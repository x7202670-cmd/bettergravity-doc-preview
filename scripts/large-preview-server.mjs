import { timingSafeEqual } from "node:crypto";
import { open, readFile, realpath, stat } from "node:fs/promises";
import { createServer } from "node:http";
import path from "node:path";
import { pipeline } from "node:stream/promises";
import { fileURLToPath } from "node:url";
import { canUseNativeWord, convertLegacyWord, docxType, legacyWordLimit, renderWordWithOriginalLayout, warmLegacyWordPreview } from "./legacy-word-preview.mjs";

const { port, token, roots, legacyWordPython } = JSON.parse(
  await readFile(new URL("./preview-config.json", import.meta.url), "utf8")
);
const allowedRoots = await Promise.all(roots.map((root) => realpath(root)));
const authority = "127.0.0.1:" + port;
const expectedAuthorization = Buffer.from("Bearer " + token);
let activeConversions = 0;
const documentTypes = new Map([
  [".pdf", "application/pdf"],
  [".doc", "application/msword"],
  [".docx", docxType]
]);

function isWithinRoot(root, filePath) {
  const relative = path.relative(root, filePath);
  return relative && relative !== ".." && !relative.startsWith(".." + path.sep) && !path.isAbsolute(relative);
}

function getRequestedPath(uri) {
  const url = new URL(uri);
  if (url.protocol !== "file:" || url.hostname || url.search || url.hash) return null;
  const filePath = fileURLToPath(url);
  if (process.platform === "win32" && filePath.slice(2).includes(":")) return null;
  return filePath;
}

async function sendWordPreview(response, readSource, render) {
  if (activeConversions >= 2) {
    response.writeHead(429).end();
    return;
  }

  activeConversions++;
  try {
    const source = await readSource();
    if (response.destroyed) return;
    const { type, bytes } = await render(source);
    if (!response.destroyed) {
      response.writeHead(200, { "Content-Type": type, "Content-Length": bytes.length }).end(bytes);
    }
  } catch (error) {
    if (response.destroyed) return;
    const status = error.code === 2 ? 422 : error.status || 500;
    response.writeHead(status).end();
  } finally {
    activeConversions--;
  }
}

async function readWordBytes(chunksSource) {
  const chunks = [];
  let size = 0;
  for await (const chunk of chunksSource) {
    size += chunk.length;
    if (size > legacyWordLimit) throw Object.assign(new Error("DOC too large"), { status: 413 });
    chunks.push(chunk);
  }
  return Buffer.concat(chunks, size);
}

const server = createServer(async (request, response) => {
  const origin = request.headers.origin;
  if (request.headers.host !== authority || !/^https:\/\/127\.0\.0\.1:[1-9]\d{0,4}$/.test(origin || "")) {
    response.writeHead(403).end();
    return;
  }

  response.setHeader("Access-Control-Allow-Origin", origin);
  response.setHeader("Vary", "Origin");
  response.setHeader("Cache-Control", "no-store");
  response.setHeader("X-Content-Type-Options", "nosniff");

  if (!request.url?.startsWith("/")) {
    response.writeHead(400).end();
    return;
  }
  const url = new URL("http://" + authority + request.url);
  const upload = url.pathname === "/legacy-word";
  if (!upload && url.pathname !== "/document") {
    response.writeHead(404).end();
    return;
  }
  if (request.method === "OPTIONS") {
    response.writeHead(204, {
      "Access-Control-Allow-Methods": upload ? "POST" : "GET",
      "Access-Control-Allow-Headers": upload ? "Authorization, Content-Type" : "Authorization"
    }).end();
    return;
  }
  if (request.method !== (upload ? "POST" : "GET")) {
    response.writeHead(405).end();
    return;
  }

  const authorization = Buffer.from(request.headers.authorization || "");
  if (authorization.length !== expectedAuthorization.length || !timingSafeEqual(authorization, expectedAuthorization)) {
    response.writeHead(401).end();
    return;
  }

  if (upload) {
    if (!legacyWordPython) {
      response.writeHead(501).end();
      return;
    }
    if (request.headers["content-type"] !== "application/octet-stream") {
      response.writeHead(415).end();
      return;
    }
    if (Number(request.headers["content-length"]) > legacyWordLimit) {
      response.writeHead(413).end();
      return;
    }
    await sendWordPreview(response, () => readWordBytes(request.iterator({ destroyOnReturn: false })),
      (source) => convertLegacyWord(source, legacyWordPython, false));
    return;
  }

  let requestedPath;
  try {
    requestedPath = getRequestedPath(url.searchParams.get("uri"));
  } catch {
    response.writeHead(400).end();
    return;
  }
  const type = requestedPath && documentTypes.get(path.extname(requestedPath).toLowerCase());
  if (!type || !allowedRoots.some((root) => isWithinRoot(root, requestedPath))) {
    response.writeHead(403).end();
    return;
  }

  try {
    const actualPath = await realpath(requestedPath);
    if (!allowedRoots.some((root) => isWithinRoot(root, actualPath)) ||
        documentTypes.get(path.extname(actualPath).toLowerCase()) !== type) {
      response.writeHead(403).end();
      return;
    }
    const info = await stat(actualPath);
    if (!info.isFile()) {
      response.writeHead(404).end();
      return;
    }
    const file = await open(actualPath, "r");
    try {
      const opened = await file.stat();
      if (!opened.isFile() || opened.dev !== info.dev || opened.ino !== info.ino) {
        response.writeHead(404).end();
        return;
      }
      const legacyDocument = type === "application/msword";
      const originalDocxLayout = type === docxType && url.searchParams.get("layout") === "word";
      if (legacyDocument || originalDocxLayout) {
        if (legacyDocument && !legacyWordPython) {
          response.writeHead(501).end();
          return;
        }
        if (opened.size > legacyWordLimit) {
          response.writeHead(413).end();
          return;
        }
        const useNative = await canUseNativeWord(actualPath);
        if (originalDocxLayout && !useNative) {
          response.writeHead(422).end();
          return;
        }
        await sendWordPreview(response, () => readWordBytes(file.createReadStream({ autoClose: false })),
          originalDocxLayout ? (source) => renderWordWithOriginalLayout(source, ".docx")
            : (source) => convertLegacyWord(source, legacyWordPython, useNative));
        return;
      }
      response.writeHead(200, { "Content-Type": type, "Content-Length": opened.size });
      await pipeline(file.createReadStream({ autoClose: false }), response);
    } finally {
      await file.close();
    }
  } catch (err) {
    if (!response.headersSent) response.writeHead(err.code === "ENOENT" ? 404 : 500).end();
    else response.destroy();
  }
});

server.listen(port, "127.0.0.1", () => {
  console.log("Doc Preview 已启动。");
  if (legacyWordPython) warmLegacyWordPreview();
});
