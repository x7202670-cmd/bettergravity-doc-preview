import { execFile, spawn } from "node:child_process";
import { createHash } from "node:crypto";
import { mkdtemp, open, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { createInterface } from "node:readline";
import { fileURLToPath } from "node:url";

export const officePreviewLimit = 128 * 1024 * 1024;
export const docxType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const workerScripts = { Word: "./export-legacy-word.vbs", PowerPoint: "./export-powerpoint.vbs" };
const converterPath = fileURLToPath(new URL("./convert-legacy-word.py", import.meta.url));
const oleSignature = Buffer.from("d0cf11e0a1b11ae1", "hex");
const zipSignature = Buffer.from([0x50, 0x4b, 0x03, 0x04]);
const cacheLimit = 48 * 1024 * 1024;

const workers = new Map();
const retryAfter = new Map();
const nativeQueues = new Map();
const cache = new Map();
const pending = new Map();
let cacheBytes = 0;

function getWorker(application) {
  if (workers.has(application)) return workers.get(application);
  if (Date.now() < retryAfter.get(application)) throw new Error(application + " automation unavailable");

  const launcher = path.join(process.env.SystemRoot || "C:\\Windows", "System32", "cscript.exe");
  const nativeWorkerPath = fileURLToPath(new URL(workerScripts[application], import.meta.url));
  const child = spawn(launcher, ["//NoLogo", "//B", "//U", nativeWorkerPath], {
    windowsHide: true,
    stdio: ["pipe", "pipe", "ignore"]
  });
  const lines = createInterface({ input: child.stdout.setEncoding("utf16le"), crlfDelay: Infinity });
  let reply;
  let idle;

  function stop(error) {
    if (workers.get(application) !== instance) return;
    workers.delete(application);
    clearTimeout(idle);
    reply?.reject(error);
    child.stdin.end();
    const forceStop = setTimeout(() => child.kill(), 3_000);
    forceStop.unref();
    child.once("close", () => clearTimeout(forceStop));
  }

  function waitForReply(timeout) {
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => stop(new Error(application + " export timed out")), timeout);
      reply = {
        resolve(value) { clearTimeout(timer); reply = undefined; resolve(value); },
        reject(error) { clearTimeout(timer); reply = undefined; reject(error); }
      };
    });
  }

  function armIdleTimer() {
    idle = setTimeout(() => stop(new Error(application + " export worker idle")), 15 * 60_000);
    idle.unref();
  }

  const instance = {
    ready: waitForReply(15_000).then((message) => {
      if (message !== "ready") throw new Error("Unexpected " + application + " worker response");
      armIdleTimer();
    }).catch((error) => {
      retryAfter.set(application, Date.now() + 60_000);
      stop(error);
      throw error;
    }),
    async export(sourcePath, outputPath) {
      await instance.ready;
      clearTimeout(idle);
      const result = waitForReply(30_000);
      child.stdin.write(Buffer.from(`${sourcePath}\r\n${outputPath}\r\n`, "utf16le"));
      if (await result !== "ok") {
        armIdleTimer();
        throw new Error(application + " could not export this document");
      }
      armIdleTimer();
    }
  };

  lines.on("line", (line) => reply?.resolve(line.replace(/^\uFEFF/, "")));
  child.on("error", stop);
  child.on("exit", () => stop(new Error(application + " export worker exited")));
  child.stdin.on("error", stop);
  workers.set(application, instance);
  return instance;
}

export function warmLegacyWordPreview() {
  if (process.platform === "win32") void getWorker("Word").ready.catch(() => {});
}

export async function canUseNativeOffice(filePath) {
  if (process.platform !== "win32") return false;
  try {
    const marker = await open(filePath + ":Zone.Identifier", "r");
    await marker.close();
  } catch (error) {
    return error.code === "ENOENT";
  }
  return false;
}

function exportNativePdf(source, extension, application) {
  const previous = nativeQueues.get(application) ?? Promise.resolve();
  const job = previous.then(async () => {
    const renderer = getWorker(application);
    await renderer.ready;
    const directory = await mkdtemp(path.join(tmpdir(), "bg-office-preview-"));
    try {
      const input = path.join(directory, "document" + extension);
      const output = path.join(directory, "preview.pdf");
      await writeFile(input, source);
      await renderer.export(input, output);
      const preview = await readFile(output);
      if (preview.length > officePreviewLimit || preview.subarray(0, 5).toString() !== "%PDF-") {
        throw new Error("Invalid Office PDF output");
      }
      return preview;
    } finally {
      await rm(directory, { recursive: true, force: true, maxRetries: 3, retryDelay: 100 }).catch(() => {});
    }
  });
  nativeQueues.set(application, job.catch(() => {}));
  return job;
}

function convertWithDocVortex(source, python, args = []) {
  return new Promise((resolve, reject) => {
    const child = execFile(python, ["-I", converterPath, ...args], {
      encoding: "buffer",
      maxBuffer: officePreviewLimit,
      timeout: 45_000,
      windowsHide: true
    }, (error, stdout) => error ? reject(error) : resolve(stdout));
    child.stdin.on("error", () => {}); // The decoder may exit before receiving the full input.
    child.stdin.end(source);
  });
}

function remember(key, result) {
  const size = result.bytes.length;
  if (size > cacheLimit) return;
  while (cache.size >= 12 || cacheBytes + size > cacheLimit) {
    const oldest = cache.keys().next().value;
    cacheBytes -= cache.get(oldest).bytes.length;
    cache.delete(oldest);
  }
  cache.set(key, result);
  cacheBytes += size;
}

function renderWithCache(key, render) {
  const cached = cache.get(key);
  if (cached) {
    cache.delete(key);
    cache.set(key, cached);
    return cached;
  }
  if (pending.has(key)) return pending.get(key);

  const job = render().then((result) => {
    remember(key, result);
    return result;
  }).finally(() => pending.delete(key));
  pending.set(key, job);
  return job;
}

export function renderWordWithOriginalLayout(source, extension) {
  const signature = extension === ".docx" ? zipSignature : oleSignature;
  if (!source.subarray(0, signature.length).equals(signature)) {
    throw Object.assign(new Error("Invalid Word document"), { status: 422 });
  }
  const digest = createHash("sha256").update(source).digest("hex");
  return renderWithCache("pdf:" + extension + ":" + digest, async () => ({
    type: "application/pdf", bytes: await exportNativePdf(source, extension, "Word")
  }));
}

export async function convertLegacyWord(source, python, native) {
  if (native && source.subarray(0, 8).equals(oleSignature)) {
    try {
      return await renderWordWithOriginalLayout(source, ".doc");
    } catch {
      // The legacy decoder also supports files that Word can't open automatically.
    }
  }
  const digest = createHash("sha256").update(source).digest("hex");
  return renderWithCache("docx:" + digest, async () => ({ type: docxType, bytes: await convertWithDocVortex(source, python) }));
}

export async function renderPresentation(source, extension, python, native) {
  const signature = extension === ".pptx" ? zipSignature : oleSignature;
  if (!source.subarray(0, signature.length).equals(signature)) {
    throw Object.assign(new Error("Invalid presentation"), { status: 422 });
  }
  const key = extension + ":" + createHash("sha256").update(source).digest("hex");
  if (native) {
    try {
      return await renderWithCache("powerpoint:" + key, async () => ({
        type: "application/pdf", bytes: await exportNativePdf(source, extension, "PowerPoint")
      }));
    } catch {
      // The offline decoder remains available when PowerPoint cannot export the file.
    }
  }
  if (!python) throw Object.assign(new Error("Presentation converter unavailable"), { status: 501 });
  return renderWithCache("slide-pdf:" + key, async () => ({
    type: "application/pdf", bytes: await convertWithDocVortex(source, python, [extension.slice(1)])
  }));
}
