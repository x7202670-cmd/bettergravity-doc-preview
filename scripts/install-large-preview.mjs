import { execFileSync } from "node:child_process";
import { randomBytes, randomUUID } from "node:crypto";
import fs from "node:fs";
import { createServer } from "node:net";
import os from "node:os";
import path from "node:path";

const [pluginDir, ...installerArgs] = process.argv.slice(2);
const configureLegacyWord = installerArgs[0] === "--legacy-word-python";
const requestedRoots = configureLegacyWord ? installerArgs.slice(2) : installerArgs;
const sidecarId = "doc-preview-local";
const configDir = path.join(os.homedir(), ".gemini", "config");
const sidecarDir = path.join(configDir, "sidecars", sidecarId);
const previewConfigPath = path.join(sidecarDir, "preview-config.json");
const sidecarManifestPath = path.join(sidecarDir, "sidecar.json");
const sidecarDisplayName = "Office & PDF Preview 大文件预览";
const marker = "const LOCAL_PREVIEW_CONFIG = null;";
const bundle = fs.readFileSync(new URL("../dist/index.js", import.meta.url), "utf8");

function writeAtomic(filePath, contents) {
  const temporaryPath = filePath + "." + randomUUID() + ".tmp";
  fs.writeFileSync(temporaryPath, contents, { mode: 0o600 });
  try {
    fs.renameSync(temporaryPath, filePath);
  } finally {
    fs.rmSync(temporaryPath, { force: true });
  }
}

function writeJson(filePath, data) {
  writeAtomic(filePath, JSON.stringify(data, null, 2) + "\n");
}

async function selectPort() {
  const socket = createServer();
  return new Promise((resolve, reject) => {
    socket.once("error", reject);
    socket.listen(0, "127.0.0.1", () => {
      const { port } = socket.address();
      socket.close((error) => error ? reject(error) : resolve(port));
    });
  });
}

function readPreviewConfig() {
  if (!fs.existsSync(previewConfigPath)) return null;
  const config = JSON.parse(fs.readFileSync(previewConfigPath, "utf8"));
  if (!Number.isInteger(config.port) || config.port < 1 || config.port > 65535 ||
      !/^[a-f0-9]{64}$/.test(config.token)) {
    throw new Error("本地预览配置有误，请重新指定工作区目录。");
  }
  return config;
}

if (!pluginDir || !bundle.includes(marker)) throw new Error("未找到 Office & PDF Preview 构建产物。");
let config = readPreviewConfig();

if (requestedRoots.length || configureLegacyWord) {
  const roots = requestedRoots.length ? [...new Set(requestedRoots.map((requestedRoot) => {
    const root = fs.realpathSync(requestedRoot);
    if (!fs.statSync(root).isDirectory()) throw new Error("预览路径不是目录：" + requestedRoot);
    return root;
  }))] : config?.roots ?? [];

  let legacyWordPython = config?.legacyWordPython;
  if (configureLegacyWord) {
    if (!installerArgs[1]) throw new Error("请指定包含 DocVortex 的 Python 路径。");
    legacyWordPython = path.resolve(installerArgs[1]);
    execFileSync(legacyWordPython, ["-I", "-c", "from docvortex import parse, render_artifact"], {
      stdio: "inherit",
      timeout: 20_000,
      windowsHide: true
    });
  }

  const settingsPath = path.join(configDir, "config.json");
  const settings = fs.existsSync(settingsPath)
    ? JSON.parse(fs.readFileSync(settingsPath, "utf8").replace(/^\uFEFF/, ""))
    : {};
  settings.sidecars ??= {};
  settings.sidecars[sidecarId] = { ...settings.sidecars[sidecarId], enabled: true };

  config = {
    port: config?.port ?? await selectPort(),
    token: config?.token ?? randomBytes(32).toString("hex"),
    roots,
    legacyWordPython
  };
  fs.mkdirSync(sidecarDir, { recursive: true });
  writeJson(previewConfigPath, config);
  writeJson(sidecarManifestPath, {
    display_name: sidecarDisplayName,
    command: process.execPath,
    args: [path.join(sidecarDir, "large-preview-server.mjs")],
    restart_policy: "on-failure"
  });
  writeJson(settingsPath, settings);
}

if (config) {
  if (fs.existsSync(sidecarManifestPath)) {
    const manifest = JSON.parse(fs.readFileSync(sidecarManifestPath, "utf8"));
    if (manifest.display_name !== sidecarDisplayName) {
      writeJson(sidecarManifestPath, { ...manifest, display_name: sidecarDisplayName });
    }
  }
  for (const name of ["convert-legacy-word.py", "normalize-docx-fonts.py", "export-legacy-word.vbs", "export-powerpoint.vbs", "legacy-word-preview.mjs", "large-preview-server.mjs"]) {
    writeAtomic(path.join(sidecarDir, name), fs.readFileSync(new URL("./" + name, import.meta.url)));
  }
}

const connection = config && { port: config.port, token: config.token };
const installedBundle = bundle.replace(marker, () => "const LOCAL_PREVIEW_CONFIG = " + JSON.stringify(connection) + ";");
writeAtomic(path.join(pluginDir, "index.js"), installedBundle);
