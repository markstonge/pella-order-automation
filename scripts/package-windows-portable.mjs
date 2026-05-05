import { spawnSync } from "node:child_process";
import { copyFileSync, existsSync, mkdirSync, rmSync, writeFileSync } from "node:fs";
import { join } from "node:path";

const projectRoot = process.cwd();
const releaseDir = join(projectRoot, "src-tauri", "target", "release");
const binariesDir = join(projectRoot, "src-tauri", "binaries");
const portableRoot = join(projectRoot, "dist", "windows-portable");
const appFolder = join(portableRoot, "Pella Order Automation");
const zipPath = join(projectRoot, "dist", "Pella Order Automation Windows Portable.zip");

if (process.platform !== "win32") {
  fail("Windows portable packaging must be run on Windows after building the Windows Tauri app.");
}

const appExe = join(releaseDir, "pella-order-automation.exe");
const releaseSidecar = join(releaseDir, "pella-generator.exe");
const sourceSidecar = join(binariesDir, "pella-generator-x86_64-pc-windows-msvc.exe");
const sidecarExe = existsSync(releaseSidecar) ? releaseSidecar : sourceSidecar;

assertFile(appExe, "Could not find the built Tauri app executable.");
assertFile(sidecarExe, "Could not find the Windows generator sidecar executable.");

rmSync(portableRoot, { recursive: true, force: true });
mkdirSync(appFolder, { recursive: true });

copyFileSync(appExe, join(appFolder, "Pella Order Automation.exe"));
copyFileSync(sidecarExe, join(appFolder, "pella-generator.exe"));

writeFileSync(
  join(appFolder, "README.txt"),
  [
    "Pella Order Automation",
    "",
    "To run the app:",
    "1. Open Pella Order Automation.exe.",
    "2. Select the purchase order and work order files.",
    "3. Generate the completed workbook.",
    "",
    "Cleanup:",
    "Delete this folder to remove the app and its app-owned logs/troubleshooting files.",
    "Generated workbooks saved outside this folder are user files and are not removed.",
    "",
  ].join("\r\n"),
);

const powershell = findPowerShell();
const archive = spawnSync(
  powershell,
  [
    "-NoProfile",
    "-ExecutionPolicy",
    "Bypass",
    "-Command",
    `Compress-Archive -LiteralPath ${quotePowerShell(appFolder)} -DestinationPath ${quotePowerShell(zipPath)} -Force`,
  ],
  {
    cwd: projectRoot,
    encoding: "utf8",
    stdio: "inherit",
  },
);

if (archive.status !== 0) {
  process.exit(archive.status || 1);
}

console.log(`Created ${zipPath}`);

function assertFile(path, message) {
  if (!existsSync(path)) {
    fail(`${message}\nMissing: ${path}`);
  }
}

function findPowerShell() {
  for (const command of ["pwsh.exe", "powershell.exe"]) {
    const result = spawnSync(command, ["-NoProfile", "-Command", "$PSVersionTable.PSVersion.ToString()"], {
      encoding: "utf8",
      stdio: "ignore",
    });
    if (result.status === 0) return command;
  }
  fail("Could not find PowerShell to create the portable zip.");
}

function quotePowerShell(value) {
  return `'${value.replace(/'/g, "''")}'`;
}

function fail(message) {
  console.error(message);
  process.exit(1);
}
