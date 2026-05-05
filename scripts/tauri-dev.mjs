import { spawn, spawnSync } from "node:child_process";
import { existsSync } from "node:fs";
import { join } from "node:path";

const projectRoot = process.cwd();

const candidates = [
  process.env.PELLA_PYTHON,
  process.env.CONDA_PREFIX ? join(process.env.CONDA_PREFIX, process.platform === "win32" ? "python.exe" : "bin/python") : "",
  join(projectRoot, ".venv", process.platform === "win32" ? "Scripts/python.exe" : "bin/python"),
  "/Users/mark/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3",
  process.platform === "win32" ? "py -3" : "python3",
  "python",
].filter(Boolean);

const python = candidates.find((candidate) => canImportOpenpyxl(candidate));

if (!python) {
  console.error("Could not find a Python installation with openpyxl available.");
  console.error("");
  console.error("Install the Python dependency with one of these:");
  console.error("  python3 -m pip install openpyxl");
  console.error("  python -m pip install openpyxl");
  console.error("");
  console.error("Or set PELLA_PYTHON to the Python executable that has openpyxl installed.");
  process.exit(1);
}

console.log(`Using Python generator runtime: ${python}`);

const child = spawn("npx", ["tauri", "dev"], {
  cwd: projectRoot,
  env: {
    ...process.env,
    PELLA_PYTHON: python,
    PYTHONPATH: join(projectRoot, "src"),
  },
  shell: process.platform === "win32",
  stdio: "inherit",
});

child.on("exit", (code, signal) => {
  if (signal) {
    process.kill(process.pid, signal);
  }
  process.exit(code ?? 1);
});

function canImportOpenpyxl(commandText) {
  const [command, ...args] = splitCommand(commandText);
  if (command.includes("/") || command.includes("\\")) {
    if (!existsSync(command)) return false;
  }
  const result = spawnSync(command, [...args, "-c", "import openpyxl"], {
    cwd: projectRoot,
    encoding: "utf8",
    shell: false,
    stdio: "ignore",
  });
  return result.status === 0;
}

function splitCommand(commandText) {
  if (commandText === "py -3") return ["py", "-3"];
  return [commandText];
}
