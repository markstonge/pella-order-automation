import { spawnSync } from "node:child_process";
import { existsSync, mkdirSync } from "node:fs";
import { join } from "node:path";

const projectRoot = process.cwd();
const python = findPython();
const targetTriple = getTargetTriple();
const extension = process.platform === "win32" ? ".exe" : "";
const sidecarName = `pella-generator-${targetTriple}${extension}`;
const binariesDir = join(projectRoot, "src-tauri", "binaries");

mkdirSync(binariesDir, { recursive: true });

console.log(`Using Python: ${python}`);
console.log(`Building generator sidecar: ${sidecarName}`);

run(python, [
  "-m",
  "PyInstaller",
  "--clean",
  "--noconfirm",
  "--onefile",
  "--name",
  sidecarName.replace(/\.exe$/, ""),
  "--distpath",
  binariesDir,
  "--workpath",
  join(projectRoot, "build", "pyinstaller"),
  "--specpath",
  join(projectRoot, "build", "pyinstaller"),
  "--paths",
  join(projectRoot, "src"),
  join(projectRoot, "scripts", "pella_generator_entry.py"),
]);

console.log(`Built ${join(binariesDir, sidecarName)}`);

function findPython() {
  const candidates = [
    process.env.PELLA_PYTHON,
    process.env.CONDA_PREFIX ? join(process.env.CONDA_PREFIX, process.platform === "win32" ? "python.exe" : "bin/python") : "",
    join(projectRoot, ".venv", process.platform === "win32" ? "Scripts/python.exe" : "bin/python"),
    "/Users/mark/opt/anaconda3/bin/python",
    "/Users/mark/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3",
    process.platform === "win32" ? "py -3" : "python3",
    "python",
  ].filter(Boolean);

  const python = candidates.find(
    (candidate) => canRun(candidate, ["-c", "import openpyxl"]) && canRun(candidate, ["-m", "PyInstaller", "--version"]),
  );
  if (!python) {
    fail("Could not find a Python runtime with both openpyxl and PyInstaller. Install dependencies or set PELLA_PYTHON.");
  }
  return python;
}

function getTargetTriple() {
  const result = spawnSync("rustc", ["--print", "host-tuple"], {
    cwd: projectRoot,
    encoding: "utf8",
  });
  if (result.status === 0 && result.stdout.trim()) {
    return result.stdout.trim();
  }

  const fallback = spawnSync("rustc", ["-vV"], {
    cwd: projectRoot,
    encoding: "utf8",
  });
  const hostLine = fallback.stdout
    .split(/\r?\n/)
    .find((line) => line.startsWith("host:"));
  const triple = hostLine?.split("host:")[1]?.trim();
  if (!triple) {
    fail("Could not determine Rust target triple. Is Rust installed?");
  }
  return triple;
}

function canRun(commandText, args) {
  const [command, ...prefixArgs] = splitCommand(commandText);
  if ((command.includes("/") || command.includes("\\")) && !existsSync(command)) return false;
  const result = spawnSync(command, [...prefixArgs, ...args], {
    cwd: projectRoot,
    encoding: "utf8",
    shell: false,
    stdio: "ignore",
  });
  return result.status === 0;
}

function run(commandText, args) {
  const [command, ...prefixArgs] = splitCommand(commandText);
  const result = spawnSync(command, [...prefixArgs, ...args], {
    cwd: projectRoot,
    env: {
      ...process.env,
      PYTHONPATH: join(projectRoot, "src"),
      PYINSTALLER_CONFIG_DIR: join(projectRoot, "build", "pyinstaller-cache"),
    },
    stdio: "inherit",
    shell: false,
  });
  if (result.status !== 0) {
    process.exit(result.status || 1);
  }
}

function splitCommand(commandText) {
  if (commandText === "py -3") return ["py", "-3"];
  return [commandText];
}

function fail(message) {
  console.error(message);
  process.exit(1);
}
