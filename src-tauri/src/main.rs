use std::env;
use std::fs::OpenOptions;
use std::io::Write;
use std::path::{Path, PathBuf};
use std::process::Command;
use std::time::{SystemTime, UNIX_EPOCH};

use tauri_plugin_shell::ShellExt;

#[tauri::command]
async fn generate_workbook(
    app: tauri::AppHandle,
    po_path: String,
    work_order_path: String,
    output_path: String,
) -> Result<String, String> {
    log_message(&format!(
        "Generate requested: po_path=\"{}\", work_order_path=\"{}\", output_path=\"{}\"",
        po_path, work_order_path, output_path
    ));

    if should_use_python_bridge() {
        log_message("Using development Python generator bridge.");
        let result = generate_with_python_bridge(po_path, work_order_path, output_path);
        log_result(&result);
        return result;
    }

    log_message("Using bundled generator sidecar.");
    let sidecar = app
        .shell()
        .sidecar("pella-generator")
        .map_err(|error| {
            let message = format!("Could not find the bundled generator sidecar: {error}");
            log_message(&message);
            message
        })?
        .args([
            "generate",
            "--po",
            &po_path,
            "--work-order",
            &work_order_path,
            "--output",
            &output_path,
            "--summary-json",
        ]);

    let output = sidecar.output().await.map_err(|error| {
        let message = format!("Could not start the bundled generator: {error}");
        log_message(&message);
        message
    })?;

    let result = handle_generator_output(output.status.success(), output.stdout, output.stderr);
    log_result(&result);
    result
}

fn generate_with_python_bridge(
    po_path: String,
    work_order_path: String,
    output_path: String,
) -> Result<String, String> {
    let project_root = project_root();
    let mut command = generator_command(&project_root);

    command
        .arg("generate")
        .arg("--po")
        .arg(po_path)
        .arg("--work-order")
        .arg(work_order_path)
        .arg("--output")
        .arg(output_path)
        .arg("--summary-json");

    let output = command
        .output()
        .map_err(|error| format!("Could not start the workbook generator: {error}"))?;

    handle_generator_output(output.status.success(), output.stdout, output.stderr)
}

fn handle_generator_output(
    success: bool,
    stdout: Vec<u8>,
    stderr: Vec<u8>,
) -> Result<String, String> {
    if success {
        String::from_utf8(stdout)
            .map_err(|error| format!("Generator returned invalid UTF-8: {error}"))
    } else {
        let stderr = String::from_utf8_lossy(&stderr);
        let stdout = String::from_utf8_lossy(&stdout);
        let details = if stderr.trim().is_empty() {
            stdout.trim()
        } else {
            stderr.trim()
        };
        Err(if details.is_empty() {
            "Workbook generation failed without diagnostic output.".to_string()
        } else {
            details.to_string()
        })
    }
}

fn log_result(result: &Result<String, String>) {
    match result {
        Ok(_) => log_message("Generation completed successfully."),
        Err(error) => log_message(&format!("Generation failed: {error}")),
    }
}

fn log_message(message: &str) {
    if !cfg!(windows) {
        return;
    }

    let Some(app_dir) = env::current_exe()
        .ok()
        .and_then(|path| path.parent().map(Path::to_path_buf))
    else {
        return;
    };

    let log_path = app_dir.join("pella-order-automation.log");
    let Ok(mut log_file) = OpenOptions::new().create(true).append(true).open(log_path) else {
        return;
    };
    let timestamp = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|duration| duration.as_secs().to_string())
        .unwrap_or_else(|_| "unknown-time".to_string());
    let _ = writeln!(log_file, "[{timestamp}] {message}");
}

fn should_use_python_bridge() -> bool {
    env::var("PELLA_GENERATOR_BIN").is_ok()
        || env::var("PELLA_PYTHON").is_ok()
        || cfg!(debug_assertions)
}

fn generator_command(project_root: &Path) -> Command {
    if let Ok(generator_bin) = env::var("PELLA_GENERATOR_BIN") {
        return Command::new(generator_bin);
    }

    let python = env::var("PELLA_PYTHON")
        .or_else(|_| {
            env::var("CONDA_PREFIX").map(|prefix| {
                Path::new(&prefix)
                    .join(conda_python_name())
                    .to_string_lossy()
                    .to_string()
            })
        })
        .unwrap_or_else(|_| "python3".to_string());
    let mut command = Command::new(python);
    command.arg("-m").arg("pella_order_automation.cli");

    let python_path = project_root.join("src");
    command.env("PYTHONPATH", python_path);
    command
}

fn conda_python_name() -> &'static str {
    if cfg!(windows) {
        "python.exe"
    } else {
        "bin/python"
    }
}

fn project_root() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .parent()
        .map(Path::to_path_buf)
        .unwrap_or_else(|| PathBuf::from("."))
}

fn main() {
    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_shell::init())
        .invoke_handler(tauri::generate_handler![generate_workbook])
        .run(tauri::generate_context!())
        .expect("error while running Pella Order Automation");
}
