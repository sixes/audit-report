# Script to bundle Python scripts and create a Windows installer
# Prerequisites: Install PyInstaller, Inno Setup, and UPX (at D:\program files\upx-5.0.0-win64)
# Run this script in your project directory to generate the installer

import os
import subprocess
import shutil

# Configuration
PROJECT_NAME = "AuditReport"
MAIN_SCRIPT = "main.py"  # Main Python script
OUTPUT_DIR = "dist"
INSTALLER_DIR = "installer"
RESOURCES = ["__init__.py", "data_loader.py", "document_generator.py", "exceptions.py", "utils.py", "template", "gui"]
PYTHON_VERSION = "3.9"  # Adjust based on your Python version
ICON_PATH = "app.ico"  # Optional: Path to an icon file
UPX_DIR = "D:\\program files\\upx-5.0.0-win64"  # UPX installation directory

def ensure_tools_installed():
    """Check if required tools are installed."""
    try:
        subprocess.run(["pyinstaller", "--version"], check=True, capture_output=True)
    except subprocess.CalledProcessError:
        print("PyInstaller not found. Installing...")
        subprocess.run(["pip", "install", "pyinstaller"], check=True)

    if not shutil.which("ISCC"):
        raise EnvironmentError("Inno Setup not found. Please install it from http://www.jrsoftware.org/isinfo.php")

    if not os.path.exists(os.path.join(UPX_DIR, "upx.exe")):
        print(f"Warning: UPX not found at {UPX_DIR}. Proceeding without compression.")

def create_requirements():
    """Generate requirements.txt if it doesn't exist."""
    if not os.path.exists("requirements.txt"):
        print("Generating requirements.txt...")
        subprocess.run(["pip", "freeze"], stdout=open("requirements.txt", "w"), check=True)

def clean_build_artifacts():
    """Clean PyInstaller build and dist directories."""
    print("Cleaning previous build artifacts...")
    for dir_name in [OUTPUT_DIR, os.path.join(OUTPUT_DIR, "build")]:
        if os.path.exists(dir_name):
            shutil.rmtree(dir_name, ignore_errors=True)

def bundle_with_pyinstaller():
    """Bundle the Python script and dependencies using PyInstaller."""
    print("Bundling application with PyInstaller...")

    if not os.path.exists(MAIN_SCRIPT):
        raise FileNotFoundError(f"Main script '{MAIN_SCRIPT}' not found in the project directory.")

    template_path = "template/temp_not_first_wocp.docx"
    if not os.path.exists(template_path):
        raise FileNotFoundError(
            f"Template file '{template_path}' not found. Ensure 'template' folder contains 'temp_not_first_wocp.docx'.")

    # Clean build artifacts before running
    clean_build_artifacts()

    pyinstaller_cmd = [
        "pyinstaller",
        "--name", PROJECT_NAME,
        "--noconsole",
        "--distpath", OUTPUT_DIR,
        "--workpath", os.path.join(OUTPUT_DIR, "build"),
        "--onefile",  # Single executable (comment and uncomment below for --onedir)
        # "--onedir",  # Uncomment for faster startup (requires Inno Setup to handle directory)
        "--strip",  # Strip debug symbols
        # Core dependencies
        "--hidden-import", "tkinter",  # Includes ttk, filedialog, messagebox
        "--hidden-import", "tkinter.ttk",
        "--hidden-import", "tkinter.filedialog",
        "--hidden-import", "tkinter.messagebox",
        "--hidden-import", "pandas",
        "--hidden-import", "numpy",
        "--hidden-import", "docxtpl",
        "--hidden-import", "python_docx",
        "--hidden-import", "lxml",
        # Custom modules (verify if needed)
        "--hidden-import", "document_generator",
        "--hidden-import", "data_loader",
        "--hidden-import", "exceptions",
        "--hidden-import", "utils",
        # TCL/TK for tkinter
        "--add-data", "C:\\Program Files\\Python39\\tcl\\tcl8.6;tcl\\tcl8.6",
        "--add-data", "C:\\Program Files\\Python39\\tcl\\tk8.6;tcl\\tk8.6"
    ]

    # Add UPX compression with exclusions for problematic DLLs
    if os.path.exists(UPX_DIR):
        pyinstaller_cmd.extend([
            "--upx-dir", UPX_DIR,
            "--upx-exclude", "pandas",  # Avoid compressing pandas DLLs
            "--upx-exclude", "numpy",   # Avoid compressing numpy DLLs
            "--upx-exclude", "lxml"     # Avoid compressing lxml DLLs
        ])

    if os.path.exists(ICON_PATH):
        pyinstaller_cmd.extend(["--icon", ICON_PATH])

    # Add resources
    for resource in RESOURCES:
        if os.path.exists(resource):
            separator = ";" if os.name == "nt" else ":"
            if resource.endswith(".py"):
                pyc_file = resource.replace(".py", ".pyc")
                if os.path.exists(pyc_file):
                    pyinstaller_cmd.extend(["--add-data", f"{pyc_file}{separator}."])
                else:
                    pyinstaller_cmd.extend(["--add-data", f"{resource}{separator}."])
            else:
                pyinstaller_cmd.extend(["--add-data", f"{resource}{separator}{resource}"])
        else:
            print(f"Warning: Resource '{resource}' not found and will be skipped.")

    pyinstaller_cmd.append(MAIN_SCRIPT)

    try:
        result = subprocess.run(
            pyinstaller_cmd,
            check=True,
            capture_output=True,
            text=True
        )
        print("PyInstaller output:", result.stdout)
    except subprocess.CalledProcessError as e:
        print("PyInstaller failed with error:", e.stderr)
        raise

    exe_path = os.path.join(OUTPUT_DIR, f"{PROJECT_NAME}\\{PROJECT_NAME}.exe")
    if not os.path.exists(exe_path):
        raise FileNotFoundError(f"PyInstaller did not create the expected executable at '{exe_path}'.")
    print(f"Executable created successfully at '{exe_path}'.")

def create_inno_setup_script():
    """Generate Inno Setup script (.iss file)."""
    print("Creating Inno Setup script...")

    iss_content = """
; Inno Setup script for {project_name}
#define MyAppName "{project_name}"
#define MyAppVersion "1.0"
#define MyAppPublisher "YourCompany"
#define MyAppExeName "{project_name}.exe"

[Setup]
AppName={{#MyAppName}}
AppVersion={{#MyAppVersion}}
AppPublisher={{#MyAppPublisher}}
DefaultDirName={{pf}}\\{{#MyAppName}}
DefaultGroupName={{#MyAppName}}
OutputDir={installer_dir}
OutputBaseFilename={project_name}_Setup
Compression=lzma
SolidCompression=yes
{icon_line}

[Files]
Source: "{exe_dir}\\*"; DestDir: "{{app}}"; Flags: ignoreversion recursesubdirs
{resource_lines}

[Icons]
Name: "{{group}}\\{{#MyAppName}}"; Filename: "{{app}}\\{{#MyAppExeName}}"
Name: "{{group}}\\Uninstall {{#MyAppName}}"; Filename: "{{uninstallexe}}"
Name: "{{userdesktop}}\\{{#MyAppName}}"; Filename: "{{app}}\\{{#MyAppExeName}}"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "{{cm:CreateDesktopIcon}}"; GroupDescription: "{{cm:AdditionalIcons}}"

[Run]
Filename: "{{app}}\\{{#MyAppExeName}}"; Description: "{{cm:LaunchProgram,{{#MyAppName}}}}"; Flags: nowait postinstall skipifsilent
""".format(
        project_name=PROJECT_NAME,
        installer_dir=INSTALLER_DIR,
        exe_dir=os.path.join(OUTPUT_DIR, PROJECT_NAME).replace(os.sep, "\\"),
        icon_line=f"SetupIconFile={ICON_PATH}" if os.path.exists(ICON_PATH) else "",
        resource_lines=''.join([
            f'Source: "{resource}\\*"; DestDir: "{{app}}\\{resource}"; Flags: ignoreversion recursesubdirs\n'
            if os.path.isdir(resource) else
            f'Source: "{resource}"; DestDir: "{{app}}"; Flags: ignoreversion\n'
            for resource in RESOURCES if os.path.exists(resource)
        ])
    )

    os.makedirs(INSTALLER_DIR, exist_ok=True)
    iss_file = os.path.join(INSTALLER_DIR, f"{PROJECT_NAME}.iss")
    with open(iss_file, "w") as f:
        f.write(iss_content.strip())
    return iss_file

def compile_installer(iss_file):
    """Compile the Inno Setup script to create the installer."""
    print("Compiling installer...")
    subprocess.run(["ISCC", iss_file], check=True)

def main():
    """Main function to create the standalone installer."""
    try:
        ensure_tools_installed()
        create_requirements()
        bundle_with_pyinstaller()
        iss_file = create_inno_setup_script()
        compile_installer(iss_file)
        print(f"Installer created successfully in {INSTALLER_DIR}")
    except Exception as e:
        print(f"Error: {str(e)}")

if __name__ == "__main__":
    main()