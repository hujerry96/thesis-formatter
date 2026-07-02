import subprocess
import sys
from pathlib import Path

def build_exe():
    project_root = Path(__file__).parent
    main_script = project_root / "main.py"
    rules_dir = project_root / "rules"
    
    sep = ";" if sys.platform == "win32" else ":"
    
    add_data_args = []
    for f in rules_dir.glob("*.yaml"):
        add_data_args.append("--add-data")
        add_data_args.append(f"{f}{sep}rules")
    
    cmd = [
        "pyinstaller",
        "--onefile",
        "--noconfirm",
        "--windowed",
        "--name", "論文格式修正工具",
        *add_data_args,
        "--hidden-import", "tkinterdnd2",
        "--hidden-import", "core.debug_log",
        "--hidden-import", "core",
        "--hidden-import", "core.rule_engine",
        "--hidden-import", "core.analyzer",
        "--hidden-import", "core.fixer",
        "--hidden-import", "gui",
        "--hidden-import", "gui.main_gui",
        "--hidden-import", "yaml",
        "--hidden-import", "lxml",
        "--hidden-import", "docx",
        "--collect-all", "docx",
        str(main_script)
    ]
    
    print("執行 PyInstaller 打包命令...")
    print(" ".join(cmd))
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        print("錯誤:", result.stderr)
    else:
        print("打包完成！執行檔在 dist/ 資料夾")
        for f in (project_root / "dist").iterdir():
            print(f"  {f}")

if __name__ == "__main__":
    build_exe()
