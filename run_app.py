import sys
import os
import streamlit.web.cli as stcli

def main():
    # 當打包成 EXE 時，sys.executable 指向 exe 檔案
    # 我們假設 'app' 資料夾位於 exe 同一層目錄
    if getattr(sys, 'frozen', False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    
    # 指向 app/src/app.py
    script_path = os.path.join(base_dir, 'app', 'src', 'app.py')
    
    if not os.path.exists(script_path):
        print(f"[ERROR] Cannot find script at: {script_path}")
        print("Please ensure 'app' folder is in the same directory.")
        input("Press Enter to exit...")
        sys.exit(1)

    # 模擬 streamlit run 命令
    sys.argv = [
        "streamlit",
        "run",
        script_path,
        "--global.developmentMode=false",
        "--server.headless=false",
    ]
    
    sys.exit(stcli.main())

if __name__ == "__main__":
    main()
