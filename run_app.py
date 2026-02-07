import os
import sys
from pathlib import Path
from streamlit.web import cli as stcli

def _abs_app_path() -> str:
    # PyInstaller(onefile/onedir) 모두에서 app.py 경로 안정화
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return str((base / "app.py").resolve())

def main():
    # "개발모드" / "node dev server(3000)"로 새는 것 차단
    os.environ["STREAMLIT_GLOBAL_DEVELOPMENT_MODE"] = "false"
    os.environ["STREAMLIT_SERVER_PORT"] = "8501"
    os.environ["STREAMLIT_SERVER_HEADLESS"] = "true"

    # (선택) CORS/XSRF는 로컬 단독 실행이면 아래처럼 단순화 가능
    os.environ["STREAMLIT_SERVER_ENABLE_CORS"] = "true"
    os.environ["STREAMLIT_SERVER_ENABLE_XSRF_PROTECTION"] = "false"

    # Streamlit 실행 인자 세팅
    sys.argv = [
        "streamlit",
        "run",
        _abs_app_path(),
        "--server.port=8501",
        "--server.headless=true",
    ]

    stcli.main()

if __name__ == "__main__":
    main()
