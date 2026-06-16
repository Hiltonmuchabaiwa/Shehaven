from pathlib import Path
import subprocess
import sys

folder = Path(__file__).resolve().parent
app = folder / "app.py"
subprocess.call([sys.executable, "-m", "pip", "install", "--upgrade", "-r", str(folder / "requirements.txt")])
subprocess.call([sys.executable, "-m", "streamlit", "run", str(app)])
