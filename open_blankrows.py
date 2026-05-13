from pathlib import Path
import subprocess
import sys
import urllib.request
import webbrowser
from datetime import datetime

BASE_URL = "https://script.google.com/macros/s/AKfycbx8nIINe088BwUOqmUXEoS4uqb3iq55YQMj7XfsXXB6qR9Ffm28njn80Oanlv4Hax4_/exec"
LIMIT = 200

url = f"{BASE_URL}?action=blankRows&limit={LIMIT}"
out_path = Path(__file__).with_name("blankRows_latest.txt")

def copy_to_clipboard(text: str):
    if sys.platform == "darwin":
        subprocess.run("pbcopy", input=text, text=True, check=True)
    elif sys.platform.startswith("win"):
        subprocess.run("clip", input=text, text=True, check=True)
    else:
        subprocess.run(["xclip", "-selection", "clipboard"], input=text, text=True, check=True)

try:
    with urllib.request.urlopen(url, timeout=30) as response:
        text = response.read().decode("utf-8")

    header = f"# fetched_at={datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n# url={url}\n\n"
    out_path.write_text(header + text, encoding="utf-8")

    try:
        copy_to_clipboard(text)
        print("已抓到 blankRows，並複製到剪貼簿。")
    except Exception:
        print("已抓到 blankRows，但剪貼簿複製失敗。")

    print(f"已存成：{out_path}")
    print("正在打開網頁...")
    webbrowser.open(url)

except Exception as e:
    print("抓取失敗：", e)
    print("改用瀏覽器打開網頁。")
    webbrowser.open(url)