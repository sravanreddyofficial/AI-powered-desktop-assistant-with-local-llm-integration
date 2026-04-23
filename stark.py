"""
╔══════════════════════════════════════════════════════════════════════╗
║                                                                      ║
║   ███████╗████████╗ █████╗ ██████╗ ██╗  ██╗                         ║
║   ██╔════╝╚══██╔══╝██╔══██╗██╔══██╗██║ ██╔╝                         ║
║   ███████╗   ██║   ███████║██████╔╝█████╔╝                          ║
║   ╚════██║   ██║   ██╔══██║██╔══██╗██╔═██╗                          ║
║   ███████║   ██║   ██║  ██║██║  ██║██║  ██╗                         ║
║   ╚══════╝   ╚═╝   ╚═╝  ╚═╝╚═╝  ╚═╝╚═╝  ╚═╝                        ║
║                                                                      ║
║   Your Personal AI Desktop Assistant — v2.0                          ║
║   Powered by LM Studio + Meta Llama 3.1 8B Instruct                  ║
║                                                                      ║
╚══════════════════════════════════════════════════════════════════════╝
"""

# ── Standard library ─────────────────────────────────────────────────
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import threading
import requests
import json
import os
import sys
import math
import shutil
import subprocess
import webbrowser
import time
import re
import fnmatch
import importlib.util
from datetime import datetime
from pathlib import Path
from urllib.parse import quote as urlquote

# ── Optional libraries ───────────────────────────────────────────────
try:
    import psutil
    HAS_PSUTIL = True
except ImportError:
    HAS_PSUTIL = False

try:
    import pyperclip
    HAS_CLIP = True
except ImportError:
    HAS_CLIP = False

try:
    import pyttsx3
    HAS_TTS = True
except ImportError:
    HAS_TTS = False

try:
    import send2trash
    HAS_TRASH = True
except ImportError:
    HAS_TRASH = False


# ════════════════════════════════════════════════════════════════════
#  CONFIGURATION
# ════════════════════════════════════════════════════════════════════
LM_STUDIO_URL = "http://localhost:1234"
MODEL_NAME    = "meta-llama-3.1-8b-instruct"
TEMPERATURE   = 0.7
MAX_TOKENS    = 1024

HOME      = Path(os.environ.get("USERPROFILE", Path.home()))
DESKTOP   = HOME / "Desktop"
DOCUMENTS = HOME / "Documents"
DOWNLOADS = HOME / "Downloads"
MUSIC_DIR = HOME / "Music"
VIDEOS_DIR= HOME / "Videos"
PICTURES  = HOME / "Pictures"

CHROME_PATHS = [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    str(HOME / r"AppData\Local\Google\Chrome\Application\chrome.exe"),
]
SPOTIFY_PATHS = [
    str(HOME / r"AppData\Roaming\Spotify\Spotify.exe"),          # standard install
    str(HOME / r"AppData\Local\Microsoft\WindowsApps\Spotify.exe"), # MS Store
    r"C:\Program Files\Spotify\Spotify.exe",                     # system-wide
    r"C:\Program Files (x86)\Spotify\Spotify.exe",
    str(HOME / r"AppData\Local\Spotify\Spotify.exe"),
]

def find_spotify() -> str | None:
    """Find Spotify exe, including dynamic WindowsApps subfolder."""
    exe = find_exe(SPOTIFY_PATHS)
    if exe:
        return exe
    # Microsoft Store apps sometimes live in versioned subfolders
    wa = HOME / r"AppData\Local\Microsoft\WindowsApps"
    if wa.exists():
        for p in wa.rglob("Spotify.exe"):
            return str(p)
    # Also try registry-style lookup via where command
    try:
        r = subprocess.run(["where", "Spotify.exe"], capture_output=True, text=True, timeout=3)
        if r.returncode == 0:
            line = r.stdout.strip().splitlines()[0].strip()
            if line: return line
    except Exception:
        pass
    return None

def find_youtube_app() -> str | None:
    """Check if YouTube app is installed (very rare on Windows, usually browser)."""
    # YouTube does not have a native Windows desktop app — always use browser
    return None
VLC_PATHS = [
    r"C:\Program Files\VideoLAN\VLC\vlc.exe",
    r"C:\Program Files (x86)\VideoLAN\VLC\vlc.exe",
]
WMP_PATHS = [
    r"C:\Program Files\Windows Media Player\wmplayer.exe",
    r"C:\Program Files (x86)\Windows Media Player\wmplayer.exe",
]

AUDIO_EXTS = {".mp3",".wav",".flac",".aac",".ogg",".wma",".m4a",".opus",".aiff"}
VIDEO_EXTS = {".mp4",".mkv",".avi",".mov",".wmv",".flv",".webm",".m4v",".ts",".mpeg"}
MEDIA_EXTS = AUDIO_EXTS | VIDEO_EXTS

# ── Colour palette: Iron Man / Arc Reactor ──────────────────────────
C_BG      = "#050810"
C_PANEL   = "#090e1a"
C_CARD    = "#0d1526"
C_BORDER  = "#1a2a4a"
C_ARC     = "#00d4ff"
C_ARC2    = "#0077aa"
C_GOLD    = "#f5a623"
C_AMBER   = "#ff8c00"
C_RED     = "#ff3b5c"
C_GREEN   = "#00ff9d"
C_WHITE   = "#e8f4ff"
C_MUTED   = "#3a5070"
C_PURPLE  = "#a29bfe"
C_PINK    = "#fd79a8"
C_TEAL    = "#00cec9"
C_ORANGE  = "#ff9f43"
C_YELLOW  = "#ffd166"

FONT_TITLE  = ("Courier New", 11, "bold")
FONT_MONO   = ("Courier New", 10)
FONT_MONO_S = ("Courier New", 9)
FONT_UI     = ("Segoe UI", 10)
FONT_UI_B   = ("Segoe UI", 10, "bold")
FONT_SMALL  = ("Segoe UI", 8)
FONT_BTN    = ("Segoe UI", 9, "bold")


# ════════════════════════════════════════════════════════════════════
#  UTILITIES
# ════════════════════════════════════════════════════════════════════
def find_exe(paths: list) -> str | None:
    for p in paths:
        if Path(p).exists():
            return p
    return None

def fmt_size(b: int) -> str:
    for u in ["B","KB","MB","GB","TB"]:
        if b < 1024: return f"{b:.1f} {u}"
        b /= 1024
    return f"{b:.1f} PB"

def fmt_time(ts: float) -> str:
    return datetime.fromtimestamp(ts).strftime("%Y-%m-%d  %H:%M:%S")

def clean_for_speech(text: str) -> str:
    text = re.sub(r'[^\x00-\x7F]+', ' ', text)
    text = re.sub(r'[*_`#\-─═╔╗╚╝║|]+', '', text)
    text = re.sub(r'[A-Za-z]:\\[^\s]+\\([^\s]+)', r'\1', text)
    text = re.sub(r'^(STARK|AI)\s*[>❯›]+\s*', '', text, flags=re.MULTILINE)
    text = re.sub(r'\s{3,}', ' ', text)
    text = re.sub(r'\n{2,}', '. ', text)
    return text.strip()

def open_in_chrome(url: str):
    chrome = find_exe(CHROME_PATHS)
    if chrome:
        try:
            subprocess.Popen([chrome, url])
            return True
        except Exception:
            pass
    webbrowser.open(url)
    return True

def get_system_info() -> str:
    if not HAS_PSUTIL:
        return "psutil not installed — run: pip install psutil"
    cpu = psutil.cpu_percent(interval=0.3)
    ram = psutil.virtual_memory()
    lines = [
        f"CPU      {cpu:.1f}%",
        f"RAM      {ram.used/1e9:.1f} / {ram.total/1e9:.1f} GB  ({ram.percent:.0f}%)",
    ]
    try:
        disk = psutil.disk_usage("C:\\")
        lines.append(f"Disk C:  {disk.used/1e9:.0f} / {disk.total/1e9:.0f} GB  ({disk.percent:.0f}%)")
    except Exception:
        pass
    try:
        b = psutil.sensors_battery()
        if b:
            lines.append(f"Battery  {b.percent:.0f}%  ({'charging ⚡' if b.power_plugged else 'on battery 🔋'})")
    except Exception:
        pass
    lines.append(f"Time     {datetime.now().strftime('%I:%M %p  —  %A, %B %d %Y')}")
    return "\n".join(lines)


# ════════════════════════════════════════════════════════════════════
#  VOICE ENGINE
# ════════════════════════════════════════════════════════════════════
class VoiceEngine:
    def __init__(self):
        self.enabled  = True
        self.rate     = 175
        self.volume   = 0.92
        self._engine  = None
        self._lock    = threading.Lock()
        self._init()

    def _init(self):
        if not HAS_TTS:
            return
        try:
            self._engine = pyttsx3.init()
            self._engine.setProperty("rate",   self.rate)
            self._engine.setProperty("volume", self.volume)
            voices = self._engine.getProperty("voices") or []
            self._voices = voices
            if voices:
                self._engine.setProperty("voice", voices[0].id)
        except Exception as e:
            print(f"[Voice] Init failed: {e}")
            self._engine = None

    def speak(self, text: str, force: bool = False):
        if not self._engine: return
        if not self.enabled and not force: return
        clean = clean_for_speech(text)
        words = clean.split()
        if len(words) > 80:
            clean = " ".join(words[:80]) + "..."
        if len(clean) < 2: return
        self.stop()
        def _run():
            with self._lock:
                try:
                    self._engine.say(clean)
                    self._engine.runAndWait()
                except Exception:
                    pass
        threading.Thread(target=_run, daemon=True).start()

    def stop(self):
        if self._engine:
            try: self._engine.stop()
            except: pass

    def get_voice_names(self) -> list:
        if not self._engine: return ["Default"]
        voices = getattr(self, "_voices", [])
        return [getattr(v, "name", str(v)).replace("Microsoft ", "").replace(" Desktop","") for v in voices] or ["Default"]

    def set_voice_idx(self, idx: int):
        voices = getattr(self, "_voices", [])
        if voices and self._engine:
            try: self._engine.setProperty("voice", voices[idx % len(voices)].id)
            except: pass

    def set_rate(self, r: int):
        self.rate = r
        if self._engine:
            try: self._engine.setProperty("rate", r)
            except: pass

    def set_volume(self, v: float):
        self.volume = v
        if self._engine:
            try: self._engine.setProperty("volume", v)
            except: pass

    def toggle(self) -> bool:
        self.enabled = not self.enabled
        if not self.enabled: self.stop()
        return self.enabled


# ════════════════════════════════════════════════════════════════════
#  AI CLIENT  (LM Studio streaming)
# ════════════════════════════════════════════════════════════════════
class AIClient:
    SYSTEM_PROMPT = (
        "You are Stark — a razor-sharp, Siri/Jarvis-inspired AI desktop assistant. "
        "Your personality: confident, witty, warm, and extremely capable — like Jarvis from Iron Man. "
        "Rules: "
        "1. Always respond in plain natural English — no markdown, no asterisks, no bullet symbols. "
        "2. Be concise. Get to the point within 2-3 sentences unless detail is specifically needed. "
        "3. Sound human and conversational — not robotic or overly formal. "
        "4. For factual questions give accurate, direct answers. "
        "5. For creative tasks (writing, brainstorming, coding) be excellent and specific. "
        "6. File operations, app launching, media playback, system info, web search, "
        "time/date, math, and voice control are all handled by local plugins automatically — "
        "so you only see messages that truly need your intelligence. "
        "7. If you do not know something, say so honestly and offer to help another way. "
        "8. Keep responses under 200 words unless the user explicitly asks for more. "
        "9. Never start with 'Certainly!', 'Of course!', 'Sure!', or similar filler openers. "
        "10. You are the user's personal assistant — be personable, helpful, and on-point."
    )

    def __init__(self):
        self.history: list = []

    def is_online(self) -> bool:
        try:
            r = requests.get(f"{LM_STUDIO_URL}/v1/models", timeout=3)
            return r.status_code == 200
        except Exception:
            return False

    def chat_stream(self, msg: str, on_token, on_done, on_error):
        self.history.append({"role": "user", "content": msg})
        messages = [{"role":"system","content":self.SYSTEM_PROMPT}] + self.history[-20:]

        def run():
            try:
                resp = requests.post(
                    f"{LM_STUDIO_URL}/v1/chat/completions",
                    json={
                        "model": MODEL_NAME,
                        "messages": messages,
                        "temperature": TEMPERATURE,
                        "max_tokens": MAX_TOKENS,
                        "stream": True,
                    },
                    stream=True,
                    timeout=60,
                )
                resp.raise_for_status()
                full = ""
                for line in resp.iter_lines():
                    if not line: continue
                    line = line.decode("utf-8")
                    if not line.startswith("data:"): continue
                    data = line[5:].strip()
                    if data == "[DONE]": break
                    try:
                        tok = json.loads(data)["choices"][0]["delta"].get("content", "")
                        if tok:
                            full += tok
                            on_token(tok)
                    except Exception:
                        pass
                self.history.append({"role":"assistant","content":full})
                on_done(full)
            except Exception as e:
                on_error(str(e))

        threading.Thread(target=run, daemon=True).start()

    def clear(self):
        self.history = []


# ════════════════════════════════════════════════════════════════════
#  FILE MANAGER
# ════════════════════════════════════════════════════════════════════
class FileManager:
    def __init__(self):
        self.locations = {
            "desktop":   DESKTOP,   "documents": DOCUMENTS, "downloads": DOWNLOADS,
            "home":      HOME,      "pictures":  PICTURES,
            "music":     MUSIC_DIR, "videos":    VIDEOS_DIR,
            "temp":      Path(os.environ.get("TEMP","C:/Temp")),
            "c":         Path("C:/"), "c drive": Path("C:/"),
            "d":         Path("D:/"), "d drive": Path("D:/"),
        }

    def resolve(self, raw: str) -> Path:
        raw   = raw.strip().strip("'\"")
        lower = raw.lower().rstrip("/\\ ")
        if lower in self.locations:
            return self.locations[lower]
        p = Path(raw)
        if p.is_absolute(): return p
        for base in [DESKTOP, DOCUMENTS, DOWNLOADS, HOME, MUSIC_DIR, VIDEOS_DIR]:
            c = base / raw
            if c.exists(): return c
        return DESKTOP / raw

    def open_file(self, s):
        p = self.resolve(s)
        if not p.exists():
            f = self._fuzzy(s, DESKTOP) or self._fuzzy(s, DOCUMENTS)
            if f: p = f
            else: return False, f"Not found: {p}"
        try: os.startfile(str(p)); return True, f"Opened: {p.name}"
        except Exception as e: return False, str(e)

    def create_file(self, s, content=""):
        p = self.resolve(s)
        try:
            p.parent.mkdir(parents=True, exist_ok=True)
            if p.exists(): return False, f"Already exists: {p.name}"
            p.write_text(content, encoding="utf-8")
            return True, f"Created: {p.name}  →  {p.parent}"
        except Exception as e: return False, str(e)

    def create_folder(self, s):
        p = self.resolve(s)
        try:
            if p.exists(): return False, f"Already exists: {p.name}"
            p.mkdir(parents=True)
            return True, f"Folder created: {p.name}  →  {p.parent}"
        except Exception as e: return False, str(e)

    def delete_path(self, s):
        p = self.resolve(s)
        if not p.exists(): return False, f"Not found: {p.name}"
        try:
            if HAS_TRASH:
                send2trash.send2trash(str(p))
                return True, f"Moved to Recycle Bin: {p.name}"
            shutil.rmtree(p) if p.is_dir() else p.unlink()
            return True, f"Deleted: {p.name}"
        except Exception as e: return False, str(e)

    def rename_path(self, old, new_name):
        p = self.resolve(old)
        if not p.exists(): return False, f"Not found: {p.name}"
        np = p.parent / new_name.strip().strip("'\"")
        if np.exists(): return False, f"'{new_name}' already exists"
        try: p.rename(np); return True, f"Renamed: {p.name}  →  {np.name}"
        except Exception as e: return False, str(e)

    def copy_path(self, src_s, dst_s):
        src = self.resolve(src_s); dst = self.resolve(dst_s)
        if not src.exists(): return False, f"Source not found: {src.name}"
        try:
            dst.mkdir(parents=True, exist_ok=True)
            dest = dst / src.name if dst.is_dir() else dst
            shutil.copytree(str(src), str(dest)) if src.is_dir() else shutil.copy2(str(src), str(dest))
            return True, f"Copied: {src.name}  →  {dest}"
        except Exception as e: return False, str(e)

    def move_path(self, src_s, dst_s):
        src = self.resolve(src_s); dst = self.resolve(dst_s)
        if not src.exists(): return False, f"Source not found: {src.name}"
        try:
            dst.mkdir(parents=True, exist_ok=True)
            dest = dst / src.name if dst.is_dir() else dst
            shutil.move(str(src), str(dest))
            return True, f"Moved: {src.name}  →  {dest}"
        except Exception as e: return False, str(e)

    def read_file(self, s, max_chars=2000):
        p = self.resolve(s)
        if not p.exists(): return False, f"Not found: {p}"
        if not p.is_file(): return False, f"'{p.name}' is a folder"
        try:
            text = p.read_text(encoding="utf-8", errors="replace")
            trunc = len(text) > max_chars
            out = f"── {p.name}  ({fmt_size(p.stat().st_size)}) ──\n{text[:max_chars]}"
            if trunc: out += f"\n... (truncated — {len(text)} chars total)"
            return True, out
        except Exception as e: return False, str(e)

    def write_file(self, s, content, append=False):
        p = self.resolve(s)
        try:
            p.parent.mkdir(parents=True, exist_ok=True)
            with p.open("a" if append else "w", encoding="utf-8") as f:
                f.write(content)
            action = "Appended to" if append else "Written to"
            return True, f"{action}: {p.name}  ({fmt_size(p.stat().st_size)})"
        except Exception as e: return False, str(e)

    def file_info(self, s):
        p = self.resolve(s)
        if not p.exists(): return False, f"Not found: {p}"
        try:
            st = p.stat()
            lines = [
                f"Name      : {p.name}",
                f"Type      : {'Folder' if p.is_dir() else 'File'}",
                f"Location  : {p.parent}",
            ]
            if p.is_file():
                lines.append(f"Size      : {fmt_size(st.st_size)}")
                lines.append(f"Extension : {p.suffix or '(none)'}")
            else:
                try: lines.append(f"Contents  : {len(list(p.iterdir()))} items")
                except: pass
            lines += [
                f"Created   : {fmt_time(st.st_ctime)}",
                f"Modified  : {fmt_time(st.st_mtime)}",
            ]
            return True, "\n".join(lines)
        except Exception as e: return False, str(e)

    def list_folder(self, s):
        p = self.resolve(s)
        if not p.exists(): return False, f"Not found: {p}"
        if not p.is_file() is False and p.is_file():
            return False, f"'{p.name}' is a file"
        try:
            items = sorted(p.iterdir(), key=lambda x:(x.is_file(), x.name.lower()))
            dirs  = [i for i in items if i.is_dir() and not i.name.startswith(".")]
            files = [i for i in items if i.is_file() and not i.name.startswith(".")]
            lines = [f"Contents: {p.name}/  ({len(dirs)} folders, {len(files)} files)", "─"*50]
            for d in dirs:  lines.append(f"  📁  {d.name}/")
            for f in files:
                try: sz = fmt_size(f.stat().st_size)
                except: sz = "?"
                lines.append(f"  📄  {f.name:<42} {sz}")
            if not dirs and not files: lines.append("  (empty)")
            return True, "\n".join(lines)
        except PermissionError: return False, f"Access denied: {p}"
        except Exception as e: return False, str(e)

    def tree_folder(self, s, max_depth=3):
        p = self.resolve(s)
        if not p.exists(): return False, f"Not found: {p}"
        lines = [f"Tree: {p}"]
        def walk(cur, prefix, depth):
            if depth > max_depth: return
            try:
                items = sorted(cur.iterdir(), key=lambda x:(x.is_file(),x.name.lower()))
                items = [i for i in items if not i.name.startswith(".")]
            except PermissionError:
                lines.append(f"{prefix}  [access denied]"); return
            for i, item in enumerate(items):
                last = i == len(items)-1
                lines.append(f"{prefix}{'└── ' if last else '├── '}{'📁 ' if item.is_dir() else '📄 '}{item.name}")
                if item.is_dir():
                    walk(item, prefix+("    " if last else "│   "), depth+1)
        walk(p, "", 1)
        if len(lines) > 60: lines = lines[:60]+["  ...(truncated)"]
        return True, "\n".join(lines)

    def search_files(self, pattern, loc="home", max_r=25):
        root = self.resolve(loc)
        if not root.exists(): return False, f"Not found: {root}"
        matches = []
        try:
            for p in root.rglob("*"):
                if fnmatch.fnmatch(p.name.lower(), pattern.lower()):
                    try: sz = fmt_size(p.stat().st_size) if p.is_file() else "folder"
                    except: sz = "?"
                    matches.append(f"  {'📁' if p.is_dir() else '📄'}  {p.name:<36} {sz:<10}  {p.parent}")
                if len(matches) >= max_r: break
        except PermissionError: pass
        if not matches: return False, f"Nothing matching '{pattern}' in {root.name}"
        return True, f"Found {len(matches)} result(s) for '{pattern}':\n"+"─"*55+"\n"+"\n".join(matches)

    def disk_usage(self, drive="C:"):
        if not HAS_PSUTIL: return False, "psutil not installed"
        try:
            u = psutil.disk_usage(drive+"\\")
            return True, (f"Disk {drive}\n"+"─"*30+
                          f"\n  Total : {fmt_size(u.total)}"
                          f"\n  Used  : {fmt_size(u.used)}  ({u.percent}%)"
                          f"\n  Free  : {fmt_size(u.free)}")
        except Exception as e: return False, str(e)

    def open_in_explorer(self, s):
        p = self.resolve(s)
        if not p.exists(): return False, f"Not found: {p}"
        try:
            subprocess.Popen(f'explorer "{p}"')
            return True, f"Opened in Explorer: {p.name}"
        except Exception as e: return False, str(e)

    def _fuzzy(self, name, in_dir):
        try:
            for p in in_dir.rglob("*"):
                if name.lower() in p.name.lower(): return p
        except: pass
        return None


# ════════════════════════════════════════════════════════════════════
#  MEDIA PLUGIN
# ════════════════════════════════════════════════════════════════════
class MediaPlugin:
    SEARCH_ROOTS = [MUSIC_DIR, VIDEOS_DIR, DESKTOP, DOWNLOADS, DOCUMENTS]

    def find_local(self, query, exts=None):
        exts  = exts or MEDIA_EXTS
        query = query.lower()
        found = []
        seen  = set()
        for root in self.SEARCH_ROOTS:
            if not root.exists(): continue
            try:
                for p in root.rglob("*"):
                    if p.suffix.lower() in exts and query in p.stem.lower() and str(p) not in seen:
                        found.append(p); seen.add(str(p))
                    if len(found) >= 25: break
            except: pass
        return found

    def list_local(self, root, exts=None):
        exts = exts or MEDIA_EXTS
        found = []
        try:
            for p in sorted(root.rglob("*"), key=lambda x: x.name.lower()):
                if p.suffix.lower() in exts:
                    found.append(p)
                if len(found) >= 100: break
        except: pass
        return found

    def youtube(self, query=None):
        url = f"https://www.youtube.com/results?search_query={urlquote(query)}" if query else "https://www.youtube.com"
        open_in_chrome(url)
        return "media", f"YouTube: '{query}'" if query else "Opening YouTube"

    def spotify(self, query=None):
        """
        Open Spotify and AUTOMATICALLY play the song — no click needed.

        Play pipeline:
        1. Get exact track ID from Spotify public API
        2. Open spotify:track:<id>  →  Spotify opens the song page
        3. Click the PLAY BUTTON in the Spotify taskbar/bottom-bar (always at
           a fixed position at bottom-center of the window) using PowerShell
           mouse_event — this is 100% reliable regardless of Spotify version.
        4. Fallback: send Space key (global play/pause for Spotify)
        """
        exe = find_spotify()

        # ── No query → just open Spotify ─────────────────────────────
        if not query:
            if exe:
                ok, msg = launch_app(exe, "Spotify")
                return ("media" if ok else "error"), msg
            try:
                os.startfile("spotify:")
                return "media", "Opening Spotify ✓"
            except Exception:
                pass
            open_in_chrome("https://open.spotify.com")
            return "media", "Spotify not installed — opened Web Player in browser."

        # ── Query → get track ID → open it → click play button ────────
        def _get_track_uri(song_query: str) -> str:
            """Get spotify:track:<id> for the top search result. Falls back to search URI."""
            fallback = f"spotify:search:{urlquote(song_query)}"
            try:
                # Try anonymous web player token (no API key needed)
                tok_resp = requests.get(
                    "https://open.spotify.com/get_access_token"
                    "?reason=transport&productType=web_player",
                    headers={"User-Agent": "Mozilla/5.0"},
                    timeout=5,
                )
                token = tok_resp.json().get("accessToken", "")
                if not token:
                    return fallback

                search_resp = requests.get(
                    f"https://api.spotify.com/v1/search"
                    f"?q={urlquote(song_query)}&type=track&limit=1",
                    headers={"Authorization": f"Bearer {token}"},
                    timeout=5,
                )
                items = search_resp.json().get("tracks", {}).get("items", [])
                if items:
                    return f"spotify:track:{items[0]['id']}"
            except Exception:
                pass
            return fallback

        # PowerShell script that:
        # 1. Waits for Spotify window  2. Brings it to front
        # 3. Gets window rect  4. Clicks the play button at bottom-center
        # 5. Fallback: sends Space key (Spotify global play/pause hotkey)
        PS_CLICK_PLAY = r"""
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class W32 {
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int n);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
    [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
    [DllImport("user32.dll")] public static extern void mouse_event(uint f, int x, int y, uint d, int e);
    [DllImport("user32.dll")] public static extern uint SendInput(uint n, INPUT[] i, int s);
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L, T, R, B; }
    [StructLayout(LayoutKind.Sequential)] public struct INPUT {
        public uint type;
        public KEYBDINPUT ki;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst=8)] public byte[] pad;
    }
    [StructLayout(LayoutKind.Sequential)] public struct KEYBDINPUT {
        public ushort wVk, wScan;
        public uint dwFlags, time;
        public IntPtr dwExtraInfo;
    }
    public const uint INPUT_KEYBOARD = 1;
    public const uint KEYEVENTF_KEYUP = 2;
    public const ushort VK_SPACE = 0x20;
    public static void Click(int x, int y) {
        SetCursorPos(x, y);
        System.Threading.Thread.Sleep(60);
        mouse_event(0x0002, x, y, 0, 0);  // MOUSEEVENTF_LEFTDOWN
        System.Threading.Thread.Sleep(60);
        mouse_event(0x0004, x, y, 0, 0);  // MOUSEEVENTF_LEFTUP
    }
    public static void PressSpace() {
        var down = new INPUT { type = INPUT_KEYBOARD, ki = new KEYBDINPUT { wVk = VK_SPACE } };
        var up   = new INPUT { type = INPUT_KEYBOARD, ki = new KEYBDINPUT { wVk = VK_SPACE, dwFlags = KEYEVENTF_KEYUP } };
        SendInput(1, new INPUT[]{down}, System.Runtime.InteropServices.Marshal.SizeOf(typeof(INPUT)));
        System.Threading.Thread.Sleep(50);
        SendInput(1, new INPUT[]{up},   System.Runtime.InteropServices.Marshal.SizeOf(typeof(INPUT)));
    }
}
"@

# Wait up to 12s for Spotify window to appear
$sw = $null
for ($i = 0; $i -lt 15; $i++) {
    $sw = Get-Process Spotify -ErrorAction SilentlyContinue |
          Where-Object { $_.MainWindowHandle -ne 0 } |
          Select-Object -First 1
    if ($sw) { break }
    Start-Sleep -Milliseconds 800
}
if (-not $sw) { exit 1 }

$hwnd = [IntPtr]$sw.MainWindowHandle
[W32]::ShowWindow($hwnd, 9) | Out-Null      # SW_RESTORE
[W32]::SetForegroundWindow($hwnd) | Out-Null
Start-Sleep -Milliseconds 1000

# Get window geometry
$rect = New-Object W32+RECT
[W32]::GetWindowRect($hwnd, [ref]$rect) | Out-Null
$W = $rect.R - $rect.L
$H = $rect.B - $rect.T

# Spotify play button is ALWAYS at bottom-center of the window
# Exact position: horizontally centered, vertically at ~93% from top
$playX = $rect.L + [int]($W * 0.50)
$playY = $rect.T + [int]($H * 0.935)

# Wait for track page to load
Start-Sleep -Milliseconds 2200

# Click the play button
[W32]::Click($playX, $playY)
Start-Sleep -Milliseconds 400

# Send Space as a backup (Spotify's global play/pause shortcut)
[W32]::SetForegroundWindow($hwnd) | Out-Null
Start-Sleep -Milliseconds 200
[W32]::PressSpace()
"""

        def _ensure_running():
            try:
                chk = subprocess.run(
                    ["tasklist", "/FI", "IMAGENAME eq Spotify.exe", "/NH"],
                    capture_output=True, text=True, timeout=4
                )
                if "Spotify.exe" in chk.stdout:
                    return True
            except Exception:
                pass
            if exe:
                try:
                    subprocess.Popen([exe])
                    time.sleep(5.0)
                    return True
                except Exception:
                    pass
            return False

        def _play_uri(uri: str):
            try:
                os.startfile(uri)
                return True
            except Exception:
                pass
            if exe:
                try:
                    subprocess.Popen([exe, uri])
                    return True
                except Exception:
                    pass
            return False

        def _click_play_button():
            try:
                subprocess.Popen(
                    ["powershell", "-NoProfile", "-NonInteractive",
                     "-WindowStyle", "Hidden", "-Command", PS_CLICK_PLAY],
                    creationflags=subprocess.CREATE_NO_WINDOW
                )
            except Exception:
                pass

        def _open_and_play():
            _ensure_running()
            track_uri = _get_track_uri(query)
            _play_uri(track_uri)
            # After opening the track page, click play + send Space
            _click_play_button()

        if exe:
            threading.Thread(target=_open_and_play, daemon=True).start()
            return "media", f"Playing '{query}' on Spotify ✓"

        # MS Store Spotify
        try:
            os.startfile(f"spotify:search:{urlquote(query)}")
            threading.Thread(target=_click_play_button, daemon=True).start()
            return "media", f"Playing '{query}' on Spotify ✓"
        except Exception:
            pass

        open_in_chrome(f"https://open.spotify.com/search/{urlquote(query)}")
        return "media", f"Spotify not installed — opened Web Player for '{query}'."

    def vlc_play(self, target=None):
        exe = find_exe(VLC_PATHS)
        if not exe: return "error", "VLC not found — install from videolan.org"
        if target:
            results = self.find_local(target)
            if results:
                subprocess.Popen([exe] + [str(r) for r in results[:30]])
                return "media", f"VLC playing {len(results)} file(s): '{target}'"
            p = Path(target)
            if p.exists(): subprocess.Popen([exe, str(p)]); return "media", f"VLC: {p.name}"
            return "error", f"No media found for '{target}'"
        subprocess.Popen([exe])
        return "media", "Opening VLC"

    def wmp_play(self, target=None):
        exe = find_exe(WMP_PATHS)
        if target:
            results = self.find_local(target)
            if results:
                p = results[0]
                if exe: subprocess.Popen([exe, str(p)])
                else: os.startfile(str(p))
                return "media", f"Playing: {p.name}"
        if exe: subprocess.Popen([exe]); return "media", "Opening Windows Media Player"
        return "error", "Windows Media Player not found"

    def smart_play(self, query):
        results = self.find_local(query)
        if results:
            p = results[0]
            vlc = find_exe(VLC_PATHS)
            if vlc: subprocess.Popen([vlc, str(p)])
            else:
                try: os.startfile(str(p))
                except: pass
            return "media", f"Playing: {p.name}"
        open_in_chrome(f"https://www.youtube.com/results?search_query={urlquote(query)}")
        return "media", f"No local file found — searching YouTube for '{query}'"

    def find_music(self, query):
        results = self.find_local(query, AUDIO_EXTS)
        if not results: return "error", f"No music matching '{query}'"
        lines = [f"Music: '{query}' ({len(results)} found)", "─"*50]
        for p in results:
            try: sz = fmt_size(p.stat().st_size)
            except: sz = "?"
            lines.append(f"  🎵  {p.name:<40}  {sz}\n      {p.parent}")
        return "info", "\n".join(lines)

    def find_video(self, query):
        results = self.find_local(query, VIDEO_EXTS)
        if not results: return "error", f"No video matching '{query}'"
        lines = [f"Videos: '{query}' ({len(results)} found)", "─"*50]
        for p in results:
            try: sz = fmt_size(p.stat().st_size)
            except: sz = "?"
            lines.append(f"  🎬  {p.name:<40}  {sz}\n      {p.parent}")
        return "info", "\n".join(lines)

    def list_music(self):
        files = self.list_local(MUSIC_DIR, AUDIO_EXTS)
        if not files: return "error", "No music files found in Music folder"
        lines = [f"Music Library ({len(files)} files)", "─"*50]
        for p in files[:40]: lines.append(f"  🎵  {p.name}")
        if len(files) > 40: lines.append(f"  ... and {len(files)-40} more")
        return "info", "\n".join(lines)

    def list_videos(self):
        files = self.list_local(VIDEOS_DIR, VIDEO_EXTS)
        if not files: return "error", "No video files found in Videos folder"
        lines = [f"Video Library ({len(files)} files)", "─"*50]
        for p in files[:40]: lines.append(f"  🎬  {p.name}")
        if len(files) > 40: lines.append(f"  ... and {len(files)-40} more")
        return "info", "\n".join(lines)


# ════════════════════════════════════════════════════════════════════
#  UNIVERSAL APP LAUNCHER ENGINE
#  Finds ANY app installed on the system through 5 discovery layers:
#  1. Known paths with existence check
#  2. Windows Start Menu / Desktop .lnk scan
#  3. Windows Registry (App Paths)
#  4. `where` command (PATH lookup)
#  5. WindowsApps folder (Microsoft Store apps)
# ════════════════════════════════════════════════════════════════════

def _find_in_registry(app_name: str) -> str | None:
    """Look up an app in Windows App Paths registry key."""
    try:
        import winreg
        key_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths"
        for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            try:
                with winreg.OpenKey(hive, key_path) as base:
                    # Try direct lookup with .exe
                    for suffix in ("", ".exe"):
                        try:
                            with winreg.OpenKey(base, app_name + suffix) as key:
                                val, _ = winreg.QueryValueEx(key, "")
                                if val and Path(val).exists():
                                    return val
                        except FileNotFoundError:
                            pass
            except Exception:
                pass
    except ImportError:
        pass
    return None


def _find_in_windowsapps(name: str) -> str | None:
    """Search Microsoft Store apps in WindowsApps folder."""
    wa = HOME / r"AppData\Local\Microsoft\WindowsApps"
    if wa.exists():
        nl = name.lower().replace(" ", "")
        try:
            for p in wa.rglob("*.exe"):
                if nl in p.stem.lower().replace(" ", ""):
                    return str(p)
        except Exception:
            pass
    return None


def _find_via_where(name: str) -> str | None:
    """Use Windows 'where' command to find exe in PATH."""
    try:
        exe = name if name.endswith(".exe") else name + ".exe"
        r = subprocess.run(
            ["where", exe], capture_output=True, text=True, timeout=3
        )
        if r.returncode == 0:
            line = r.stdout.strip().splitlines()[0].strip()
            if line and Path(line).exists():
                return line
    except Exception:
        pass
    return None


def smart_find_app(name: str) -> str | None:
    """
    Universal app finder. Given an app name (like 'spotify', 'chrome', 'discord'),
    searches all 5 layers and returns the full exe path or URI, or None.
    """
    nl = name.lower().strip()

    # ── Layer 1: winreg App Paths ─────────────────────────────────
    reg = _find_in_registry(nl)
    if reg:
        return reg

    # ── Layer 2: `where` command (PATH) ──────────────────────────
    wh = _find_via_where(nl)
    if wh:
        return wh

    # ── Layer 3: WindowsApps (Microsoft Store) ────────────────────
    wa = _find_in_windowsapps(nl)
    if wa:
        return wa

    # ── Layer 4: Start Menu .lnk scan ─────────────────────────────
    search_bases = [
        DESKTOP,
        Path("C:/Users/Public/Desktop"),
        HOME / "AppData/Roaming/Microsoft/Windows/Start Menu/Programs",
        Path("C:/ProgramData/Microsoft/Windows/Start Menu/Programs"),
    ]
    for base in search_bases:
        if not base.exists():
            continue
        try:
            for p in base.rglob("*.lnk"):
                if nl in p.stem.lower():
                    return str(p)
        except Exception:
            pass

    return None


def launch_app(path_or_uri: str, name: str = "") -> tuple[bool, str]:
    """
    Launch any application given its path, .lnk, or URI scheme.
    Returns (success, message).
    """
    label = name.title() if name else Path(path_or_uri).stem.title()
    try:
        if path_or_uri.startswith("ms-") or path_or_uri.startswith("spotify:"):
            # URI scheme — use os.startfile
            os.startfile(path_or_uri)
            return True, f"Opening {label} ✓"

        p = Path(path_or_uri)

        if p.suffix.lower() == ".lnk":
            # Shortcut — use os.startfile which resolves .lnk natively
            os.startfile(str(p))
            return True, f"Opening {label} ✓"

        if p.suffix.lower() in (".exe", ".msc", ".bat", ".cmd"):
            subprocess.Popen([str(p)], shell=False)
            return True, f"Opening {label} ✓"

        # Fallback: shell=True handles edge cases
        os.startfile(str(p))
        return True, f"Opening {label} ✓"

    except Exception as e:
        # Last resort: try shell=True
        try:
            subprocess.Popen(path_or_uri, shell=True)
            return True, f"Opening {label} ✓"
        except Exception:
            return False, f"Could not open {label}: {e}"


class AppScanner:
    def __init__(self):
        self.apps: dict[str, str] = {}
        self._build()

    def _build(self):
        """
        Build the app registry. Only registers apps that actually exist.
        Uses smart_find_app() as the universal discovery engine.
        """
        # ── Windows built-ins (always available, no path check needed) ─
        builtins = {
            "notepad":            "notepad.exe",
            "calculator":         "calc.exe",
            "paint":              "mspaint.exe",
            "file explorer":      "explorer.exe",
            "explorer":           "explorer.exe",
            "task manager":       "taskmgr.exe",
            "taskmgr":            "taskmgr.exe",
            "control panel":      "control.exe",
            "cmd":                "cmd.exe",
            "command prompt":     "cmd.exe",
            "powershell":         "powershell.exe",
            "windows powershell": "powershell.exe",
            "settings":           "ms-settings:",
            "windows settings":   "ms-settings:",
            "snipping tool":      "snippingtool.exe",
            "snip":               "snippingtool.exe",
            "wordpad":            "wordpad.exe",
            "regedit":            "regedit.exe",
            "registry editor":    "regedit.exe",
            "msconfig":           "msconfig.exe",
            "device manager":     "devmgmt.msc",
            "disk management":    "diskmgmt.msc",
            "event viewer":       "eventvwr.msc",
            "services":           "services.msc",
            "resource monitor":   "resmon.exe",
            "performance monitor":"perfmon.exe",
            "character map":      "charmap.exe",
            "magnifier":          "magnify.exe",
            "on-screen keyboard": "osk.exe",
            "narrator":           "narrator.exe",
        }
        self.apps.update(builtins)

        # ── Apps with known candidate paths (check existence first) ───
        known = [
            # ( [aliases],  [candidate paths] )
            (
                ["chrome", "google chrome", "google"],
                CHROME_PATHS
            ),
            (
                ["firefox", "mozilla firefox", "mozilla"],
                [
                    r"C:\Program Files\Mozilla Firefox\firefox.exe",
                    r"C:\Program Files (x86)\Mozilla Firefox\firefox.exe",
                    str(HOME / r"AppData\Local\Mozilla Firefox\firefox.exe"),
                ]
            ),
            (
                ["edge", "microsoft edge", "msedge"],
                [
                    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
                    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
                    str(HOME / r"AppData\Local\Microsoft\Edge\Application\msedge.exe"),
                ]
            ),
            (
                ["spotify"],
                SPOTIFY_PATHS
            ),
            (
                ["vlc", "vlc media player", "video lan"],
                VLC_PATHS
            ),
            (
                ["windows media player", "media player", "wmp"],
                WMP_PATHS
            ),
            (
                ["word", "microsoft word", "ms word"],
                [
                    r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",
                    r"C:\Program Files\Microsoft Office\root\Office17\WINWORD.EXE",
                    r"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE",
                    r"C:\Program Files\Microsoft Office 15\root\Office15\WINWORD.EXE",
                    "winword.exe",
                ]
            ),
            (
                ["excel", "microsoft excel", "ms excel"],
                [
                    r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE",
                    r"C:\Program Files\Microsoft Office\root\Office17\EXCEL.EXE",
                    r"C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE",
                    "excel.exe",
                ]
            ),
            (
                ["powerpoint", "microsoft powerpoint", "ms powerpoint", "ppt"],
                [
                    r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE",
                    r"C:\Program Files\Microsoft Office\root\Office17\POWERPNT.EXE",
                    r"C:\Program Files (x86)\Microsoft Office\root\Office16\POWERPNT.EXE",
                    "powerpnt.exe",
                ]
            ),
            (
                ["outlook", "microsoft outlook", "ms outlook"],
                [
                    r"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE",
                    r"C:\Program Files\Microsoft Office\root\Office17\OUTLOOK.EXE",
                    r"C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE",
                    "outlook.exe",
                ]
            ),
            (
                ["onenote", "microsoft onenote", "one note"],
                [
                    r"C:\Program Files\Microsoft Office\root\Office16\ONENOTE.EXE",
                    r"C:\Program Files (x86)\Microsoft Office\root\Office16\ONENOTE.EXE",
                    "onenote.exe",
                ]
            ),
            (
                ["teams", "microsoft teams", "ms teams"],
                [
                    str(HOME / r"AppData\Local\Microsoft\Teams\current\Teams.exe"),
                    str(HOME / r"AppData\Local\Microsoft\Teams\Teams.exe"),
                    str(HOME / r"AppData\Local\Microsoft\TeamsMeetingAddin\Teams.exe"),
                    "msteams.exe",
                ]
            ),
            (
                ["discord"],
                [
                    str(HOME / r"AppData\Local\Discord\Update.exe"),
                    str(HOME / r"AppData\Roaming\Discord\Discord.exe"),
                    "discord.exe",
                ]
            ),
            (
                ["telegram", "telegram desktop"],
                [
                    str(HOME / r"AppData\Roaming\Telegram Desktop\Telegram.exe"),
                    str(HOME / r"AppData\Local\Telegram Desktop\Telegram.exe"),
                    str(HOME / r"AppData\Roaming\Telegram\Telegram.exe"),
                ]
            ),
            (
                ["whatsapp", "whats app"],
                [
                    str(HOME / r"AppData\Local\WhatsApp\WhatsApp.exe"),
                    str(HOME / r"AppData\Roaming\WhatsApp\WhatsApp.exe"),
                    str(HOME / r"AppData\Local\Programs\WhatsApp\WhatsApp.exe"),
                ]
            ),
            (
                ["zoom", "zoom meetings"],
                [
                    str(HOME / r"AppData\Roaming\Zoom\bin\Zoom.exe"),
                    str(HOME / r"AppData\Roaming\Zoom\Zoom.exe"),
                    r"C:\Program Files\Zoom\bin\Zoom.exe",
                    "zoom.exe",
                ]
            ),
            (
                ["steam"],
                [
                    r"C:\Program Files (x86)\Steam\steam.exe",
                    r"C:\Program Files\Steam\steam.exe",
                ]
            ),
            (
                ["obs", "obs studio"],
                [
                    r"C:\Program Files\obs-studio\bin\64bit\obs64.exe",
                    r"C:\Program Files (x86)\obs-studio\bin\32bit\obs32.exe",
                ]
            ),
            (
                ["notepad++", "notepad plus", "npp"],
                [
                    r"C:\Program Files\Notepad++\notepad++.exe",
                    r"C:\Program Files (x86)\Notepad++\notepad++.exe",
                ]
            ),
            (
                ["winrar"],
                [
                    r"C:\Program Files\WinRAR\WinRAR.exe",
                    r"C:\Program Files (x86)\WinRAR\WinRAR.exe",
                ]
            ),
            (
                ["7zip", "7-zip", "7 zip"],
                [
                    r"C:\Program Files\7-Zip\7zFM.exe",
                    r"C:\Program Files (x86)\7-Zip\7zFM.exe",
                ]
            ),
            (
                ["git", "git bash"],
                [
                    r"C:\Program Files\Git\bin\git.exe",
                    r"C:\Program Files (x86)\Git\bin\git.exe",
                    r"C:\Program Files\Git\git-bash.exe",
                ]
            ),
            (
                ["python", "python3"],
                [
                    str(HOME / r"AppData\Local\Programs\Python\Python312\python.exe"),
                    str(HOME / r"AppData\Local\Programs\Python\Python311\python.exe"),
                    str(HOME / r"AppData\Local\Programs\Python\Python310\python.exe"),
                    r"C:\Python312\python.exe",
                    r"C:\Python311\python.exe",
                    "python.exe",
                ]
            ),
            (
                ["pycharm"],
                [
                    str(HOME / r"AppData\Local\JetBrains\Toolbox\apps\PyCharm-P\ch-0\bin\pycharm64.exe"),
                    r"C:\Program Files\JetBrains\PyCharm Community Edition\bin\pycharm64.exe",
                    r"C:\Program Files\JetBrains\PyCharm Professional Edition\bin\pycharm64.exe",
                ]
            ),
            (
                ["android studio"],
                [
                    r"C:\Program Files\Android\Android Studio\bin\studio64.exe",
                    str(HOME / r"AppData\Local\Programs\Android Studio\bin\studio64.exe"),
                ]
            ),
            (
                ["figma"],
                [
                    str(HOME / r"AppData\Local\Figma\Figma.exe"),
                    str(HOME / r"AppData\Roaming\Figma\Figma.exe"),
                ]
            ),
            (
                ["postman"],
                [
                    str(HOME / r"AppData\Local\Postman\Postman.exe"),
                    str(HOME / r"AppData\Roaming\Postman\Postman.exe"),
                ]
            ),
            (
                ["slack"],
                [
                    str(HOME / r"AppData\Local\slack\slack.exe"),
                    str(HOME / r"AppData\Roaming\Slack\slack.exe"),
                    "slack.exe",
                ]
            ),
            (
                ["skype"],
                [
                    str(HOME / r"AppData\Roaming\Skype\Skype.exe"),
                    str(HOME / r"AppData\Local\Microsoft\Skype\Skype.exe"),
                    "skype.exe",
                ]
            ),
            (
                ["epic games", "epic games launcher"],
                [
                    r"C:\Program Files (x86)\Epic Games\Launcher\Portal\Binaries\Win32\EpicGamesLauncher.exe",
                    r"C:\Program Files\Epic Games\Launcher\Portal\Binaries\Win64\EpicGamesLauncher.exe",
                ]
            ),
            (
                ["blender"],
                [
                    r"C:\Program Files\Blender Foundation\Blender\blender.exe",
                    r"C:\Program Files\Blender Foundation\Blender 4.0\blender.exe",
                    r"C:\Program Files\Blender Foundation\Blender 3.6\blender.exe",
                ]
            ),
            (
                ["gimp"],
                [
                    r"C:\Program Files\GIMP 2\bin\gimp-2.10.exe",
                    r"C:\Program Files\GIMP 3\bin\gimp-3.0.exe",
                ]
            ),
            (
                ["audacity"],
                [
                    r"C:\Program Files\Audacity\Audacity.exe",
                    r"C:\Program Files (x86)\Audacity\Audacity.exe",
                ]
            ),
            (
                ["winamp"],
                [
                    r"C:\Program Files\Winamp\winamp.exe",
                    r"C:\Program Files (x86)\Winamp\winamp.exe",
                ]
            ),
            (
                ["itunes"],
                [
                    r"C:\Program Files\iTunes\iTunes.exe",
                    r"C:\Program Files (x86)\iTunes\iTunes.exe",
                ]
            ),
            (
                ["foobar", "foobar2000"],
                [
                    r"C:\Program Files\foobar2000\foobar2000.exe",
                    r"C:\Program Files (x86)\foobar2000\foobar2000.exe",
                ]
            ),
            (
                ["brave", "brave browser"],
                [
                    str(HOME / r"AppData\Local\BraveSoftware\Brave-Browser\Application\brave.exe"),
                    r"C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe",
                ]
            ),
            (
                ["opera"],
                [
                    str(HOME / r"AppData\Local\Programs\Opera\opera.exe"),
                    r"C:\Program Files\Opera\opera.exe",
                ]
            ),
            (
                ["vivaldi"],
                [
                    str(HOME / r"AppData\Local\Vivaldi\Application\vivaldi.exe"),
                ]
            ),
            (
                ["tor", "tor browser"],
                [
                    str(DESKTOP / r"Tor Browser\Browser\firefox.exe"),
                    str(HOME / r"Desktop\Tor Browser\Browser\firefox.exe"),
                ]
            ),
        ]

        for aliases, paths in known:
            exe = find_exe(paths)
            if not exe:
                # Try smart_find_app with the primary name
                exe = smart_find_app(aliases[0])
            if exe:
                for alias in aliases:
                    self.apps[alias] = exe

        # ── VS Code (special path pattern) ────────────────────────────
        vsc_paths = [
            HOME / "AppData/Local/Programs/Microsoft VS Code/Code.exe",
            Path(r"C:\Program Files\Microsoft VS Code\Code.exe"),
        ]
        for vsc in vsc_paths:
            if vsc.exists():
                for k in ["vs code", "vscode", "visual studio code", "code"]:
                    self.apps[k] = str(vsc)
                break

        # ── Scan Start Menu + Desktop for .lnk shortcuts ──────────────
        # This catches anything not in our known list
        scan_bases = [
            DESKTOP,
            Path("C:/Users/Public/Desktop"),
            HOME / "AppData/Roaming/Microsoft/Windows/Start Menu/Programs",
            Path("C:/ProgramData/Microsoft/Windows/Start Menu/Programs"),
        ]
        skip = ("uninstall", "setup", "installer", "update", "repair",
                "remove", "readme", "help", "crash", "bug", "log")
        for base in scan_bases:
            if not base.exists():
                continue
            for pat in ("**/*.lnk", "**/*.exe", "**/*.url"):
                try:
                    for p in base.glob(pat):
                        n = p.stem.lower().strip()
                        if n and not any(x in n for x in skip) and n not in self.apps:
                            self.apps[n] = str(p)
                except Exception:
                    pass

    # ──────────────────────────────────────────────────────────────────
    def find(self, name: str) -> str | None:
        """
        Find an app path by name. Multi-layer lookup:
        1. Exact match in registry
        2. Partial/prefix/substring match in registry
        3. smart_find_app() universal search
        """
        n = name.lower().strip()

        # Exact match
        if n in self.apps:
            return self.apps[n]

        # Prefix or suffix match (e.g. "google" → "google chrome")
        for k, v in self.apps.items():
            if k.startswith(n) or n.startswith(k):
                return v

        # Substring match
        for k, v in self.apps.items():
            if n in k or k in n:
                return v

        # Universal fallback search (registry + windowsapps + where + lnk)
        found = smart_find_app(n)
        if found:
            self.apps[n] = found   # cache it for next time
            return found

        return None

    def list_all(self) -> list:
        return sorted(self.apps.keys())


# ════════════════════════════════════════════════════════════════════
#  UNIFIED COMMAND PARSER
# ════════════════════════════════════════════════════════════════════
class CommandParser:
    def __init__(self, fm: FileManager, media: MediaPlugin,
                 apps: AppScanner, voice: VoiceEngine):
        self.fm    = fm
        self.media = media
        self.apps  = apps
        self.voice = voice

    # ── Siri-like NLU helpers ─────────────────────────────────────────
    @staticmethod
    def _strip_prefix(text: str, prefixes: list) -> str | None:
        """Remove a matching prefix and return the rest, or None."""
        tl = text.lower()
        for p in sorted(prefixes, key=len, reverse=True):
            if tl.startswith(p):
                return text[len(p):].strip().strip("'\"")
        return None

    @staticmethod
    def _eval_math(expr: str) -> str | None:
        """Safely evaluate a math expression. Returns result string or None."""
        try:
            # normalise spoken math
            expr = expr.replace("^", "**").replace("×","*").replace("÷","/")
            expr = re.sub(r'\bsquared\b', '**2', expr)
            expr = re.sub(r'\bcubed\b',   '**3', expr)
            expr = re.sub(r'\bsquare root of\b', 'math.sqrt', expr)
            expr = re.sub(r'\bsqrt\b',           'math.sqrt', expr)
            expr = re.sub(r'\bpi\b',             'math.pi',   expr)
            # Only allow safe characters
            if re.search(r'[a-zA-Z]', expr.replace("math.sqrt","").replace("math.pi","")):
                return None
            result = eval(expr, {"__builtins__": {}, "math": math})
            if isinstance(result, float) and result == int(result):
                return str(int(result))
            return f"{result:.6g}"
        except Exception:
            return None

    @staticmethod
    def _greet_response() -> str:
        import random
        h = datetime.now().hour
        time_greet = "Good morning" if h < 12 else ("Good afternoon" if h < 17 else "Good evening")
        responses = [
            f"{time_greet}! How can I help you today?",
            f"{time_greet}! I'm Stark, your personal assistant. What do you need?",
            f"Hey! I'm here and ready. What can I do for you?",
            f"{time_greet}! All systems online. What's on your mind?",
            f"Hello! Ready to assist. Just say the word.",
        ]
        return random.choice(responses)

    @staticmethod
    def _farewell_response() -> str:
        import random
        responses = [
            "Goodbye! I'll be right here if you need me.",
            "See you! Stay productive.",
            "Later! The orb is always watching.",
            "Goodbye! Have a great day.",
            "Signing off. Click the orb anytime.",
        ]
        return random.choice(responses)

    @staticmethod
    def _thanks_response() -> str:
        import random
        responses = [
            "You're welcome! Anything else?",
            "Happy to help! What's next?",
            "Anytime! That's what I'm here for.",
            "No problem at all. Need anything else?",
            "Of course! Just ask.",
        ]
        return random.choice(responses)

    @staticmethod
    def _joke() -> str:
        import random
        jokes = [
            "Why do programmers prefer dark mode? Because light attracts bugs! 🐛",
            "I asked my AI to tell me a joke. It said 'Error 404: Humor not found.' 😅",
            "Why was the computer cold? Because it left its Windows open! 🖥️",
            "What do you call a fish with no eyes? A fsh. 🐟",
            "How many programmers does it take to change a light bulb? None — that's a hardware problem.",
            "I told my computer I needed a break... Now it won't stop sending me Kit-Kat ads. 🍫",
            "Why do Java developers wear glasses? Because they don't C#! 👓",
            "A SQL query walks into a bar, walks up to two tables and asks... 'Can I join you?'",
        ]
        return random.choice(jokes)

    @staticmethod
    def _motivate() -> str:
        import random
        quotes = [
            "The secret of getting ahead is getting started. — Mark Twain",
            "It always seems impossible until it's done. — Nelson Mandela",
            "Don't watch the clock; do what it does. Keep going. — Sam Levenson",
            "Success is not final, failure is not fatal: it is the courage to continue that counts.",
            "The only way to do great work is to love what you do. — Steve Jobs",
            "Believe you can and you're halfway there. — Theodore Roosevelt",
            "You are capable of more than you know. Keep pushing.",
            "Every expert was once a beginner. Keep going.",
        ]
        return random.choice(quotes)

    def parse(self, text: str):
        """
        Siri-like NLU engine.
        Returns (tag, response) if handled locally, else None → AI.
        Handles: greetings, farewells, gratitude, identity, time/date,
        math, system info, voice control, media, files, apps, web search,
        jokes, motivation, clipboard, help, and more.
        """
        import random
        t  = text.strip()
        tl = t.lower().strip()
        # remove filler words at the very start for cleaner matching
        tl_clean = re.sub(r'^(hey\s+stark[,!]?\s*|stark[,!]?\s*|ok\s+stark[,!]?\s*)', '', tl).strip()

        # ── 1. GREETINGS ──────────────────────────────────────────────
        _greet_kw = ["hello","hi","hey","howdy","good morning","good afternoon",
                     "good evening","what's up","sup","greetings","yo stark",
                     "wake up","are you there","you there","hello stark"]
        if any(tl_clean == k or tl_clean.startswith(k+" ") for k in _greet_kw):
            r = self._greet_response()
            self.voice.speak(r, force=True)
            return "stark", r

        # ── 2. FAREWELLS ──────────────────────────────────────────────
        _bye_kw = ["bye","goodbye","see you","see ya","later","good night",
                   "goodnight","that's all","that will be all","close","exit"]
        if any(tl_clean == k or tl_clean.startswith(k) for k in _bye_kw):
            r = self._farewell_response()
            self.voice.speak(r, force=True)
            return "stark", r

        # ── 3. THANKS / GRATITUDE ─────────────────────────────────────
        _thanks_kw = ["thank you","thanks","thank u","thx","cheers","appreciate it",
                      "great job","well done","nice work","good job","awesome","perfect"]
        if any(tl_clean == k or tl_clean.startswith(k) for k in _thanks_kw):
            r = self._thanks_response()
            self.voice.speak(r, force=True)
            return "stark", r

        # ── 4. IDENTITY / WHO ARE YOU ─────────────────────────────────
        _id_kw = ["who are you","what are you","what's your name","your name",
                  "introduce yourself","tell me about yourself","what can you do",
                  "what do you do","help me","show me what you can do","capabilities",
                  "features","what are your features"]
        if any(k in tl_clean for k in _id_kw):
            r = (
                "I'm Stark — your personal AI desktop assistant, inspired by Jarvis.\n\n"
                "Here's what I can do for you:\n"
                "  🗣  Voice  — speak responses, adjust speed & volume\n"
                "  📁  Files  — create, delete, rename, copy, move, read, search\n"
                "  🎵  Media  — play music/videos, open YouTube & Spotify\n"
                "  🖥  Apps   — launch any installed application\n"
                "  🌐  Web    — search Google, open Wikipedia\n"
                "  ⏱  Time   — tell the time, date, and day\n"
                "  💻  System — CPU, RAM, battery, disk info\n"
                "  🧮  Math   — calculate any expression instantly\n"
                "  🤖  AI     — chat with LM Studio for anything else\n\n"
                "Just talk to me naturally — I'll figure out what you need."
            )
            self.voice.speak("I'm Stark, your personal AI desktop assistant. I can help with files, apps, music, system info, and much more. Just ask!", force=True)
            return "stark", r

        # ── 5. HOW ARE YOU ────────────────────────────────────────────
        _how_kw = ["how are you","how are you doing","how do you do","you okay",
                   "are you okay","you good","you alright"]
        if any(k in tl_clean for k in _how_kw):
            responses = [
                "Running at peak efficiency! All systems nominal. How about you?",
                "I'm operating perfectly, thanks for asking! What can I help with?",
                "Excellent — no bugs, no errors. Ready to assist!",
                "I'm great! Fully charged and ready. What do you need?",
            ]
            r = random.choice(responses)
            self.voice.speak(r, force=True)
            return "stark", r

        # ── 6. JOKES ──────────────────────────────────────────────────
        _joke_kw = ["tell me a joke","joke","make me laugh","say something funny",
                    "funny","humor me","entertain me"]
        if any(k in tl_clean for k in _joke_kw):
            r = self._joke()
            self.voice.speak(r, force=True)
            return "stark", r

        # ── 7. MOTIVATION / QUOTES ────────────────────────────────────
        _motive_kw = ["motivate me","inspire me","motivation","inspirational quote",
                      "quote of the day","give me a quote","say something inspiring",
                      "encouragement","encourage me"]
        if any(k in tl_clean for k in _motive_kw):
            r = self._motivate()
            self.voice.speak(r, force=True)
            return "stark", r

        # ── 8. VOICE CONTROL ──────────────────────────────────────────
        m = re.match(r"(?:say|speak|read\s+aloud|tell\s+me)\s+(.+)", tl_clean)
        if m:
            phrase = m.group(1).strip("'\"")
            self.voice.speak(phrase, force=True)
            return "action", f"Speaking: \"{phrase}\""

        _stop_kw = ["stop speaking","stop talking","be quiet","silence","shush",
                    "quiet","stop voice","shut up","stop it","enough"]
        if any(k in tl_clean for k in _stop_kw):
            self.voice.stop()
            return "action", "Stopped."

        _von_kw = ["voice on","enable voice","turn on voice","unmute",
                   "start speaking","speak to me"]
        if any(k in tl_clean for k in _von_kw):
            if not self.voice.enabled: self.voice.toggle()
            return "action", "Voice enabled — I'll speak my responses."

        _voff_kw = ["voice off","disable voice","turn off voice","mute voice",
                    "stop speaking automatically","no voice","silent mode"]
        if any(k in tl_clean for k in _voff_kw):
            if self.voice.enabled: self.voice.toggle()
            return "action", "Voice disabled — silent mode on."

        m = re.search(r"(?:set\s+)?(?:voice\s+)?speed\s+(?:to\s+)?(slow|normal|fast|\d+)", tl_clean)
        if m:
            val  = m.group(1)
            rate = {"slow":120,"normal":175,"fast":260}.get(val, int(val) if val.isdigit() else 175)
            self.voice.set_rate(rate)
            return "action", f"Voice speed set to {rate} wpm."

        m = re.search(r"(?:set\s+)?(?:voice\s+)?volume\s+(?:to\s+)?(\d+)", tl_clean)
        if m:
            vol = max(0, min(100, int(m.group(1))))
            self.voice.set_volume(vol/100)
            return "action", f"Volume set to {vol}%."

        # ── 9. TIME & DATE ────────────────────────────────────────────
        _time_kw = ["what time","time now","current time","what's the time",
                    "tell me the time","time is it","clock"]
        if any(k in tl_clean for k in _time_kw):
            now = datetime.now()
            r = f"It's {now.strftime('%I:%M %p')}  —  {now.strftime('%A, %B %d %Y')}"
            self.voice.speak(r)
            return "info", r

        _date_kw = ["what date","today's date","what day","current date",
                    "what is today","date today","what's today"]
        if any(k in tl_clean for k in _date_kw):
            r = f"Today is {datetime.now().strftime('%A, %B %d, %Y')}"
            self.voice.speak(r)
            return "info", r

        # day of week
        m = re.search(r"what day (?:is it|of the week)", tl_clean)
        if m:
            r = f"Today is {datetime.now().strftime('%A')}."
            self.voice.speak(r)
            return "info", r

        # ── 10. MATH / CALCULATOR ─────────────────────────────────────
        _calc_kw = ["calculate","compute","what is","whats","eval","solve",
                    "how much is","work out","figure out","what's"]
        # Direct math expressions like "12 * 34" or "sqrt(9)"
        _math_expr = re.match(
            r"(?:(?:calculate|compute|what(?:'?s|\s+is)\s+|eval|solve|how\s+much\s+is\s+|work\s+out\s+|figure\s+out\s+)?)"
            r"([\d\s\.\+\-\*\/\^\(\)%]+(?:[\^\*]{2}[\d]+)?|sqrt\([\d\.]+\)|(?:square\s+root\s+of\s+\d+))",
            tl_clean
        )
        if _math_expr:
            expr = _math_expr.group(1).strip()
            if any(op in expr for op in ["+","-","*","/","^","%","sqrt","**"]) or re.search(r"\d\s+\d", expr):
                result = self._eval_math(expr)
                if result:
                    r = f"  {expr.strip()} = {result}"
                    self.voice.speak(f"The answer is {result}")
                    return "info", r

        # natural language math: "what is 25 times 4"
        _spoken_math = re.search(
            r"(?:what(?:'?s|\s+is)\s+|calculate\s+|compute\s+)"
            r"(\d[\d\s]*(?:plus|minus|times|multiplied\s+by|divided\s+by|to\s+the\s+power\s+of|mod|percent\s+of)[\d\s]+\d)",
            tl_clean
        )
        if _spoken_math:
            expr = _spoken_math.group(1)
            expr = expr.replace("plus","+").replace("minus","-") \
                       .replace("times","*").replace("multiplied by","*") \
                       .replace("divided by","/").replace("to the power of","**") \
                       .replace("mod","%")
            # handle "X percent of Y"
            pm = re.match(r"([\d.]+)\s*percent\s*(?:of\s*)?([\d.]+)", expr)
            if pm:
                expr = f"({pm.group(1)}/100)*{pm.group(2)}"
            result = self._eval_math(expr)
            if result:
                r = f"  {_spoken_math.group(1).strip()} = {result}"
                self.voice.speak(f"The answer is {result}")
                return "info", r

        # ── 11. SYSTEM INFO ───────────────────────────────────────────
        _sys_kw = ["system info","system status","pc status","computer status",
                   "how's my pc","how is my pc","show me system","hardware info",
                   "performance","what's my cpu","check system"]
        if any(k in tl_clean for k in _sys_kw):
            r = "System Status\n"+"─"*40+"\n"+get_system_info()
            return "info", r

        _cpu_kw = ["cpu","processor","cpu usage","cpu load"]
        if any(k == tl_clean or k in tl_clean for k in _cpu_kw):
            return "info", "System Status\n"+"─"*40+"\n"+get_system_info()

        _ram_kw = ["ram","memory","memory usage","ram usage"]
        if any(k == tl_clean or k in tl_clean for k in _ram_kw):
            return "info", "System Status\n"+"─"*40+"\n"+get_system_info()

        _bat_kw = ["battery","battery level","battery status","power"]
        if any(k in tl_clean for k in _bat_kw) and not any(x in tl_clean for x in ["play","vlc","media","power shell","powershell"]):
            return "info", "System Status\n"+"─"*40+"\n"+get_system_info()

        if "disk" in tl_clean and any(k in tl_clean for k in ["usage","space","free","info","storage","how much"]):
            drive = "C:"
            m2 = re.search(r"\b([c-z]):\b", tl_clean)
            if m2: drive = m2.group(1).upper()+":"
            ok, msg = self.fm.disk_usage(drive)
            return ("info" if ok else "error"), msg

        # ── 12. YOUTUBE ───────────────────────────────────────────────
        if "youtube" in tl_clean:
            # Patterns: "play X on youtube", "youtube X", "search youtube for X", "open youtube"
            q = None
            for pat in [
                r"(?:play|watch|search|find|search\s+for|look\s+up)\s+(.+?)\s+(?:on\s+|in\s+)?youtube",
                r"youtube\s+(?:play|watch|search\s+for|search|find)?\s*(.+)",
                r"(.+?)\s+on\s+youtube",
            ]:
                m2 = re.search(pat, tl_clean)
                if m2:
                    candidate = m2.group(1).strip()
                    if candidate and candidate not in ("","open","start","please","now"):
                        q = candidate; break
            if q:
                return self.media.youtube(q)
            return self.media.youtube()

        # ── 13. SPOTIFY ───────────────────────────────────────────────
        if "spotify" in tl_clean:
            q = None
            for pat in [
                r"(?:play|listen\s+to|search\s+for|find)\s+(.+?)\s+(?:on\s+)?spotify",
                r"spotify\s+(?:play|search\s+for|search|find|open)?\s*(.+)",
            ]:
                m2 = re.search(pat, tl_clean)
                if m2:
                    candidate = m2.group(1).strip()
                    if candidate and candidate not in ("","open","start","please"):
                        q = candidate; break
            if q:
                return self.media.spotify(q)
            return self.media.spotify()

        # ── 14. VLC ───────────────────────────────────────────────────
        if "vlc" in tl_clean:
            m2 = re.search(r"vlc\s+(.+)", tl_clean)
            return self.media.vlc_play(m2.group(1).strip() if m2 else None)

        # ── 15. WINDOWS MEDIA PLAYER ──────────────────────────────────
        if any(k in tl_clean for k in ["windows media player","media player","wmp"]):
            m2 = re.search(r"(?:windows\s+media\s+player|media\s+player|wmp)\s+(.+)", tl_clean)
            return self.media.wmp_play(m2.group(1).strip() if m2 else None)

        # ── 16. PLAY MUSIC / SMART PLAY ───────────────────────────────
        m = re.match(r"(?:play|listen\s+to|put\s+on|start\s+playing)\s+(.+)", tl_clean)
        if m:
            q = m.group(1).strip("'\"")
            # if query has "on youtube", redirect
            if "on youtube" in q:
                return self.media.youtube(q.replace("on youtube","").strip())
            if "on spotify" in q:
                return self.media.spotify(q.replace("on spotify","").strip())
            return self.media.smart_play(q)

        # ── 17. WATCH / VIDEO ─────────────────────────────────────────
        m = re.match(r"(?:watch|stream)\s+(.+)", tl_clean)
        if m:
            q = m.group(1).strip()
            if "on youtube" in q: q = q.replace("on youtube","").strip()
            return self.media.youtube(q)

        # ── 18. FIND MUSIC / VIDEO ────────────────────────────────────
        m = re.match(r"find\s+(?:music|song|songs|audio|track)\s+(?:called\s+|named\s+|by\s+)?(.+)", tl_clean)
        if m: return self.media.find_music(m.group(1).strip())

        m = re.match(r"find\s+(?:video|movie|film|videos|movies)\s+(?:called\s+|named\s+|by\s+)?(.+)", tl_clean)
        if m: return self.media.find_video(m.group(1).strip())

        _music_list_kw = ["list music","show music","my music","music library",
                          "what music","show my songs","list songs","all music",
                          "music collection","browse music"]
        if any(k in tl_clean for k in _music_list_kw):
            return self.media.list_music()

        _video_list_kw = ["list videos","show videos","my videos","video library",
                          "what videos","show my videos","all videos","browse videos",
                          "video collection"]
        if any(k in tl_clean for k in _video_list_kw):
            return self.media.list_videos()

        # ── 19. FILE OPERATIONS ───────────────────────────────────────
        # Read / View file
        m = re.match(r"(?:read|show|cat|view|print|open|display)\s+(?:the\s+)?(?:file\s+|contents?\s+of\s+)?(.+)", tl_clean)
        if m and not any(k in tl_clean for k in ["folder","app","music","video","youtube","spotify","calculator","notepad"]):
            target = m.group(1).strip("'\"")
            # Make sure it looks like a filename
            if "." in target or any(w in target for w in ["txt","log","py","json","csv","md","html","xml"]):
                ok, msg = self.fm.read_file(target)
                return ("info" if ok else "error"), msg

        # Create file
        m = re.match(
            r"(?:create|make|new)\s+(?:a\s+)?(?:new\s+)?(?:text\s+)?file\s+"
            r"(?:called\s+|named\s+|with\s+name\s+)?(.+?)(?:\s+with\s+(?:content\s+|text\s+)?['\"]?(.+?)['\"]?)?$",
            tl_clean
        )
        if m:
            ok, msg = self.fm.create_file(m.group(1).strip(), m.group(2) or "")
            return ("action" if ok else "error"), msg

        # Create folder
        m = re.match(
            r"(?:create|make|new)\s+(?:a\s+)?(?:new\s+)?(?:folder|directory|dir)\s+"
            r"(?:called\s+|named\s+|with\s+name\s+)?(.+)",
            tl_clean
        )
        if m:
            ok, msg = self.fm.create_folder(m.group(1).strip())
            return ("action" if ok else "error"), msg

        # Delete file/folder
        m = re.match(r"(?:delete|remove|trash|erase|get\s+rid\s+of)\s+(?:the\s+)?(?:file\s+|folder\s+)?(.+)", tl_clean)
        if m and not any(k in tl_clean for k in ["youtube","spotify","app"]):
            ok, msg = self.fm.delete_path(m.group(1).strip())
            return ("action" if ok else "error"), msg

        # Rename
        m = re.match(r"(?:rename|change\s+name\s+of)\s+(.+?)\s+(?:to|as)\s+(.+)", tl_clean)
        if m:
            ok, msg = self.fm.rename_path(m.group(1).strip(), m.group(2).strip())
            return ("action" if ok else "error"), msg

        # Copy file
        m = re.match(r"copy\s+(?:the\s+)?(?:file\s+)?(.+?)\s+to\s+(.+)", tl_clean)
        if m and "clipboard" not in tl_clean:
            ok, msg = self.fm.copy_path(m.group(1).strip(), m.group(2).strip())
            return ("action" if ok else "error"), msg

        # Move file
        m = re.match(r"(?:move|mv|transfer)\s+(?:the\s+)?(?:file\s+)?(.+?)\s+to\s+(.+)", tl_clean)
        if m:
            ok, msg = self.fm.move_path(m.group(1).strip(), m.group(2).strip())
            return ("action" if ok else "error"), msg

        # Write to file
        m = re.match(r"(?:write|save)\s+['\"]?(.+?)['\"]?\s+to\s+(?:file\s+)?(.+)", tl_clean)
        if m:
            ok, msg = self.fm.write_file(m.group(2), m.group(1)+"\n")
            return ("action" if ok else "error"), msg

        # Append to file
        m = re.match(r"append\s+['\"]?(.+?)['\"]?\s+to\s+(?:file\s+)?(.+)", tl_clean)
        if m:
            ok, msg = self.fm.write_file(m.group(2), m.group(1)+"\n", append=True)
            return ("action" if ok else "error"), msg

        # File info / properties
        m = re.match(r"(?:info|details|properties|stat|size)\s+(?:of\s+|for\s+|about\s+)?(.+)", tl_clean)
        if m and not any(k in tl_clean for k in ["system","cpu","ram","disk"]):
            ok, msg = self.fm.file_info(m.group(1))
            return ("info" if ok else "error"), msg

        # List folder
        m = re.match(r"(?:list|ls|dir|show|contents?\s+of)\s+(?:folder\s+|directory\s+|the\s+)?(.+)", tl_clean)
        if m and not any(k in tl_clean for k in ["app","music","video","song","clipboard","apps"]):
            ok, msg = self.fm.list_folder(m.group(1).strip())
            return ("info" if ok else "error"), msg

        # Tree
        m = re.match(r"(?:tree|folder\s+tree|show\s+tree)\s+(?:of\s+)?(.+)", tl_clean)
        if m:
            ok, msg = self.fm.tree_folder(m.group(1).strip())
            return ("info" if ok else "error"), msg

        # Open folder in Explorer — handles BOTH word orders:
        #   "open folder STARK"  AND  "open STARK folder"
        m = re.match(r"(?:open|browse|explore)\s+(?:the\s+)?(?:folder|directory)\s+(.+)", tl_clean)
        if not m:
            m = re.match(r"(?:open|browse|explore)\s+(.+?)\s+(?:folder|directory)$", tl_clean)
        if m:
            ok, msg = self.fm.open_in_explorer(m.group(1).strip())
            return ("action" if ok else "error"), msg

        # Search files with location
        m = re.match(r"(?:find|search)\s+(.+?)\s+(?:in|inside|within)\s+(.+)", tl_clean)
        if m and not any(k in tl_clean for k in ["youtube","web","google","music","video"]):
            pat = m.group(1).strip()
            if "*" not in pat and "." not in pat: pat = f"*{pat}*"
            ok, msg = self.fm.search_files(pat, m.group(2).strip())
            return ("info" if ok else "error"), msg

        # Search files globally
        m = re.match(r"(?:find|search\s+for|look\s+for)\s+(?:the\s+)?(?:file\s+)?(.+)", tl_clean)
        if m and not any(k in tl_clean for k in ["youtube","google","web","music","video","song","movie","app"]):
            pat = m.group(1).strip()
            if "*" not in pat and "." not in pat: pat = f"*{pat}*"
            ok, msg = self.fm.search_files(pat, "home")
            return ("info" if ok else "error"), msg

        # ── 20. OPEN / LAUNCH APPS ────────────────────────────────────
        # Strip trailing noise words so "open Spotify app" → "spotify",
        # "open STARK folder" already handled above; strip here for safety.
        _open_clean = re.sub(r'\s+(app|application)$', '', tl_clean).strip()
        m = re.match(r"(?:open|launch|start|run|load|execute|bring\s+up)\s+(.+)", _open_clean)
        if m:
            name = m.group(1).strip("'\"").strip()

            # YouTube has no Windows app — always browser
            if re.search(r"\byoutube\b", name):
                return self.media.youtube()

            # Spotify — use dedicated media handler for proper play support
            if re.search(r"\bspotify\b", name):
                return self.media.spotify()

            # Look up in app registry (covers 50+ apps + Start Menu scan)
            app_path = self.apps.find(name)

            if app_path:
                ok, msg = launch_app(app_path, name)
                if ok:
                    self.voice.speak(f"Opening {name}")
                return ("action" if ok else "error"), msg

            # Not in registry — try smart universal finder live
            found = smart_find_app(name)
            if found:
                self.apps.apps[name.lower()] = found   # cache it for next time
                ok, msg = launch_app(found, name)
                if ok:
                    self.voice.speak(f"Opening {name}")
                return ("action" if ok else "error"), msg

            # Try as a file path
            ok, msg = self.fm.open_file(name)
            if ok:
                return "action", msg

            return "error", (
                f"I couldn't find '{name}' on your system.\n"
                f"  → Make sure it's installed and try again.\n"
                f"  → Or say 'open [full path]' to launch it directly."
            )

        # ── 21. WEB SEARCH ────────────────────────────────────────────
        _search_prefixes = [
            "search for","search the web for","google","look up","look for",
            "search","find me information about","find info on","browse for",
            "can you search","search online for","search google for",
        ]
        q = self._strip_prefix(tl_clean, _search_prefixes)
        if q and not any(k in tl_clean for k in ["file","folder","music","video","app"]):
            webbrowser.open(f"https://www.google.com/search?q={urlquote(q)}")
            r = f"Searching Google for: \"{q}\" 🔍"
            self.voice.speak(f"Searching for {q}")
            return "action", r

        # ── 22. WIKIPEDIA ─────────────────────────────────────────────
        m = re.match(r"(?:wikipedia|wiki|wikepedia)\s+(.+)", tl_clean)
        if not m: m = re.match(r"(?:what is|who is|tell me about|explain)\s+(.+)", tl_clean)
        if m:
            topic = m.group(1).strip()
            # Only handle simple factual lookups here; complex questions → AI
            if len(topic.split()) <= 5 and not "?" in topic:
                webbrowser.open(f"https://en.wikipedia.org/wiki/{urlquote(topic)}")
                r = f"Opening Wikipedia: {topic.title()} 📖"
                self.voice.speak(f"Looking up {topic} on Wikipedia")
                return "action", r

        # ── 23. CLIPBOARD ─────────────────────────────────────────────
        _clip_kw = ["read clipboard","show clipboard","what's in clipboard",
                    "what is in clipboard","clipboard contents","get clipboard",
                    "paste","clipboard"]
        if any(k in tl_clean for k in _clip_kw) and "copy" not in tl_clean:
            if not HAS_CLIP: return "error", "pyperclip not installed — run: pip install pyperclip"
            try:
                clip = pyperclip.paste()
                return "info", f"Clipboard:\n{'─'*40}\n{clip or '(empty)'}"
            except Exception as e: return "error", str(e)

        m = re.match(r"copy\s+['\"]?(.+?)['\"]?\s+to\s+clipboard", tl_clean)
        if m:
            if not HAS_CLIP: return "error", "pyperclip not installed"
            try:
                pyperclip.copy(m.group(1))
                return "action", f"Copied to clipboard: \"{m.group(1)}\""
            except Exception as e: return "error", str(e)

        # ── 24. LIST APPS ─────────────────────────────────────────────
        _apps_kw = ["list apps","show apps","what apps","list applications",
                    "available apps","what can you open","installed apps",
                    "what applications","my apps"]
        if any(k in tl_clean for k in _apps_kw):
            apps = self.apps.list_all()[:50]
            return "info", "Installed Applications:\n\n" + "\n".join(f"  • {a.title()}" for a in apps)

        # ── 25. HELP / COMMANDS ───────────────────────────────────────
        _help_kw = ["help","commands","what can i say","show commands","guide",
                    "how do i","instructions","tutorial","tips"]
        if any(k in tl_clean for k in _help_kw):
            r = (
                "STARK  —  Quick Command Guide\n"
                "═"*44+"\n\n"
                "🗓  TIME & DATE\n"
                "   What time is it  |  What's today's date\n\n"
                "💻  SYSTEM\n"
                "   System info  |  CPU  |  RAM  |  Battery  |  Disk usage\n\n"
                "🧮  MATH\n"
                "   Calculate 25 * 4  |  What is 100 / 5  |  Sqrt 144\n\n"
                "🎵  MEDIA\n"
                "   Play [song]  |  YouTube [query]  |  Spotify [artist]\n"
                "   List music  |  List videos  |  VLC [file]\n\n"
                "📁  FILES\n"
                "   Create file notes.txt  |  Delete file test.txt\n"
                "   Rename old.txt to new.txt  |  Read file readme.txt\n"
                "   Copy file a.txt to desktop  |  Find *.pdf in documents\n\n"
                "🖥  APPS\n"
                "   Open Notepad  |  Open Chrome  |  Launch Calculator\n\n"
                "🌐  WEB\n"
                "   Search for Python tutorials  |  Google best laptops\n"
                "   Wikipedia Albert Einstein\n\n"
                "🗣  VOICE\n"
                "   Voice on/off  |  Set speed fast/slow  |  Set volume 80\n"
                "   Say hello world  |  Stop speaking\n\n"
                "💬  CHAT\n"
                "   Anything else goes to your local AI model!"
            )
            return "info", r

        # ── 26. SCREENSHOT ────────────────────────────────────────────
        _ss_kw = ["screenshot","take a screenshot","capture screen","screen capture",
                  "snip","snipping tool"]
        if any(k in tl_clean for k in _ss_kw):
            try:
                subprocess.Popen(["snippingtool.exe"])
                return "action", "Opening Snipping Tool — select your area."
            except Exception:
                try:
                    subprocess.Popen(["SnippingTool.exe"])
                    return "action", "Opening Snipping Tool."
                except:
                    return "error", "Could not open Snipping Tool."

        # ── 27. LOCK / SLEEP / SHUTDOWN ───────────────────────────────
        if any(k in tl_clean for k in ["lock pc","lock computer","lock screen","lock my pc"]):
            subprocess.Popen("rundll32.exe user32.dll,LockWorkStation", shell=True)
            return "action", "PC locked. See you later!"

        if any(k in tl_clean for k in ["sleep","put to sleep","sleep mode"]):
            subprocess.Popen("rundll32.exe powrprof.dll,SetSuspendState 0,1,0", shell=True)
            return "action", "Going to sleep. Goodnight!"

        if any(k in tl_clean for k in ["restart","reboot","restart pc"]):
            r = "Restarting PC in 10 seconds. Close your work!"
            self.voice.speak(r)
            subprocess.Popen("shutdown /r /t 10", shell=True)
            return "action", r

        if any(k in tl_clean for k in ["shutdown","turn off","power off","shut down pc","shut down computer"]):
            r = "Shutting down PC in 10 seconds. Save your work!"
            self.voice.speak(r)
            subprocess.Popen("shutdown /s /t 10", shell=True)
            return "action", r

        if any(k in tl_clean for k in ["cancel shutdown","abort shutdown","cancel restart","stop shutdown"]):
            subprocess.Popen("shutdown /a", shell=True)
            return "action", "Shutdown/restart cancelled. You're safe."

        # ── 28. VOLUME / BRIGHTNESS (system) ─────────────────────────
        if any(k in tl_clean for k in ["mute system","mute sound","mute pc","mute audio"]):
            # Send mute key via nircmd if available, fallback message
            try:
                subprocess.Popen(["nircmd.exe","mutesysvolume","2"])
                return "action", "System audio toggled."
            except:
                return "info", "Tip: Press the mute key on your keyboard, or use the taskbar volume icon."

        # ── 29. OPEN WEBSITE ──────────────────────────────────────────
        m = re.match(r"(?:go\s+to|navigate\s+to|visit)\s+(?:the\s+)?(?:website\s+)?(.+)", tl_clean)
        if m:
            target = m.group(1).strip()
            # Only handle actual URLs here
            if re.match(r'[\w\-]+\.(com|org|net|io|edu|gov|co|uk|in|ai)', target):
                url = target if target.startswith("http") else f"https://{target}"
                open_in_chrome(url)
                return "action", f"Opening {target} in browser."

        # ── 30. WEATHER (graceful fallback) ───────────────────────────
        _weather_kw = ["weather","forecast","temperature","rain","sunny","cloudy",
                       "will it rain","how hot","how cold","weather today"]
        if any(k in tl_clean for k in _weather_kw):
            webbrowser.open("https://www.google.com/search?q=weather+today")
            return "action", "Opened Google Weather for you. 🌤"

        # ── 31. NEWS ──────────────────────────────────────────────────
        _news_kw = ["news","latest news","today's news","current news","headlines"]
        if any(k in tl_clean for k in _news_kw):
            webbrowser.open("https://news.google.com")
            return "action", "Opening Google News. 📰"

        # ── 32. MAPS ──────────────────────────────────────────────────
        m = re.match(r"(?:directions?\s+to|navigate\s+to|maps?\s+to|how\s+to\s+get\s+to)\s+(.+)", tl_clean)
        if m:
            dest = urlquote(m.group(1).strip())
            open_in_chrome(f"https://www.google.com/maps/dir/?api=1&destination={dest}")
            return "action", f"Opening Google Maps directions to {m.group(1).strip()}."

        m = re.match(r"(?:find\s+on\s+map|map|show\s+me)\s+(.+)", tl_clean)
        if m and "file" not in tl_clean:
            dest = urlquote(m.group(1).strip())
            open_in_chrome(f"https://www.google.com/maps/search/{dest}")
            return "action", f"Opening Google Maps: {m.group(1).strip()}."

        # ── 33. TRANSLATE ─────────────────────────────────────────────
        m = re.match(r"(?:translate|translation\s+of)\s+(.+?)(?:\s+(?:to|in|into)\s+(\w+))?$", tl_clean)
        if m:
            phrase = urlquote(m.group(1).strip())
            lang   = m.group(2) or ""
            url = f"https://translate.google.com/?sl=auto&tl={lang}&text={phrase}&op=translate" if lang \
                  else f"https://translate.google.com/?sl=auto&text={phrase}"
            open_in_chrome(url)
            return "action", f"Opening Google Translate."

        # ── Fallback → AI ─────────────────────────────────────────────
        return None


# ════════════════════════════════════════════════════════════════════
#  FLOATING ORB WIDGET
# ════════════════════════════════════════════════════════════════════
class FloatingOrb(tk.Toplevel):
    SIZE = 76

    def __init__(self, root, on_click):
        super().__init__(root)
        self.on_click    = on_click
        self._phase      = 0.0
        self._state      = "idle"
        self._panel_open = False
        self._moved      = False
        self._dx = self._dy = 0

        self.overrideredirect(True)
        self.attributes("-topmost", True)
        self.attributes("-transparentcolor", "#010101")
        self.configure(bg="#010101")
        self.resizable(False, False)

        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"{self.SIZE}x{self.SIZE}+{sw-100}+{sh-150}")

        self.cv = tk.Canvas(self, width=self.SIZE, height=self.SIZE,
                            bg="#010101", highlightthickness=0)
        self.cv.pack()
        self.cv.bind("<ButtonPress-1>",   self._d_start)
        self.cv.bind("<B1-Motion>",       self._d_move)
        self.cv.bind("<ButtonRelease-1>", self._d_end)

        self._draw(); self._tick()

    def set_state(self, s): self._state = s
    def set_open(self, v):  self._panel_open = v

    def _draw(self):
        cv  = self.cv; cv.delete("all")
        cx  = cy = self.SIZE // 2
        phi = self._phase
        n   = 16
        color = {
            "idle":     C_ARC,
            "thinking": C_GOLD,
            "speaking": C_GREEN,
            "busy":     C_AMBER,
        }.get(self._state, C_ARC)

        cv.create_oval(3, 3, self.SIZE-3, self.SIZE-3,
                       fill="#06101e", outline=C_BORDER, width=1)

        for i in range(n):
            angle = (2*math.pi/n)*i
            wave  = (math.sin(phi + i*0.5) + 1)/2

            if self._state == "idle":
                amp = 1.5 + wave*2.5
            elif self._state == "thinking":
                amp = 2 + abs(math.sin(phi*3 + i*0.8))*7
            elif self._state == "speaking":
                amp = 1 + wave*9
            else:
                amp = 2 + wave*4

            r_in  = 26
            r_out = 26 + amp
            x1 = cx + r_in  * math.cos(angle)
            y1 = cy + r_in  * math.sin(angle)
            x2 = cx + r_out * math.cos(angle)
            y2 = cy + r_out * math.sin(angle)
            cv.create_line(x1,y1,x2,y2, fill=color,
                           width=2.5 if self._state != "idle" else 2,
                           capstyle=tk.ROUND)

        ri = 22
        cv.create_oval(cx-ri, cy-ri, cx+ri, cy+ri,
                       fill="#060f1c", outline=color, width=1)

        pulse = 0.88 + 0.12*math.sin(phi*1.2)
        fsz   = max(9, int(15*pulse))
        cv.create_text(cx, cy, text="S",
                       font=("Segoe UI", fsz, "bold"), fill=color)

        if self._panel_open:
            cv.create_oval(cx+13, cy-19, cx+19, cy-13, fill=C_GOLD, outline="")

    def _tick(self):
        speeds = {"idle":0.035,"thinking":0.13,"speaking":0.08,"busy":0.1}
        self._phase += speeds.get(self._state, 0.035)
        self._draw()
        self.after(38, self._tick)

    def _d_start(self, e):
        self._dx = e.x_root - self.winfo_x()
        self._dy = e.y_root - self.winfo_y()
        self._moved = False

    def _d_move(self, e):
        self.geometry(f"+{e.x_root-self._dx}+{e.y_root-self._dy}")
        self._moved = True

    def _d_end(self, e):
        if not self._moved: self.on_click()


# ════════════════════════════════════════════════════════════════════
#  HELPER DIALOG
# ════════════════════════════════════════════════════════════════════
def ask_input(parent, title, prompt):
    dlg = tk.Toplevel(parent)
    dlg.title(title); dlg.geometry("380x110")
    dlg.configure(bg=C_PANEL); dlg.grab_set(); dlg.resizable(False,False)
    tk.Label(dlg, text=prompt, font=FONT_UI, bg=C_PANEL, fg=C_WHITE).pack(pady=(14,4), padx=16)
    var   = tk.StringVar()
    entry = tk.Entry(dlg, textvariable=var, font=FONT_MONO, bg=C_CARD, fg=C_WHITE,
                     relief="flat", insertbackground=C_ARC, bd=0,
                     highlightthickness=1, highlightcolor=C_ARC, highlightbackground=C_BORDER)
    entry.pack(fill="x", padx=16, ipady=6); entry.focus()
    result = [None]
    def ok(e=None): result[0]=var.get().strip(); dlg.destroy()
    entry.bind("<Return>", ok)
    tk.Button(dlg, text="OK", font=FONT_BTN, bg=C_ARC, fg=C_BG,
              relief="flat", padx=16, command=ok).pack(pady=8)
    dlg.wait_window()
    return result[0] or None


# ════════════════════════════════════════════════════════════════════
#  FILE BROWSER PANEL
# ════════════════════════════════════════════════════════════════════
class FileBrowserPanel(tk.Toplevel):
    def __init__(self, parent, fm: FileManager):
        super().__init__(parent)
        self.fm = fm
        self.title("File Browser"); self.geometry("740x520")
        self.configure(bg=C_BG); self.resizable(True,True)
        self._items = []; self.cur = DESKTOP
        self._build(); self._load(self.cur); self.grab_set()

    def _build(self):
        top = tk.Frame(self, bg=C_PANEL, pady=6, padx=8); top.pack(fill="x")
        for lbl, path in [("🖥 Desktop",DESKTOP),("📄 Docs",DOCUMENTS),
                           ("⬇ Downloads",DOWNLOADS),("🎵 Music",MUSIC_DIR),
                           ("🎬 Videos",VIDEOS_DIR),("🏠 Home",HOME)]:
            tk.Button(top, text=lbl, font=FONT_MONO_S, bg=C_CARD, fg=C_MUTED,
                      relief="flat", padx=6, cursor="hand2",
                      command=lambda p=path: self._load(p)).pack(side="left", padx=2)
        tk.Button(top, text="⬆ Up", font=FONT_MONO_S, bg=C_CARD, fg=C_YELLOW,
                  relief="flat", padx=6, cursor="hand2",
                  command=self._up).pack(side="left", padx=(10,2))
        self.pvar = tk.StringVar()
        pe = tk.Entry(self, textvariable=self.pvar, font=FONT_MONO_S,
                      bg=C_CARD, fg=C_ARC, relief="flat", insertbackground=C_ARC, bd=0,
                      highlightthickness=1, highlightbackground=C_BORDER)
        pe.pack(fill="x", padx=8, pady=(4,0), ipady=5)
        pe.bind("<Return>", lambda e: self._load(Path(self.pvar.get())))
        lf = tk.Frame(self, bg=C_BG); lf.pack(fill="both", expand=True, padx=8, pady=8)
        sb = tk.Scrollbar(lf, bg=C_CARD, troughcolor=C_PANEL, bd=0); sb.pack(side="right", fill="y")
        self.lb = tk.Listbox(lf, font=FONT_MONO_S, bg=C_CARD, fg=C_WHITE,
                             relief="flat", selectbackground=C_ARC2,
                             selectforeground="#fff", yscrollcommand=sb.set,
                             activestyle="none", highlightthickness=0, bd=0)
        self.lb.pack(fill="both", expand=True); sb.config(command=self.lb.yview)
        self.lb.bind("<Double-Button-1>", self._dbl)
        bot = tk.Frame(self, bg=C_PANEL, pady=6, padx=8); bot.pack(fill="x")
        self.slbl = tk.StringVar()
        tk.Label(bot, textvariable=self.slbl, font=FONT_MONO_S, bg=C_PANEL, fg=C_MUTED).pack(side="left")
        for txt,fg,cmd in [("Open",C_ARC,self._open),("New File",C_GREEN,self._new_file),
                           ("New Folder",C_YELLOW,self._new_folder),
                           ("Rename",C_AMBER,self._rename),("Delete",C_RED,self._delete)]:
            tk.Button(bot, text=txt, font=FONT_BTN, bg=C_CARD, fg=fg,
                      relief="flat", padx=8, cursor="hand2", command=cmd).pack(side="right", padx=2)

    def _load(self, p):
        p = Path(p)
        if not p.exists(): self.slbl.set(f"Not found: {p}"); return
        self.cur = p; self.pvar.set(str(p))
        self.lb.delete(0,"end"); self._items=[]
        try: items = sorted(p.iterdir(), key=lambda x:(x.is_file(),x.name.lower()))
        except PermissionError: self.slbl.set("Access denied"); return
        for item in items:
            if item.name.startswith("."): continue
            icon = "📁  " if item.is_dir() else "📄  "
            try: sz = fmt_size(item.stat().st_size) if item.is_file() else ""
            except: sz=""
            self.lb.insert("end", f"{icon}{item.name:<50} {sz}")
            self._items.append(item)
        self.slbl.set(f"{len(self._items)} items in {p.name}/")

    def _up(self):
        p = self.cur.parent
        if p != self.cur: self._load(p)
    def _sel(self):
        s = self.lb.curselection()
        return self._items[s[0]] if s else None
    def _dbl(self, e=None):
        item = self._sel()
        if item:
            if item.is_dir(): self._load(item)
            else:
                try: os.startfile(str(item))
                except: pass
    def _open(self):
        item = self._sel()
        if item:
            if item.is_dir(): subprocess.Popen(f'explorer "{item}"')
            else:
                try: os.startfile(str(item))
                except: pass
    def _new_file(self):
        n = ask_input(self, "New File", "Enter file name:")
        if n:
            ok,msg = self.fm.create_file(str(self.cur/n))
            messagebox.showinfo("Result",msg,parent=self); self._load(self.cur)
    def _new_folder(self):
        n = ask_input(self, "New Folder", "Enter folder name:")
        if n:
            ok,msg = self.fm.create_folder(str(self.cur/n))
            messagebox.showinfo("Result",msg,parent=self); self._load(self.cur)
    def _rename(self):
        item = self._sel()
        if not item: return
        n = ask_input(self, "Rename", f"Rename '{item.name}' to:")
        if n:
            ok,msg = self.fm.rename_path(str(item),n)
            messagebox.showinfo("Result",msg,parent=self); self._load(self.cur)
    def _delete(self):
        item = self._sel()
        if not item: return
        if messagebox.askyesno("Delete",f"Delete '{item.name}'?",parent=self):
            ok,msg = self.fm.delete_path(str(item))
            messagebox.showinfo("Result",msg,parent=self); self._load(self.cur)


# ════════════════════════════════════════════════════════════════════
#  MEDIA BROWSER PANEL
# ════════════════════════════════════════════════════════════════════
class MediaBrowserPanel(tk.Toplevel):
    def __init__(self, parent, media: MediaPlugin):
        super().__init__(parent)
        self.media = media
        self.title("Media Browser"); self.geometry("760x540")
        self.configure(bg=C_BG); self.resizable(True,True)
        self._all = []; self._visible = []
        self._build(); self._load_all(); self.grab_set()

    def _build(self):
        top = tk.Frame(self, bg=C_PANEL, pady=6, padx=8); top.pack(fill="x")
        for lbl, fn in [("🎵 Music",lambda:self._load_dir(MUSIC_DIR,AUDIO_EXTS)),
                        ("🎬 Videos",lambda:self._load_dir(VIDEOS_DIR,VIDEO_EXTS)),
                        ("⬇ Downloads",lambda:self._load_dir(DOWNLOADS)),
                        ("🖥 Desktop",lambda:self._load_dir(DESKTOP)),
                        ("📚 All Media",self._load_all)]:
            tk.Button(top, text=lbl, font=FONT_MONO_S, bg=C_CARD, fg=C_MUTED,
                      relief="flat", padx=7, cursor="hand2",
                      command=fn).pack(side="left", padx=2)
        sf = tk.Frame(self, bg=C_BG, pady=4); sf.pack(fill="x", padx=8)
        self.sv = tk.StringVar(); self.sv.trace("w", lambda *a: self._filter())
        tk.Label(sf, text="🔎", bg=C_BG, fg=C_MUTED).pack(side="left", padx=(0,4))
        tk.Entry(sf, textvariable=self.sv, font=FONT_MONO_S, bg=C_CARD, fg=C_WHITE,
                 relief="flat", insertbackground=C_ARC, bd=0,
                 highlightthickness=1, highlightcolor=C_ARC, highlightbackground=C_BORDER
                 ).pack(side="left", fill="x", expand=True, ipady=5)
        lf = tk.Frame(self, bg=C_BG); lf.pack(fill="both", expand=True, padx=8, pady=4)
        sb = tk.Scrollbar(lf, bg=C_CARD, troughcolor=C_PANEL, bd=0); sb.pack(side="right", fill="y")
        self.lb = tk.Listbox(lf, font=FONT_MONO_S, bg=C_CARD, fg=C_WHITE,
                             relief="flat", selectbackground=C_ARC2, selectforeground="#fff",
                             yscrollcommand=sb.set, activestyle="none",
                             highlightthickness=0, bd=0)
        self.lb.pack(fill="both", expand=True); sb.config(command=self.lb.yview)
        self.lb.bind("<Double-Button-1>", lambda e: self._play_default())
        bot = tk.Frame(self, bg=C_PANEL, pady=6, padx=8); bot.pack(fill="x")
        self.slbl = tk.StringVar()
        tk.Label(bot, textvariable=self.slbl, font=FONT_MONO_S, bg=C_PANEL, fg=C_MUTED).pack(side="left")
        for txt,fg,cmd in [("▶ Play VLC",C_GREEN,self._play_vlc),
                           ("🎵 WMP",C_YELLOW,self._play_wmp),
                           ("▶ Default",C_ARC,self._play_default),
                           ("📂 Show",C_TEAL,self._show_folder)]:
            tk.Button(bot, text=txt, font=FONT_BTN, bg=C_CARD, fg=fg,
                      relief="flat", padx=8, cursor="hand2", command=cmd).pack(side="right",padx=2)

    def _load_all(self):
        files = []
        for r in MediaPlugin.SEARCH_ROOTS:
            files += MediaPlugin().list_local(r)
        self._all = files; self._render(files)

    def _load_dir(self, d, exts=None):
        files = MediaPlugin().list_local(d, exts)
        self._all = files; self._render(files)

    def _render(self, files):
        self.lb.delete(0,"end"); self._visible=files
        for p in files:
            icon = "🎵 " if p.suffix.lower() in AUDIO_EXTS else "🎬 "
            self.lb.insert("end", f"{icon} {p.name}")
        self.slbl.set(f"{len(files)} files")

    def _filter(self):
        q = self.sv.get().lower()
        self._render([p for p in self._all if q in p.name.lower()] if q else self._all)

    def _sel(self):
        s = self.lb.curselection()
        return self._visible[s[0]] if s else None

    def _play_default(self, e=None):
        p = self._sel()
        if p:
            try: os.startfile(str(p))
            except Exception as ex: messagebox.showerror("Error",str(ex),parent=self)

    def _play_vlc(self):
        p = self._sel()
        if p:
            exe = find_exe(VLC_PATHS)
            if exe: subprocess.Popen([exe, str(p)])
            else: messagebox.showwarning("VLC","VLC not found",parent=self)

    def _play_wmp(self):
        p = self._sel()
        if p:
            exe = find_exe(WMP_PATHS)
            if exe: subprocess.Popen([exe, str(p)])
            else:
                try: os.startfile(str(p))
                except: pass

    def _show_folder(self):
        p = self._sel()
        if p: subprocess.Popen(f'explorer /select,"{p}"')


# ════════════════════════════════════════════════════════════════════
#  VOICE SETTINGS PANEL
# ════════════════════════════════════════════════════════════════════
class VoiceSettingsPanel(tk.Toplevel):
    def __init__(self, parent, voice: VoiceEngine):
        super().__init__(parent)
        self.voice = voice
        self.title("Voice Settings"); self.geometry("400x340")
        self.configure(bg=C_BG); self.resizable(False,False); self.grab_set()
        self._build()

    def _build(self):
        tk.Label(self, text="VOICE OUTPUT", font=("Courier New",12,"bold"),
                 bg=C_BG, fg=C_ARC).pack(pady=(18,2))
        tk.Label(self, text="Powered by Windows SAPI5 — no internet needed",
                 font=FONT_SMALL, bg=C_BG, fg=C_MUTED).pack(pady=(0,14))

        card = tk.Frame(self, bg=C_CARD, padx=20, pady=16); card.pack(fill="x",padx=20)

        def row(lbl, widget_fn):
            f = tk.Frame(card, bg=C_CARD); f.pack(fill="x", pady=7)
            tk.Label(f, text=lbl, font=FONT_UI, bg=C_CARD, fg=C_MUTED, width=10, anchor="w").pack(side="left")
            widget_fn(f)

        self.tbtn = None
        def make_toggle(p):
            self.tbtn = tk.Button(p, font=FONT_BTN, relief="flat", padx=14, pady=3, cursor="hand2",
                                  command=self._toggle)
            self._update_tbtn()
            self.tbtn.pack(side="left")
        row("Voice:", make_toggle)

        vnames = self.voice.get_voice_names()
        self.vvar = tk.StringVar(value=vnames[0] if vnames else "Default")
        def make_voice(p):
            cb = ttk.Combobox(p, textvariable=self.vvar, values=vnames,
                              state="readonly", font=FONT_MONO_S, width=24)
            cb.pack(side="left")
            cb.bind("<<ComboboxSelected>>", lambda e: self.voice.set_voice_idx(vnames.index(self.vvar.get())))
        row("Voice:", make_voice)

        self.rvar = tk.IntVar(value=self.voice.rate)
        self.rlbl = None
        def make_rate(p):
            f2 = tk.Frame(p, bg=C_CARD); f2.pack(side="left")
            self.rlbl = tk.Label(f2, text=f"{self.voice.rate} wpm",
                                 font=FONT_MONO_S, bg=C_CARD, fg=C_ARC, width=8)
            self.rlbl.pack(side="right")
            tk.Scale(f2, from_=60, to=380, orient="horizontal", variable=self.rvar,
                     bg=C_CARD, fg=C_MUTED, troughcolor=C_BORDER, highlightthickness=0,
                     showvalue=False, length=160,
                     command=lambda v:(self.voice.set_rate(int(v)),
                                      self.rlbl.config(text=f"{int(v)} wpm"))).pack(side="left")
        row("Speed:", make_rate)

        self.volvr = tk.DoubleVar(value=self.voice.volume*100)
        self.vollbl = None
        def make_vol(p):
            f2 = tk.Frame(p, bg=C_CARD); f2.pack(side="left")
            self.vollbl = tk.Label(f2, text=f"{int(self.voice.volume*100)}%",
                                   font=FONT_MONO_S, bg=C_CARD, fg=C_ARC, width=5)
            self.vollbl.pack(side="right")
            tk.Scale(f2, from_=0, to=100, orient="horizontal", variable=self.volvr,
                     bg=C_CARD, fg=C_MUTED, troughcolor=C_BORDER, highlightthickness=0,
                     showvalue=False, length=160,
                     command=lambda v:(self.voice.set_volume(float(v)/100),
                                      self.vollbl.config(text=f"{int(float(v))}%"))).pack(side="left")
        row("Volume:", make_vol)

        pf = tk.Frame(card, bg=C_CARD); pf.pack(fill="x", pady=(10,0))
        tk.Label(pf, text="Preset:", font=FONT_UI, bg=C_CARD, fg=C_MUTED).pack(side="left",padx=(0,8))
        for lbl,r,v in [("Slow",120,0.9),("Normal",175,0.9),("Fast",260,0.9)]:
            tk.Button(pf, text=lbl, font=FONT_BTN, bg=C_BORDER, fg=C_MUTED,
                      relief="flat", padx=10, cursor="hand2",
                      command=lambda r=r,v=v:self._preset(r,v)).pack(side="left",padx=4)

        bf = tk.Frame(self, bg=C_BG); bf.pack(pady=16)
        tk.Button(bf, text="🔊 Test", font=FONT_BTN, bg=C_ARC, fg=C_BG,
                  relief="flat", padx=14, pady=5, cursor="hand2",
                  command=lambda:self.voice.speak("Hello. I am Stark, your personal AI assistant.", force=True)
                  ).pack(side="left", padx=5)
        tk.Button(bf, text="⏹ Stop", font=FONT_BTN, bg=C_CARD, fg=C_RED,
                  relief="flat", padx=14, pady=5, cursor="hand2",
                  command=self.voice.stop).pack(side="left", padx=5)
        tk.Button(bf, text="Close", font=FONT_BTN, bg=C_CARD, fg=C_MUTED,
                  relief="flat", padx=14, pady=5, cursor="hand2",
                  command=self.destroy).pack(side="left", padx=5)

        if not HAS_TTS:
            tk.Label(self, text="⚠  pyttsx3 not installed\npip install pyttsx3",
                     font=FONT_UI, bg=C_BG, fg=C_RED, justify="center").pack(pady=8)

    def _update_tbtn(self):
        if self.tbtn:
            on = self.voice.enabled
            self.tbtn.config(text="● ON" if on else "○ OFF",
                             bg=C_GREEN if on else C_RED,
                             fg="#000" if on else "#fff")

    def _toggle(self):
        self.voice.toggle(); self._update_tbtn()

    def _preset(self, r, v):
        self.voice.set_rate(r); self.voice.set_volume(v)
        self.rvar.set(r); self.volvr.set(v*100)
        if self.rlbl:  self.rlbl.config(text=f"{r} wpm")
        if self.vollbl: self.vollbl.config(text=f"{int(v*100)}%")


# ════════════════════════════════════════════════════════════════════
#  STARK MINI PANEL  (Siri-style popup)
#  FIX 1: _greet() now uses self._write() instead of self._add()
#  FIX 2: Input box is at the very BOTTOM (packed last)
# ════════════════════════════════════════════════════════════════════
class StarkMiniPanel(tk.Toplevel):
    W = 450
    H = 600

    def __init__(self, root, orb: FloatingOrb, parser: CommandParser,
                 ai: AIClient, voice: VoiceEngine,
                 fm: FileManager, media: MediaPlugin):
        super().__init__(root)
        self.orb    = orb
        self.parser = parser
        self.ai     = ai
        self.voice  = voice
        self.fm     = fm
        self.media  = media

        self.overrideredirect(True)
        self.attributes("-topmost", True)
        self.configure(bg=C_BG)
        self.resizable(True, True)

        self._busy   = False
        self._ai_buf = ""
        self._hist   = []
        self._hidx   = -1
        self._px = self._py = 0

        self._position()
        self._build()
        self._greet()
        self._status_loop()

    def _position(self):
        ox, oy = self.orb.winfo_x(), self.orb.winfo_y()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        x = ox - self.W - 10
        y = oy - self.H + self.orb.SIZE
        if x < 10: x = ox + self.orb.SIZE + 10
        if y < 10: y = 10
        if y + self.H > sh - 50: y = sh - self.H - 50
        self.geometry(f"{self.W}x{self.H}+{x}+{y}")

    def _build(self):
        # ── Title bar ────────────────────────────────────────────
        bar = tk.Frame(self, bg=C_PANEL, height=44)
        bar.pack(fill="x")
        bar.pack_propagate(False)
        bar.bind("<ButtonPress-1>",  lambda e: self._drag_start(e))
        bar.bind("<B1-Motion>",      lambda e: self._drag_move(e))

        tk.Label(bar, text="S", font=("Segoe UI",15,"bold"), bg=C_PANEL, fg=C_ARC
                 ).pack(side="left", padx=(12,4), pady=8)
        tk.Label(bar, text="STARK", font=("Courier New",11,"bold"), bg=C_PANEL, fg=C_WHITE
                 ).pack(side="left", pady=8)
        tk.Label(bar, text="  AI ASSISTANT", font=("Segoe UI",7), bg=C_PANEL, fg=C_MUTED
                 ).pack(side="left", pady=(18,0))

        self.dot = tk.Label(bar, text="●", font=("Segoe UI",9), bg=C_PANEL, fg=C_MUTED)
        self.dot.pack(side="right", padx=(0,4))
        self.dlbl = tk.Label(bar, text="offline", font=("Segoe UI",7), bg=C_PANEL, fg=C_MUTED)
        self.dlbl.pack(side="right", padx=(0,2))
        tk.Button(bar, text="✕", font=("Segoe UI",10), bg=C_PANEL, fg=C_MUTED,
                  relief="flat", padx=8, cursor="hand2",
                  activebackground=C_RED, activeforeground=C_WHITE,
                  command=self._close).pack(side="right", pady=2)

        # ── Quick action row ──────────────────────────────────────
        qf = tk.Frame(self, bg=C_CARD)
        qf.pack(fill="x")
        for lbl, cmd in [("⏱ Time","what time is it"),("📊 System","system info"),
                         ("▶ YouTube","youtube "),("🎵 Spotify","spotify"),
                         ("🔍 Search","search for "),("📁 Desktop","open folder desktop")]:
            tk.Button(qf, text=lbl, font=FONT_SMALL, bg=C_CARD, fg=C_MUTED,
                      relief="flat", padx=7, pady=4, cursor="hand2",
                      activeforeground=C_ARC,
                      command=lambda c=cmd: self._quick(c)).pack(side="left", padx=1, pady=3)

        # ── Media row ─────────────────────────────────────────────
        mf = tk.Frame(self, bg=C_CARD)
        mf.pack(fill="x")
        for lbl, cmd in [("🎵 Music","list music"),("🎬 Videos","list videos"),
                         ("🎵 Find","find music "),("🎬 Find","find video "),
                         ("▶ Play","play "),("📁 Files",""),("🎵 Browse","")]:
            if lbl == "📁 Files":
                tk.Button(mf, text=lbl, font=FONT_SMALL, bg=C_CARD, fg=C_TEAL,
                          relief="flat", padx=7, pady=4, cursor="hand2",
                          command=lambda: FileBrowserPanel(self, self.fm)).pack(side="left", padx=1, pady=3)
            elif lbl == "🎵 Browse":
                tk.Button(mf, text=lbl, font=FONT_SMALL, bg=C_CARD, fg=C_PINK,
                          relief="flat", padx=7, pady=4, cursor="hand2",
                          command=lambda: MediaBrowserPanel(self, self.media)).pack(side="left", padx=1, pady=3)
            else:
                tk.Button(mf, text=lbl, font=FONT_SMALL, bg=C_CARD, fg=C_PINK,
                          relief="flat", padx=7, pady=4, cursor="hand2",
                          activeforeground=C_PURPLE,
                          command=lambda c=cmd: self._quick(c)).pack(side="left", padx=1, pady=3)

        tk.Frame(self, bg=C_BORDER, height=1).pack(fill="x")

        # ── Input bar — pack(side="bottom") so it anchors to the very bottom ──
        inf = tk.Frame(self, bg=C_PANEL, pady=8, padx=10)
        inf.pack(side="bottom", fill="x")

        tk.Button(inf, text="⊞", font=("Segoe UI",11), bg=C_CARD, fg=C_MUTED,
                  relief="flat", padx=6, pady=4, cursor="hand2",
                  activeforeground=C_ARC,
                  command=self._open_full_window).pack(side="right", padx=(0,4))

        self.sbtn = tk.Button(inf, text="▶", font=("Segoe UI",12,"bold"),
                              bg=C_ARC, fg=C_BG, relief="flat", padx=12, pady=4,
                              cursor="hand2", activebackground=C_GOLD,
                              command=self._send)
        self.sbtn.pack(side="right")

        bdr = tk.Frame(inf, bg=C_ARC, padx=1, pady=1)
        bdr.pack(side="left", fill="x", expand=True, padx=(0,8))
        self.ivar = tk.StringVar()
        self.ibox = tk.Entry(bdr, textvariable=self.ivar, font=FONT_MONO,
                             bg=C_CARD, fg=C_WHITE, insertbackground=C_ARC,
                             relief="flat", bd=0)
        self.ibox.pack(fill="x", ipady=7, padx=2, pady=1)
        self.ibox.bind("<Return>",  lambda e: self._send())
        self.ibox.bind("<Up>",      self._hist_up)
        self.ibox.bind("<Down>",    self._hist_down)
        self.ibox.bind("<Escape>",  lambda e: self._close())
        self.ibox.focus()

        # ── Status bar — just above the input bar ─────────────────
        sb2 = tk.Frame(self, bg=C_PANEL, height=22)
        sb2.pack(side="bottom", fill="x")
        sb2.pack_propagate(False)
        self.stlbl = tk.Label(sb2, text="Ready", font=FONT_SMALL, bg=C_PANEL, fg=C_MUTED)
        self.stlbl.pack(side="left", padx=10)
        self.vlbl = tk.Label(sb2, font=FONT_SMALL, bg=C_PANEL, cursor="hand2")
        self.vlbl.pack(side="right", padx=8)
        self.vlbl.bind("<Button-1>", lambda e: self._toggle_voice())
        self._refresh_voice_lbl()
        tk.Label(sb2, text="voice:", font=FONT_SMALL, bg=C_PANEL, fg=C_MUTED).pack(side="right")
        tk.Label(sb2, text="⚙", font=FONT_SMALL, bg=C_PANEL, fg=C_PURPLE, cursor="hand2"
                 ).pack(side="right", padx=(0,2))

        # ── Chat display — fills all remaining space between buttons and input ──
        cf = tk.Frame(self, bg=C_BG)
        cf.pack(fill="both", expand=True)
        sb_scroll = tk.Scrollbar(cf, bg=C_CARD, troughcolor=C_PANEL, bd=0, relief="flat")
        sb_scroll.pack(side="right", fill="y")
        self.chat = tk.Text(cf, font=FONT_MONO, bg=C_BG, fg=C_WHITE,
                            relief="flat", state="disabled", wrap="word",
                            padx=14, pady=10, cursor="arrow",
                            selectbackground=C_ARC2, spacing1=2, spacing3=4)
        self.chat.pack(fill="both", expand=True)
        self.chat.config(yscrollcommand=sb_scroll.set)
        sb_scroll.config(command=self.chat.yview)

        self.chat.tag_config("stark",  foreground=C_ARC,    font=("Courier New",9,"bold"))
        self.chat.tag_config("you",    foreground=C_GOLD,   font=("Courier New",9,"bold"))
        self.chat.tag_config("info",   foreground=C_WHITE,  font=FONT_MONO)
        self.chat.tag_config("action", foreground=C_GREEN,  font=FONT_MONO)
        self.chat.tag_config("media",  foreground=C_AMBER,  font=FONT_MONO)
        self.chat.tag_config("error",  foreground=C_RED,    font=FONT_MONO)
        self.chat.tag_config("muted",  foreground=C_MUTED,  font=FONT_MONO_S)
        self.chat.tag_config("stream", foreground=C_WHITE,  font=FONT_MONO)

    # ── Drag ──────────────────────────────────────────────────
    def _drag_start(self, e): self._px=e.x_root-self.winfo_x(); self._py=e.y_root-self.winfo_y()
    def _drag_move(self, e):  self.geometry(f"+{e.x_root-self._px}+{e.y_root-self._py}")

    # ── Write helpers ─────────────────────────────────────────
    def _write(self, text, tag="info", nl=True):
        """Write a line to the chat display."""
        self.chat.config(state="normal")
        self.chat.insert("end", (text+"\n") if nl else text, tag)
        self.chat.config(state="disabled")
        self.chat.see("end")

    def _write_token(self, t):
        """Append a streaming token."""
        self.chat.config(state="normal")
        self.chat.insert("end", t, "stream")
        self.chat.config(state="disabled")
        self.chat.see("end")

    def _greet(self):
        import random
        h  = datetime.now().hour
        g  = "Good morning" if h < 12 else ("Good afternoon" if h < 17 else "Good evening")
        day = datetime.now().strftime("%A")
        phrases = [
            f"{g}! All systems online. What can I do for you?",
            f"{g}! Ready and at your service. Just ask.",
            f"{g}! I'm Stark — your personal AI assistant. What do you need?",
            f"Hey! Happy {day}. I'm online and ready.",
            f"{g}! Fully operational. Ask me anything or give me a command.",
        ]
        greeting = random.choice(phrases)
        self._write(f"STARK  ›  {greeting}", "stark")
        self._write("         Ask questions  |  Give commands  |  Try: 'help' for a full list", "muted")
        self._write("")
        self.voice.speak(greeting, force=True)

    def _set_state(self, s):
        self.stlbl.config(text=s)
        self.orb.set_state({"Ready":"idle","Thinking...":"thinking","Speaking...":"speaking"}.get(s,"idle"))

    def _quick(self, cmd):
        self.ivar.set(cmd); self.ibox.icursor("end"); self.ibox.focus()
        if not cmd.endswith(" "): self._send()

    def _refresh_voice_lbl(self):
        on = self.voice.enabled
        self.vlbl.config(text="🔊 ON" if on else "🔇 OFF",
                         fg=C_GREEN if on else C_MUTED)

    def _toggle_voice(self):
        self.voice.toggle()
        self._refresh_voice_lbl()
        if self.voice.enabled: self.voice.speak("Voice enabled.", force=True)
        else: self.voice.stop()

    # ── Send ──────────────────────────────────────────────────
    def _send(self):
        if self._busy: return
        msg = self.ivar.get().strip()
        if not msg: return
        self._hist.append(msg); self._hidx=-1; self.ivar.set("")
        self._write(f"YOU    ›  {msg}", "you")

        result = self.parser.parse(msg)
        if result:
            tag, resp = result
            self._write(f"STARK  ›  {resp}", tag)
            self._write("")
            self.voice.speak(resp)
            self._set_state("Ready")
            return

        if not self.ai.is_online():
            self._write("STARK  ›  LM Studio is offline. Start the local server on port 1234.", "error")
            self._write("")
            return

        self._busy=True; self._ai_buf=""
        self.sbtn.config(state="disabled", text="…")
        self._set_state("Thinking...")
        self._write("STARK  ›  ", "stark", nl=False)

        self.ai.chat_stream(
            msg,
            on_token = lambda t: self.after(0, self._write_token, t),
            on_done  = lambda full: self.after(0, self._done, full),
            on_error = lambda e: self.after(0, self._err, e),
        )

    def _done(self, full):
        self._write("\n"); self._write("")
        self._busy=False; self.sbtn.config(state="normal", text="▶")
        self._set_state("Speaking...")
        self.voice.speak(full)
        delay = max(2000, min(len(full.split())*320, 14000))
        self.after(delay, lambda: self._set_state("Ready"))

    def _err(self, e):
        self._write(f"\nSTARK  ›  Error: {e}", "error"); self._write("")
        self._busy=False; self.sbtn.config(state="normal", text="▶")
        self._set_state("Ready")

    def _hist_up(self, e):
        if not self._hist: return
        self._hidx=min(self._hidx+1,len(self._hist)-1)
        self.ivar.set(self._hist[-(self._hidx+1)]); self.ibox.icursor("end")

    def _hist_down(self, e):
        if self._hidx<=0: self._hidx=-1; self.ivar.set(""); return
        self._hidx-=1; self.ivar.set(self._hist[-(self._hidx+1)]); self.ibox.icursor("end")

    def _status_loop(self):
        def check():
            online = self.ai.is_online()
            self.dot.config(fg=C_GREEN if online else C_RED)
            self.dlbl.config(text="online" if online else "offline",
                             fg=C_GREEN if online else C_RED)
            self.after(6000, check)
        threading.Thread(target=lambda:(time.sleep(0.8),self.after(0,check)),daemon=True).start()

    def _open_full_window(self):
        FullAssistantWindow(self, self.parser, self.ai, self.voice, self.fm, self.media)

    def _close(self):
        self.orb.set_open(False); self.destroy()


# ════════════════════════════════════════════════════════════════════
#  FULL ASSISTANT WINDOW
# ════════════════════════════════════════════════════════════════════
class FullAssistantWindow(tk.Toplevel):
    def __init__(self, parent, parser: CommandParser, ai: AIClient,
                 voice: VoiceEngine, fm: FileManager, media: MediaPlugin):
        super().__init__(parent)
        self.parser = parser
        self.ai     = ai
        self.voice  = voice
        self.fm     = fm
        self.media  = media

        self.title("Stark — Full Assistant")
        self.geometry("1100x760"); self.minsize(860,580)
        self.configure(bg=C_BG); self.resizable(True,True)

        self._busy   = False
        self._ai_buf = ""
        self._hist   = []
        self._hidx   = -1

        self._build()
        self._status_loop()
        self._welcome()

    def _build(self):
        hdr = tk.Frame(self, bg=C_PANEL, height=52); hdr.pack(fill="x"); hdr.pack_propagate(False)
        tk.Label(hdr, text="S", font=("Segoe UI",18,"bold"), bg=C_PANEL, fg=C_ARC
                 ).pack(side="left", padx=(14,4))
        tk.Label(hdr, text="STARK", font=("Courier New",13,"bold"), bg=C_PANEL, fg=C_WHITE
                 ).pack(side="left")
        tk.Label(hdr, text="  FULL MODE", font=("Segoe UI",8), bg=C_PANEL, fg=C_MUTED
                 ).pack(side="left", pady=(18,0))
        self.dot  = tk.Label(hdr, text="●", font=("Segoe UI",10), bg=C_PANEL, fg=C_MUTED)
        self.dlbl = tk.Label(hdr, text="checking...", font=FONT_SMALL, bg=C_PANEL, fg=C_MUTED)
        self.dlbl.pack(side="right", padx=(0,16)); self.dot.pack(side="right", padx=(0,4))

        for txt,fg,cmd in [
            ("🎵 Media Browser", C_PINK,   lambda: MediaBrowserPanel(self, self.media)),
            ("📁 File Browser",  C_TEAL,   lambda: FileBrowserPanel(self, self.fm)),
            ("⚙ Voice",         C_PURPLE,  lambda: VoiceSettingsPanel(self, self.voice)),
        ]:
            tk.Button(hdr, text=txt, font=FONT_BTN, bg=C_CARD, fg=fg, relief="flat",
                      padx=10, pady=4, cursor="hand2", command=cmd).pack(side="right", padx=(0,6), pady=10)

        self.vtbtn = tk.Button(hdr, font=FONT_BTN, bg=C_CARD, relief="flat",
                               padx=10, pady=4, cursor="hand2", command=self._toggle_voice)
        self.vtbtn.pack(side="right", padx=(0,4), pady=10)
        self._refresh_vtbtn()

        tabs = tk.Frame(self, bg=C_CARD); tabs.pack(fill="x")

        def qrow(label, color, commands, bg=C_CARD):
            rf = tk.Frame(tabs, bg=bg); rf.pack(fill="x", padx=4, pady=(3,0))
            tk.Label(rf, text=label, font=FONT_MONO_S, bg=bg, fg=C_MUTED, width=9).pack(side="left")
            for lbl,cmd in commands:
                tk.Button(rf, text=lbl, font=FONT_MONO_S, bg=bg, fg=color,
                          relief="flat", padx=6, pady=3, cursor="hand2",
                          activeforeground=C_ARC if color==C_MUTED else C_YELLOW,
                          command=lambda c=cmd: self._quick(c)).pack(side="left", padx=1, pady=2)

        qrow("GENERAL:", C_MUTED, [
            ("📊 System","system info"),("📋 Clipboard","show clipboard"),
            ("🔍 Google","search for "),("📱 Apps","list apps"),("🕐 Time","what time is it"),
        ])
        qrow("FILES:", C_ORANGE, [
            ("📄 New File","create file "),("📁 New Folder","create folder "),
            ("🗑 Delete","delete file "),("✏ Rename","rename "),("📋 Copy","copy file "),
            ("🚚 Move","move file "),("📂 List","list folder desktop"),
            ("🔎 Find","find *."),("📖 Read","read file "),("🌳 Tree","tree desktop"),
        ])
        qrow("MEDIA:", C_PINK, [
            ("▶ YouTube","youtube "),("🎵 Spotify","spotify"),("🎬 VLC","vlc "),
            ("📻 WMP","windows media player"),("🎵 Music","list music"),
            ("🎬 Videos","list videos"),("🔎 Find♪","find music "),
            ("🔎 Find▶","find video "),("▶ Play","play "),("👁 Watch","watch "),
        ])
        qrow("VOICE:", C_ARC, [
            ("🔊 Say","say "),("⏹ Stop","stop speaking"),
            ("▶ Slow","set speed slow"),("▶ Normal","set speed normal"),
            ("▶ Fast","set speed fast"),("📢 Vol+","set volume 90"),("🔉 Vol-","set volume 50"),
        ])

        tk.Frame(self, bg=C_BORDER, height=1).pack(fill="x")

        main = tk.Frame(self, bg=C_BG); main.pack(fill="both", expand=True)

        side = tk.Frame(main, bg=C_PANEL, width=190); side.pack(side="right", fill="y"); side.pack_propagate(False)
        tk.Label(side, text="LIVE STATS", font=FONT_SMALL, bg=C_PANEL, fg=C_MUTED).pack(pady=(10,2), padx=10, anchor="w")
        tk.Frame(side, bg=C_BORDER, height=1).pack(fill="x", padx=8)
        self.stats = tk.Text(side, font=FONT_MONO_S, bg=C_PANEL, fg=C_MUTED, relief="flat",
                             state="disabled", wrap="word", padx=8, pady=6, height=8)
        self.stats.pack(fill="x", padx=4, pady=4)
        tk.Frame(side, bg=C_BORDER, height=1).pack(fill="x", padx=8)
        tk.Label(side, text="QUICK OPEN", font=FONT_SMALL, bg=C_PANEL, fg=C_MUTED).pack(pady=(10,4), padx=10, anchor="w")
        for lbl,cmd in [("🖥 Desktop","open folder desktop"),("📄 Documents","open folder documents"),
                        ("⬇ Downloads","open folder downloads"),("🎵 Music","open folder music"),
                        ("🎬 Videos","open folder videos"),("🧮 Calculator","open calculator"),
                        ("📝 Notepad","open notepad"),("💻 CMD","open command prompt"),
                        ("⚙ Settings","open settings")]:
            tk.Button(side, text=lbl, font=FONT_MONO_S, bg=C_PANEL, fg=C_MUTED, relief="flat",
                      padx=8, pady=3, cursor="hand2", anchor="w", width=20,
                      activeforeground=C_ARC,
                      command=lambda c=cmd: self._quick(c)).pack(fill="x", padx=6, pady=1)
        tk.Button(side, text="⟳ Refresh", font=FONT_SMALL, bg=C_CARD, fg=C_MUTED,
                  relief="flat", cursor="hand2", command=self._refresh_stats).pack(pady=(10,4))
        self._refresh_stats()

        cf = tk.Frame(main, bg=C_BG); cf.pack(side="left", fill="both", expand=True)
        sb = tk.Scrollbar(cf, bg=C_CARD, troughcolor=C_PANEL, bd=0, relief="flat")
        sb.pack(side="right", fill="y")
        self.chat = tk.Text(cf, font=FONT_MONO, bg=C_BG, fg=C_WHITE, relief="flat",
                            state="disabled", wrap="word", padx=16, pady=12,
                            cursor="arrow", selectbackground=C_ARC2, spacing1=2, spacing3=4)
        self.chat.pack(fill="both", expand=True)
        self.chat.config(yscrollcommand=sb.set); sb.config(command=self.chat.yview)
        for tag,fg,fnt in [
            ("stark",  C_ARC,    ("Courier New",10,"bold")),
            ("you",    C_GOLD,   ("Courier New",10,"bold")),
            ("info",   C_WHITE,  FONT_MONO),
            ("action", C_GREEN,  FONT_MONO),
            ("media",  C_AMBER,  FONT_MONO),
            ("error",  C_RED,    FONT_MONO),
            ("muted",  C_MUTED,  FONT_MONO_S),
            ("stream", C_WHITE,  FONT_MONO),
        ]:
            self.chat.tag_config(tag, foreground=fg, font=fnt)

        st = tk.Frame(self, bg=C_PANEL, height=22); st.pack(fill="x"); st.pack_propagate(False)
        self.stlbl = tk.Label(st, text="Ready", font=FONT_SMALL, bg=C_PANEL, fg=C_MUTED)
        self.stlbl.pack(side="left", padx=10)

        inf = tk.Frame(self, bg=C_PANEL, pady=10, padx=12); inf.pack(fill="x")
        bdr = tk.Frame(inf, bg=C_ARC, padx=1, pady=1)
        bdr.pack(side="left", fill="x", expand=True, padx=(0,10))
        self.ivar = tk.StringVar()
        self.ibox = tk.Entry(bdr, textvariable=self.ivar, font=FONT_MONO,
                             bg=C_CARD, fg=C_WHITE, insertbackground=C_ARC, relief="flat", bd=0)
        self.ibox.pack(fill="x", ipady=8, padx=2, pady=1)
        self.ibox.bind("<Return>", lambda e: self._send())
        self.ibox.bind("<Up>",     self._hist_up)
        self.ibox.bind("<Down>",   self._hist_down)
        self.ibox.focus()
        self.sbtn = tk.Button(inf, text="SEND  ▶", font=FONT_BTN, bg=C_ARC, fg=C_BG,
                              relief="flat", padx=18, pady=6, cursor="hand2",
                              activebackground=C_GOLD, command=self._send)
        self.sbtn.pack(side="left")
        tk.Button(inf, text="Clear", font=FONT_BTN, bg=C_CARD, fg=C_MUTED, relief="flat",
                  padx=10, pady=6, cursor="hand2", activeforeground=C_RED,
                  command=self._clear).pack(side="left", padx=(6,0))

    def _write(self, text, tag="info", nl=True):
        self.chat.config(state="normal")
        self.chat.insert("end", (text+"\n") if nl else text, tag)
        self.chat.config(state="disabled"); self.chat.see("end")

    def _write_token(self, t):
        self.chat.config(state="normal"); self.chat.insert("end",t,"stream")
        self.chat.config(state="disabled"); self.chat.see("end")

    def _welcome(self):
        self._write("═"*66, "muted")
        self._write("  STARK FULL ASSISTANT  — All plugins active", "stark")
        self._write(f"  {datetime.now().strftime('%A, %B %d %Y  —  %I:%M %p')}", "muted")
        self._write("═"*66, "muted"); self._write("")
        self._write("  VOICE:  say Hello  |  voice on/off  |  set speed fast  |  stop speaking", "muted")
        self._write("  FILES:  create file x.txt  |  delete file x  |  rename a to b", "muted")
        self._write("  MEDIA:  play [song]  |  youtube [query]  |  spotify [artist]", "muted")
        self._write("  FIND:   find *.pdf in documents  |  find music beethoven", "muted")
        self._write("  SYS:    system info  |  what time is it  |  disk usage", "muted")
        self._write("  Use the quick buttons above for one-click commands.", "muted")
        self._write("")

    def _set_state(self, s): self.stlbl.config(text=s)

    def _quick(self, cmd):
        self.ivar.set(cmd); self.ibox.icursor("end"); self.ibox.focus()
        if not cmd.endswith(" "): self._send()

    def _refresh_vtbtn(self):
        on = self.voice.enabled
        self.vtbtn.config(text="🔊 Voice ON" if on else "🔇 Voice OFF",
                          bg=C_GREEN if on else C_CARD,
                          fg=C_BG if on else C_MUTED)

    def _toggle_voice(self):
        self.voice.toggle(); self._refresh_vtbtn()
        if self.voice.enabled: self.voice.speak("Voice enabled.", force=True)

    def _refresh_stats(self):
        info = get_system_info()
        self.stats.config(state="normal"); self.stats.delete("1.0","end")
        self.stats.insert("end", info); self.stats.config(state="disabled")

    def _status_loop(self):
        def check():
            online = self.ai.is_online()
            self.dot.config(fg=C_GREEN if online else C_RED)
            self.dlbl.config(text="LM Studio Online" if online else "LM Studio Offline",
                             fg=C_GREEN if online else C_RED)
            self._refresh_stats(); self.after(5000, check)
        threading.Thread(target=lambda:(time.sleep(0.6),self.after(0,check)),daemon=True).start()

    def _send(self):
        if self._busy: return
        msg = self.ivar.get().strip()
        if not msg: return
        self._hist.append(msg); self._hidx=-1; self.ivar.set("")
        self._write(f"YOU    ›  {msg}", "you")
        result = self.parser.parse(msg)
        if result:
            tag,resp = result
            self._write(f"STARK  ›  {resp}", tag); self._write("")
            self.voice.speak(resp); self._set_state("Ready"); return
        if not self.ai.is_online():
            self._write("STARK  ›  LM Studio offline — start the local server on port 1234.", "error")
            self._write(""); return
        self._busy=True; self._ai_buf=""
        self.sbtn.config(state="disabled", text="…")
        self._set_state("Thinking...")
        self._write("STARK  ›  ", "stark", nl=False)
        self.ai.chat_stream(
            msg,
            on_token = lambda t: self.after(0, self._write_token, t),
            on_done  = lambda full: self.after(0, self._done, full),
            on_error = lambda e: self.after(0, self._err, e),
        )

    def _done(self, full):
        self._write("\n"); self._write(""); self._busy=False
        self.sbtn.config(state="normal", text="SEND  ▶"); self._set_state("Ready")
        self.voice.speak(full)

    def _err(self, e):
        self._write(f"\nSTARK  ›  Error: {e}", "error"); self._write("")
        self._busy=False; self.sbtn.config(state="normal", text="SEND  ▶")

    def _hist_up(self, e):
        if not self._hist: return
        self._hidx=min(self._hidx+1,len(self._hist)-1)
        self.ivar.set(self._hist[-(self._hidx+1)]); self.ibox.icursor("end")

    def _hist_down(self, e):
        if self._hidx<=0: self._hidx=-1; self.ivar.set(""); return
        self._hidx-=1; self.ivar.set(self._hist[-(self._hidx+1)]); self.ibox.icursor("end")

    def _clear(self):
        self.chat.config(state="normal"); self.chat.delete("1.0","end")
        self.chat.config(state="disabled"); self.ai.clear(); self._welcome()


# ════════════════════════════════════════════════════════════════════
#  STARK APPLICATION
# ════════════════════════════════════════════════════════════════════
class StarkApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.withdraw()
        self.root.title("Stark")

        self.voice  = VoiceEngine()
        self.ai     = AIClient()
        self.fm     = FileManager()
        self.media  = MediaPlugin()
        self.apps   = AppScanner()
        self.parser = CommandParser(self.fm, self.media, self.apps, self.voice)

        self._mini: StarkMiniPanel | None = None

        self.orb = FloatingOrb(self.root, on_click=self._toggle)
        self.orb.set_open(False)

        self.root.protocol("WM_DELETE_WINDOW", self._quit)
        self._watch()

    def _toggle(self):
        if self._mini and self._mini.winfo_exists():
            self._mini._close()
            self._mini = None
            self.orb.set_open(False)
        else:
            self._mini = StarkMiniPanel(
                self.root, self.orb, self.parser,
                self.ai, self.voice, self.fm, self.media
            )
            self.orb.set_open(True)

    def _watch(self):
        if not self.orb.winfo_exists(): self._quit(); return
        self.root.after(400, self._watch)

    def _quit(self):
        try: self.voice.stop()
        except: pass
        try: self.root.quit(); self.root.destroy()
        except: pass

    def run(self):
        self.root.mainloop()


# ════════════════════════════════════════════════════════════════════
#  ENTRY POINT
# ════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    print("╔══════════════════════════════════════╗")
    print("║   STARK  —  AI Desktop Assistant     ║")
    print("╚══════════════════════════════════════╝")

    pkgs = {"requests":"requests","psutil":"psutil","pyperclip":"pyperclip"}
    missing = [pkg for mod,pkg in pkgs.items() if not importlib.util.find_spec(mod)]
    if missing:
        print(f"Installing: {', '.join(missing)}")
        subprocess.check_call([sys.executable,"-m","pip","install"]+missing, stdout=subprocess.DEVNULL)
        print("Done.")

    for mod,pkg,note in [
        ("pyttsx3",   "pyttsx3",   "voice output"),
        ("send2trash","send2trash","safe deletes"),
    ]:
        if not importlib.util.find_spec(mod):
            print(f"Installing {pkg} ({note})...")
            try:
                subprocess.check_call([sys.executable,"-m","pip","install",pkg], stdout=subprocess.DEVNULL)
                print(f"  ✓ {pkg} installed")
            except Exception as e:
                print(f"  ✗ {pkg} failed: {e}")

    try:
        import pyttsx3; HAS_TTS=True
    except ImportError:
        pass
    try:
        import send2trash; HAS_TRASH=True
    except ImportError:
        pass
    try:
        import psutil; HAS_PSUTIL=True
    except ImportError:
        pass
    try:
        import pyperclip; HAS_CLIP=True
    except ImportError:
        pass

    print()
    print("Starting Stark...")
    print("  → A glowing orb will appear on your desktop")
    print("  → Click the orb to open / close the panel")
    print("  → Drag the orb anywhere on your screen")
    print("  → Click ⊞ in the panel for full window mode")
    print("  → Make sure LM Studio is running on port 1234")
    print()

    StarkApp().run()