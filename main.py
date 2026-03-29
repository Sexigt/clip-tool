import os, sys, re, time, wave, threading, subprocess, json, logging, glob as globmod, shutil
from datetime import datetime
from collections import deque
from pathlib import Path

import cv2, numpy as np, bettercam, mss, pyaudiowpatch as pyaudio
from PIL import Image, ImageDraw
import pystray
from plyer import notification
from faster_whisper import WhisperModel
import customtkinter as ctk
import keyboard

# ── Config ────────────────────────────────────────────────────────────────────
CLIPS_DIR = Path.home() / "clip-tool" / "clips"
CLIPS_DIR.mkdir(parents=True, exist_ok=True)
TEMP_DIR = CLIPS_DIR / "_temp"
TEMP_DIR.mkdir(exist_ok=True)
AUDIO_DIR = Path(__file__).resolve().parent / "audio"
DEFAULT_TARGET_H = 1080
DEFAULT_FPS = 60
DEFAULT_CLIP_SEC = 60
BUFFER_SEC = 300
JPEG_QUALITY = 95
RECORD_RATE, CHUNK = 44100, 1024
WAKE_WORD, HOTKEY = "computer", "ctrl+shift+f9"
FFMPEG = None  # auto-detected
FFPLAY = None
CONFIG_FILE = CLIPS_DIR / "config.json"

# logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler(CLIPS_DIR / "cliptool.log", encoding="utf-8"),
    ])
logger = logging.getLogger("cliptool")
# rotate log if > 1MB
_log_path = CLIPS_DIR / "cliptool.log"
if _log_path.exists() and _log_path.stat().st_size > 1_000_000:
    _bak = CLIPS_DIR / "cliptool.log.bak"
    if _bak.exists(): _bak.unlink()
    _log_path.rename(_bak)


def load_config():
    defaults = {
        "monitor": 1, "mic_device": "", "desk_device": "",
        "resolution": f"{DEFAULT_TARGET_H}p", "fps": str(DEFAULT_FPS),
        "quality": str(JPEG_QUALITY), "default_clip_sec": DEFAULT_CLIP_SEC,
        "mic_volume": 1.0, "desk_volume": 0.5,
        "mic_on": True, "desk_on": True,
    }
    try:
        if CONFIG_FILE.exists():
            with open(CONFIG_FILE, "r") as f:
                saved = json.load(f)
            defaults.update({k: v for k, v in saved.items() if k in defaults})
    except (OSError, json.JSONDecodeError) as e:
        logger.warning(f"Config load error: {e}")
    return defaults


def save_config(cfg):
    try:
        with open(CONFIG_FILE, "w") as f:
            json.dump(cfg, f, indent=2)
    except OSError as e:
        logger.error(f"Config save error: {e}")


def cleanup_temps():
    try:
        shutil.rmtree(TEMP_DIR, ignore_errors=True)
        TEMP_DIR.mkdir(exist_ok=True)
    except OSError:
        pass


def fmt_duration(secs):
    if secs >= 60:
        m, s = divmod(int(secs), 60)
        return f"{m}m{s:02d}s"
    return f"{int(secs)}s"


def find_ffmpeg():
    global FFMPEG, FFPLAY
    checks = [
        ("ffmpeg", "ffplay"),
        (r"C:\ffmpeg\bin\ffmpeg.exe", r"C:\ffmpeg\bin\ffplay.exe"),
        (r"C:\Program Files\ffmpeg\bin\ffmpeg.exe", r"C:\Program Files\ffmpeg\bin\ffplay.exe"),
        (str(Path.home() / "ffmpeg" / "bin" / "ffmpeg.exe"),
         str(Path.home() / "ffmpeg" / "bin" / "ffplay.exe")),
    ]
    # also check WinGet packages
    local = Path(os.environ.get("LOCALAPPDATA", ""))
    for p in local.glob("Microsoft/WinGet/Packages/Gyan.FFmpeg*/ffmpeg-*/bin/ffmpeg.exe"):
        checks.insert(1, (str(p), str(p.parent / "ffplay.exe")))
    for p in local.glob("Microsoft/WinGet/Packages/FFmpeg*/ffmpeg-*/bin/ffmpeg.exe"):
        checks.insert(1, (str(p), str(p.parent / "ffplay.exe")))

    for ffmpeg_path, ffplay_path in checks:
        try:
            r = subprocess.run([ffmpeg_path, "-version"], capture_output=True, timeout=5)
            if r.returncode == 0:
                FFMPEG = ffmpeg_path
                FFPLAY = ffplay_path
                logger.info(f"found: {ffmpeg_path}")
                return True
        except Exception:
            continue
    logger.warning("NOT FOUND — clips will be saved as .avi (lower quality)")
    return False


def play_sound(name):
    path = AUDIO_DIR / name
    if not path.exists(): return
    if not FFPLAY: return
    try:
        subprocess.Popen([FFPLAY, "-nodisp", "-autoexit", "-loglevel", "quiet", str(path)],
                         stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception as e:
        logger.error(f"sound: {e}")


# ── Monitors (skip virtual combined) ──────────────────────────────────────────
def get_monitors():
    with mss.mss() as sct:
        return [{"index": i, "w": m["width"], "h": m["height"]}
                for i, m in enumerate(sct.monitors) if i > 0]


def get_audio_devices():
    pa = pyaudio.PyAudio()
    mics, loops = [], []
    for i in range(pa.get_device_count()):
        info = pa.get_device_info_by_index(i)
        if info["maxInputChannels"] > 0:
            e = {"index": i, "name": info["name"]}
            (loops if "loopback" in info["name"].lower() else mics).append(e)
    pa.terminate()
    return mics, loops


# ── Screen ────────────────────────────────────────────────────────────────────
class Screen:
    def __init__(self, mon=1, target_h=DEFAULT_TARGET_H, fps=DEFAULT_FPS,
                 buf_sec=BUFFER_SEC, jpeg_quality=JPEG_QUALITY):
        self.mon, self.fps = mon - 1, fps  # bettercam is 0-indexed
        self.target_h, self.jpeg_quality = target_h, jpeg_quality
        self._jpg_params = [cv2.IMWRITE_JPEG_QUALITY, jpeg_quality]
        self.buf = deque(maxlen=buf_sec * fps)
        self._raw_q = deque(maxlen=fps)  # small queue, frames get consumed fast
        self.frame_size = None
        self._last_jpg = None
        self._cam = None
        self._stop = threading.Event()

    def start(self):
        self._stop.clear()
        # get monitor dimensions via mss
        with mss.mss() as sct:
            mon_bounds = sct.monitors[self.mon + 1]  # +1 because mss[0] is virtual
            mw, mh = mon_bounds["width"], mon_bounds["height"]
        w = int(mw * self.target_h / mh)
        self.frame_size = (w, self.target_h)
        self._cam = bettercam.create(output_idx=self.mon,
                                     output_color="BGR",
                                     max_buffer_len=64)
        self._cam.start()
        threading.Thread(target=self._loop, daemon=True).start()
        threading.Thread(target=self._encoder, daemon=True).start()

    def stop(self):
        self._stop.set()
        if self._cam:
            try: self._cam.stop()
            except Exception:
                pass

    def _loop(self):
        interval = 1.0 / self.fps
        next_t = time.perf_counter()
        while not self._stop.is_set():
            frame = self._cam.get_latest_frame()
            if frame is not None:
                self._raw_q.append((time.time(), frame))
            next_t += interval
            wait = next_t - time.perf_counter()
            if wait > 0.003:
                time.sleep(wait - 0.002)
            while next_t - time.perf_counter() > 0:
                pass

    def _encoder(self):
        while not self._stop.is_set():
            if not self._raw_q:
                time.sleep(0.002)
                continue
            ts, frame = self._raw_q.popleft()
            resized = cv2.resize(frame, self.frame_size)
            _, jpg = cv2.imencode(".jpg", resized, self._jpg_params)
            self._last_jpg = jpg
            self.buf.append((ts, jpg))

    def last_frame(self):
        if self._last_jpg is not None:
            return cv2.imdecode(np.frombuffer(self._last_jpg, dtype=np.uint8), cv2.IMREAD_COLOR)
        return None

    def duration(self): return self.buf[-1][0] - self.buf[0][0] if len(self.buf) > 1 else 0
    def get_frames(self, s, e): return [(t, j) for t, j in self.buf if s <= t <= e]


# ── Audio ─────────────────────────────────────────────────────────────────────
class Audio:
    def __init__(self, mic_idx=None, desk_idx=None):
        self.mic_idx, self.desk_idx = mic_idx, desk_idx
        self.mic_buf = deque(maxlen=BUFFER_SEC * RECORD_RATE // CHUNK)
        self.desk_buf = deque(maxlen=BUFFER_SEC * RECORD_RATE // CHUNK)
        self.voice_buf = deque(maxlen=RECORD_RATE // CHUNK * 6)
        self._stop, self._pa, self._lock = threading.Event(), None, threading.Lock()
        self._desk_ch = 2
        self._desk_rate = 48000

    def start(self):
        self._pa = pyaudio.PyAudio()
        self._stop.clear()
        threading.Thread(target=self._mic, daemon=True).start()
        threading.Thread(target=self._desk, daemon=True).start()

    def stop(self):
        self._stop.set()
        time.sleep(0.15)
        if self._pa: self._pa.terminate()

    def _mic(self):
        try:
            kw = dict(format=pyaudio.paInt16, channels=1, rate=RECORD_RATE,
                      input=True, frames_per_buffer=CHUNK)
            if self.mic_idx is not None: kw["input_device_index"] = self.mic_idx
            s = self._pa.open(**kw)
            while not self._stop.is_set():
                d = s.read(CHUNK, exception_on_overflow=False)
                self.mic_buf.append((time.time(), d))
                with self._lock: self.voice_buf.append(d)
            s.stop_stream(); s.close()
        except Exception as e: logger.error(f"audio mic: {e}")

    def _desk(self):
        try:
            idx = self.desk_idx
            if idx is None:
                for i in range(self._pa.get_device_count()):
                    d = self._pa.get_device_info_by_index(i)
                    if d["maxInputChannels"] > 0 and "loopback" in d["name"].lower():
                        idx = i; break
            if idx is None: return
            info = self._pa.get_device_info_by_index(idx)
            self._desk_ch = int(info["maxInputChannels"])
            self._desk_rate = int(info["defaultSampleRate"])
            s = self._pa.open(format=pyaudio.paInt16, channels=self._desk_ch,
                              rate=self._desk_rate, input=True,
                              input_device_index=idx, frames_per_buffer=CHUNK)
            while not self._stop.is_set():
                self.desk_buf.append((time.time(), s.read(CHUNK, exception_on_overflow=False)))
            s.stop_stream(); s.close()
        except Exception as e: logger.error(f"audio desktop: {e}")

    def read_voice(self):
        with self._lock:
            c = list(self.voice_buf); self.voice_buf.clear(); return c

    def get_range(self, buf, s, e):
        return b"".join(d for t, d in buf if s <= t <= e)


# ── Voice ─────────────────────────────────────────────────────────────────────
class Voice:
    def __init__(self, audio):
        self.audio, self.on_cmd, self._stop, self._model = audio, None, threading.Event(), None

    def start(self, cb):
        self.on_cmd = cb
        if self._model is None:
            self._model = WhisperModel("base", device="cpu", compute_type="int8")
        self._stop.clear()
        threading.Thread(target=self._loop, daemon=True).start()

    def stop(self): self._stop.set()

    def _loop(self):
        last = 0
        while not self._stop.is_set():
            time.sleep(2)
            chunks = self.audio.read_voice()
            if not chunks: continue
            raw = np.frombuffer(b"".join(chunks), dtype=np.int16).astype(np.float32)
            ratio = RECORD_RATE / 16000
            idx = np.arange(0, len(raw), ratio).astype(int)
            idx = idx[idx < len(raw)]
            audio = raw[idx] / 32768.0
            if len(audio) < 8000: continue
            segs, _ = self._model.transcribe(audio, language="en", vad_filter=True,
                vad_parameters=dict(min_silence_duration_ms=300, speech_pad_ms=200))
            text = " ".join(s.text.strip() for s in segs).strip().lower()
            if not text: continue
            logger.info(f"voice: {text}")
            for w in text.split():
                clean = re.sub(r"[,.\!?;:]", "", w)
                if clean in ("computer", "comp") and time.time() - last > 3:
                    last = time.time(); self.on_cmd(text); break

    @staticmethod
    def parse(text, default_sec=DEFAULT_CLIP_SEC):
        text = text.lower()
        text = re.sub(r"[,.\!?;:]", "", text)
        text = re.sub(r"comp\w*", "computer", text)
        text = re.sub(r"equip\w*", "clip", text)
        text = re.sub(r"reclip\w*", "clip", text)
        text = re.sub(r"clip\w*", "clip", text)
        # normalize misheard phrases
        text = re.sub(r"clip\s+(at|out|of|it|up)\b", "clip that", text)
        text = re.sub(r"clipped?\s+(at|out|of|it|up)\b", "clip that", text)
        text = re.sub(r"\s+", " ", text).strip()
        if re.search(r"clip\s+(it|that|this|now)\s*$", text) or re.search(r"clip\s*$", text):
            return (0, default_sec)
        m = re.search(r"clip.*?(\d+)\s*min.*?(\d+)\s*min", text)
        if m: return (int(m.group(1)) * 60, int(m.group(2)) * 60)
        m = re.search(r"clip.*?last\s+(\d+)\s*(second|sec|minute|min)", text)
        if m:
            v, u = int(m.group(1)), m.group(2)
            return (0, v if u.startswith("sec") else v * 60)
        m = re.search(r"clip.*?(\d+):(\d+).*?(\d+):(\d+)", text)
        if m: return (int(m.group(1))*60+int(m.group(2)), int(m.group(3))*60+int(m.group(4)))
        return None


# ── Save ──────────────────────────────────────────────────────────────────────
def _wav(path, data, ch, rate=RECORD_RATE):
    with wave.open(path, "wb") as w:
        w.setnchannels(ch); w.setsampwidth(2); w.setframerate(rate); w.writeframes(data)


def save(screen, audio, so, eo, fps, mic_vol=1.0, desk_vol=0.5,
         mic_on=True, desk_on=True, progress_cb=None):
    now = time.time()
    cs, ce = now - eo, now - so
    frames = screen.get_frames(cs, ce)
    if not frames: return None

    w, h = screen.frame_size
    ts = datetime.now().strftime('%Y%m%d_%H%M%S')
    clip_dur = frames[-1][0] - frames[0][0]
    dur_str = fmt_duration(clip_dur)
    final = str(CLIPS_DIR / f"clip_{ts}_{dur_str}.mp4")

    if FFMPEG:
        TEMP_DIR.mkdir(exist_ok=True)
        enc = str(TEMP_DIR / f"_e_{ts}.mp4")
        concat = str(TEMP_DIR / f"_c_{ts}.txt")
        tm = str(TEMP_DIR / f"_m_{ts}.wav")
        td = str(TEMP_DIR / f"_d_{ts}.wav")

        # write JPEG frames + concat file with actual per-frame durations (VFR)
        with open(concat, "w") as cf:
            for i, (t, jpg) in enumerate(frames):
                jpath = str(TEMP_DIR / f"_j_{ts}_{i:06d}.jpg")
                with open(jpath, "wb") as jf:
                    jf.write(jpg.tobytes())
                if i + 1 < len(frames):
                    dur = frames[i + 1][0] - frames[i][0]
                else:
                    dur = (frames[i][0] - frames[i - 1][0]) if i > 0 else 1.0 / fps
                cf.write(f"file '{jpath}'\n")
                cf.write(f"duration {dur:.6f}\n")
        if progress_cb: progress_cb(20)

        # pass 1: concat JPEGs → VFR video
        subprocess.run([FFMPEG, "-y", "-loglevel", "error",
            "-f", "concat", "-safe", "0", "-i", concat,
            "-c:v", "libx264", "-preset", "fast", "-crf", "16",
            "-pix_fmt", "yuv420p", enc], capture_output=True)
        if progress_cb: progress_cb(45)

        # cleanup JPEG frames + concat
        for i in range(len(frames)):
            jpath = str(TEMP_DIR / f"_j_{ts}_{i:06d}.jpg")
            try:
                if os.path.exists(jpath): os.remove(jpath)
            except OSError: pass
        try:
            if os.path.exists(concat): os.remove(concat)
        except OSError: pass

        if not os.path.exists(enc) or os.path.getsize(enc) < 500:
            logger.error("save failed: video encode")
            return None

        # audio
        mic = audio.get_range(audio.mic_buf, cs, ce) if mic_on else b""
        desk = audio.get_range(audio.desk_buf, cs, ce) if desk_on else b""
        hm, hd = bool(mic), bool(desk)
        if hm: _wav(tm, mic, 1, RECORD_RATE)
        if hd: _wav(td, desk, audio._desk_ch, audio._desk_rate)
        if progress_cb: progress_cb(65)

        # pass 2: mux video + audio
        if hm and hd:
            subprocess.run([FFMPEG, "-y", "-loglevel", "error",
                "-i", enc, "-i", tm, "-i", td,
                "-filter_complex",
                f"[1]aresample={RECORD_RATE},volume={mic_vol}[a];[2]aresample={RECORD_RATE},volume={desk_vol}[b];[a][b]amix=inputs=2:duration=shortest[m]",
                "-map", "0:v", "-map", "[m]",
                "-c:v", "copy", "-c:a", "aac", "-b:a", "192k",
                "-movflags", "+faststart", final], capture_output=True)
        elif hm:
            subprocess.run([FFMPEG, "-y", "-loglevel", "error",
                "-i", enc, "-i", tm,
                "-af", f"aresample={RECORD_RATE},volume={mic_vol}",
                "-c:v", "copy", "-c:a", "aac", "-b:a", "192k",
                "-movflags", "+faststart", final], capture_output=True)
        elif hd:
            subprocess.run([FFMPEG, "-y", "-loglevel", "error",
                "-i", enc, "-i", td,
                "-map", "0:v", "-map", "1:a",
                "-c:v", "copy", "-c:a", "aac", "-b:a", "192k",
                "-af", f"aresample={RECORD_RATE},volume={desk_vol}",
                "-movflags", "+faststart", final], capture_output=True)
        else:
            os.replace(enc, final)
        if progress_cb: progress_cb(90)

        for f in [enc, tm, td]:
            try:
                if os.path.exists(f): os.remove(f)
            except OSError: pass

        if not os.path.exists(final) or os.path.getsize(final) < 1000:
            logger.error("save failed: mux")
            return None
    else:
        # no ffmpeg — decode JPEGs and save as avi
        avi = str(CLIPS_DIR / f"clip_{ts}_{dur_str}.avi")
        fourcc = cv2.VideoWriter_fourcc(*"XVID")
        out = cv2.VideoWriter(avi, fourcc, fps, (w, h))
        if not out.isOpened():
            logger.error("save failed: VideoWriter")
            return None
        for _, jpg in frames:
            frame = cv2.imdecode(np.frombuffer(jpg.tobytes(), dtype=np.uint8), cv2.IMREAD_COLOR)
            out.write(frame)
        out.release()
        final = avi
        if progress_cb: progress_cb(90)

    if progress_cb: progress_cb(100)
    dur = frames[-1][0] - frames[0][0]
    mb = os.path.getsize(final) / 1048576
    logger.info(f"saved {Path(final).name} ({dur:.1f}s, {mb:.1f}MB)")
    return Path(final)


# ── Minimal UI ────────────────────────────────────────────────────────────────
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("dark-blue")
        self.title("ClipTool")
        self.geometry("340x370")
        self.resizable(False, False)
        self.protocol("WM_DELETE_WINDOW", self._quit)

        self.screen, self.audio, self.voice = None, None, None
        self.active, self._tray_icon = False, None
        self.clips = []
        self._save_timer = None
        self._loading = False

        cleanup_temps()
        self.config = load_config()
        self.mics, self.loops = get_audio_devices()
        self.mons = get_monitors()
        find_ffmpeg()

        self._build()
        self._loading = True
        self._apply_config()
        self._loading = False
        self._tray()
        self._hotkey()
        self.after(500, self._start)
        self.after(2000, self._update_status)

    def _build(self):
        # header
        ctk.CTkLabel(self, text="ClipTool", font=("Segoe UI", 22, "bold")).pack(pady=(12, 0))
        self.status = ctk.CTkLabel(self, text="Starting...", font=("Segoe UI", 13), text_color="#ffa500")
        self.status.pack()
        self.buf_label = ctk.CTkLabel(self, text="", font=("Segoe UI", 11), text_color="#666")
        self.buf_label.pack(pady=(0, 5))

        # device selectors — all centered at width 280
        f = ctk.CTkFrame(self, fg_color="transparent", width=300)
        f.pack(pady=3)
        f.pack_propagate(False)

        for label, values in [
            ("Monitor", [f"Monitor {m['index']} ({m['w']}x{m['h']})" for m in self.mons] or ["None"]),
            ("Mic", [x["name"][:40] for x in self.mics] or ["Default"]),
            ("Desktop", [x["name"][:40] for x in self.loops] or ["None"]),
        ]:
            r = ctk.CTkFrame(f, fg_color="transparent")
            r.pack(fill="x", pady=1)
            ctk.CTkLabel(r, text=f"{label}:", font=("Segoe UI", 9), width=50,
                         anchor="e").pack(side="left")
            var = ctk.StringVar(value=values[0])
            ctk.CTkOptionMenu(r, variable=var, values=values, width=240, height=24,
                              font=("Segoe UI", 9)).pack(side="left", padx=(5, 0))
            if label == "Monitor":
                self.mon_v = var
                self.mon_v.trace_add("write", self._restart_screen)
            elif label == "Mic":
                self.mic_v = var
            else:
                self.desk_v = var

        # capture settings — centered row
        sf = ctk.CTkFrame(f, fg_color="transparent")
        sf.pack(pady=(5, 0))
        for lbl, var_name, opts, default in [
            ("Res", "res_v", ["720p", "1080p", "1440p", "Native"], f"{DEFAULT_TARGET_H}p"),
            ("FPS", "fps_v", ["30", "60", "120"], str(DEFAULT_FPS)),
            ("Q", "qual_v", ["70", "85", "95"], str(JPEG_QUALITY)),
        ]:
            ctk.CTkLabel(sf, text=f"{lbl}:", font=("Segoe UI", 9)).pack(side="left", padx=(8, 1))
            var = ctk.StringVar(value=default)
            ctk.CTkOptionMenu(sf, variable=var, values=opts, width=65, height=22,
                              font=("Segoe UI", 9)).pack(side="left")
            setattr(self, var_name, var)
            var.trace_add("write", self._restart_screen)

        # volume sliders — grid-aligned
        vf = ctk.CTkFrame(self, fg_color="transparent", width=300)
        vf.pack(pady=5)
        vf.pack_propagate(False)

        for label, var_name, default, lbl_name in [
            ("Mic", "mic_vol", 1.0, "mic_vol_lbl"),
            ("Desktop", "desk_vol", 0.5, "desk_vol_lbl"),
        ]:
            r = ctk.CTkFrame(vf, fg_color="transparent")
            r.pack(fill="x", pady=1)
            ctk.CTkLabel(r, text=f"{label}:", font=("Segoe UI", 9), width=50,
                         anchor="e").pack(side="left")
            var = ctk.DoubleVar(value=default)
            ctk.CTkSlider(r, from_=0, to=2, variable=var, width=190).pack(side="left", padx=(5, 4))
            lbl = ctk.CTkLabel(r, text=f"{int(default*100)}%", font=("Segoe UI", 9), width=32)
            lbl.pack(side="left")
            setattr(self, var_name, var)
            setattr(self, lbl_name, lbl)
            var.trace_add("write", lambda *a, v=var, l=lbl: l.configure(text=f"{int(v.get()*100)}%"))

        # record toggles
        tf = ctk.CTkFrame(vf, fg_color="transparent")
        tf.pack(pady=(5, 0))
        self.mic_on = ctk.BooleanVar(value=True)
        ctk.CTkSwitch(tf, text="Mic", variable=self.mic_on, font=("Segoe UI", 9),
                      command=self._toggle_mic).pack(side="left", padx=(0, 15))
        self.desk_on = ctk.BooleanVar(value=True)
        ctk.CTkSwitch(tf, text="Desktop", variable=self.desk_on, font=("Segoe UI", 9),
                      command=self._toggle_desk).pack(side="left", padx=(0, 15))
        ctk.CTkLabel(tf, text="Default:", font=("Segoe UI", 9)).pack(side="left")
        self.default_sec_v = ctk.StringVar(value=f"{DEFAULT_CLIP_SEC // 60}min")
        ctk.CTkOptionMenu(tf, variable=self.default_sec_v,
                          values=["30s", "1min", "2min", "5min"],
                          width=55, height=22,
                          font=("Segoe UI", 9)).pack(side="left", padx=(2, 0))

        # action feedback
        self.action_lbl = ctk.CTkLabel(self, text="", font=("Segoe UI", 12, "bold"),
                                        text_color="#4ecca3", height=24)
        self.action_lbl.pack(pady=(3, 0))

        # buttons
        bf = ctk.CTkFrame(self, fg_color="transparent")
        bf.pack(pady=(3, 12))
        self.clip_btn = ctk.CTkButton(bf, text="Clip That", width=100, height=30,
                      font=("Segoe UI", 11), fg_color="#e94560", hover_color="#c73e54",
                      command=lambda: self._clip_default())
        self.clip_btn.pack(side="left", padx=2)
        ctk.CTkButton(bf, text="1min", width=40, height=30, font=("Segoe UI", 11),
                      fg_color="#1a1a2e", hover_color="#16213e",
                      command=lambda: self._clip(0, 60)).pack(side="left", padx=2)
        ctk.CTkButton(bf, text="2min", width=40, height=30, font=("Segoe UI", 11),
                      fg_color="#1a1a2e", hover_color="#16213e",
                      command=lambda: self._clip(0, 120)).pack(side="left", padx=2)
        ctk.CTkButton(bf, text="5min", width=40, height=30, font=("Segoe UI", 11),
                      fg_color="#1a1a2e", hover_color="#16213e",
                      command=lambda: self._clip(0, 300)).pack(side="left", padx=2)

        # wire config saving to all variables
        for v in [self.mon_v, self.mic_v, self.desk_v, self.res_v,
                  self.fps_v, self.qual_v, self.mic_vol, self.desk_vol,
                  self.default_sec_v]:
            v.trace_add("write", lambda *a: self._schedule_save())
        self.mic_on.trace_add("write", lambda *a: self._schedule_save())
        self.desk_on.trace_add("write", lambda *a: self._schedule_save())

    def _apply_config(self):
        c = self.config
        # monitor
        for m in self.mons:
            if m["index"] == c["monitor"]:
                self.mon_v.set(f"Monitor {m['index']} ({m['w']}x{m['h']})")
                break
        # devices
        if c["mic_device"]:
            for d in self.mics:
                if d["name"][:40] == c["mic_device"]:
                    self.mic_v.set(c["mic_device"]); break
        if c["desk_device"]:
            for d in self.loops:
                if d["name"][:40] == c["desk_device"]:
                    self.desk_v.set(c["desk_device"]); break
        # capture settings
        self.res_v.set(c["resolution"])
        self.fps_v.set(c["fps"])
        self.qual_v.set(c["quality"])
        # default clip duration
        dsec = c.get("default_clip_sec", DEFAULT_CLIP_SEC)
        self.default_sec_v.set(self._sec_to_label(dsec))
        # volumes
        self.mic_vol.set(c["mic_volume"])
        self.desk_vol.set(c["desk_volume"])
        # toggles
        self.mic_on.set(c["mic_on"])
        self.desk_on.set(c["desk_on"])

    def _schedule_save(self):
        if self._loading: return
        if self._save_timer:
            self.after_cancel(self._save_timer)
        self._save_timer = self.after(500, self._save_config)

    def _save_config(self):
        self._save_timer = None
        if self._loading: return
        mon_idx = 1
        for m in self.mons:
            if f"Monitor {m['index']}" in self.mon_v.get(): mon_idx = m["index"]
        self.config = {
            "monitor": mon_idx,
            "mic_device": self.mic_v.get(),
            "desk_device": self.desk_v.get(),
            "resolution": self.res_v.get(),
            "fps": self.fps_v.get(),
            "quality": self.qual_v.get(),
            "default_clip_sec": self._get_default_sec(),
            "mic_volume": round(self.mic_vol.get(), 2),
            "desk_volume": round(self.desk_vol.get(), 2),
            "mic_on": self.mic_on.get(),
            "desk_on": self.desk_on.get(),
        }
        save_config(self.config)

    def _update_status(self):
        if self.active and self.screen:
            d = self.screen.duration()
            mins, secs = divmod(int(d), 60)
            self.buf_label.configure(text=f"Buffer: {mins}m {secs}s")
            self.status.configure(text="● REC", text_color="#e94560")
        else:
            self.buf_label.configure(text="")
            self.status.configure(text="● OFF", text_color="#666")
        self.after(1000, self._update_status)

    # ── Start ─────────────────────────────────────────────────────────────
    def _start(self):
        mon = 1
        for m in self.mons:
            if f"Monitor {m['index']}" in self.mon_v.get(): mon = m["index"]
        mi = self._find_idx(self.mics, self.mic_v.get())
        di = self._find_idx(self.loops, self.desk_v.get())

        target_h = self._get_target_h()
        fps = int(self.fps_v.get())
        quality = int(self.qual_v.get())

        self.screen = Screen(mon, target_h=target_h, fps=fps, jpeg_quality=quality)
        self.audio = Audio(mi, di)
        self.voice = Voice(self.audio)
        self.screen.start()
        self.audio.start()
        self.voice.start(self._on_voice)
        self.active = True
        logger.info(f"running — {target_h}p {fps}fps q{quality} — say 'computer clip that'")

    def _stop(self):
        self.active = False
        if self.voice: self.voice.stop()
        if self.audio: self.audio.stop()
        if self.screen: self.screen.stop()

    def _restart_screen(self, *a):
        if self._loading: return
        if self.screen:
            self.screen.stop()
            mon = 1
            for m in self.mons:
                if f"Monitor {m['index']}" in self.mon_v.get(): mon = m["index"]
            target_h = self._get_target_h()
            fps = int(self.fps_v.get())
            quality = int(self.qual_v.get())
            self.screen = Screen(mon, target_h=target_h, fps=fps, jpeg_quality=quality)
            self.screen.start()

    def _get_target_h(self):
        r = self.res_v.get()
        if r == "Native" and self.mons:
            for m in self.mons:
                if f"Monitor {m['index']}" in self.mon_v.get():
                    return m["h"]
            return self.mons[0]["h"]
        return int(r.replace("p", ""))

    def _find_idx(self, lst, name):
        for d in lst:
            if d["name"][:40] == name: return d["index"]
        return None

    def _get_default_sec(self):
        return self._label_to_sec(self.default_sec_v.get())

    @staticmethod
    def _label_to_sec(label):
        label = label.strip().lower()
        if label.endswith("min"):
            return int(label.replace("min", "")) * 60
        if label.endswith("s"):
            return int(label.replace("s", ""))
        return DEFAULT_CLIP_SEC

    @staticmethod
    def _sec_to_label(sec):
        if sec >= 60:
            return f"{sec // 60}min"
        return f"{sec}s"

    def _clip_default(self):
        self._clip(0, self._get_default_sec())

    # ── Clipping ──────────────────────────────────────────────────────────
    def _toggle_mic(self):
        play_sound("sound-on.mp3" if self.mic_on.get() else "sound-off.mp3")

    def _toggle_desk(self):
        play_sound("sound-on.mp3" if self.desk_on.get() else "sound-off.mp3")

    def _on_voice(self, text):
        p = Voice.parse(text, default_sec=self._get_default_sec())
        if p:
            s, e = p
            d = self.screen.duration() if self.screen else 0
            if d > 0:
                self.after(0, lambda: self._show_action("✂ Clipping..."))
                self.after(0, lambda: self._clip(s, min(e, d)))
        else:
            logger.debug(f"can't parse: {text}")

    def _clip(self, s, e):
        if not self.active or not self.screen: return
        d = self.screen.duration()
        if d < 1: self._toast("Buffering..."); return
        play_sound("clipped.mp3")
        self._show_action("✂ Clipping...")
        threading.Thread(target=self._do_clip, args=(s, min(e, d)), daemon=True).start()

    def _do_clip(self, s, e):
        try:
            fps = int(self.fps_v.get())
            def on_progress(pct):
                self.after(0, lambda: self._show_action(f"Saving... {pct}%"))
            p = save(self.screen, self.audio, s, e, fps=fps,
                     mic_vol=self.mic_vol.get(), desk_vol=self.desk_vol.get(),
                     mic_on=self.mic_on.get(), desk_on=self.desk_on.get(),
                     progress_cb=on_progress)
            if p: self.after(0, lambda: self._saved(p))
            else: self.after(0, lambda: self._show_action("Save failed", err=True))
        except Exception as ex:
            logger.error(f"save error: {ex}", exc_info=True)
            self.after(0, lambda: self._show_action(f"Error: {ex}", err=True))

    def _saved(self, p):
        self.clips.insert(0, p)
        self._show_action(f"Saved: {p.name}")
        self._toast(f"Saved: {p.name}")
        # clear after 4 seconds
        self.after(4000, lambda: self.action_lbl.configure(text=""))

    def _show_action(self, text, err=False):
        self.action_lbl.configure(text=text, text_color="#e94560" if err else "#4ecca3")

    def _open_folder(self):
        if sys.platform == "win32": os.startfile(str(CLIPS_DIR))

    # ── Hotkey ────────────────────────────────────────────────────────────
    def _hotkey(self):
        def h():
            d = self.screen.duration() if self.screen else 0
            sec = self._get_default_sec()
            self.after(0, lambda: self._clip(0, min(sec, d)))
        try: keyboard.add_hotkey(HOTKEY, h)
        except Exception: pass

    # ── Tray ──────────────────────────────────────────────────────────────
    def _tray(self):
        img = Image.new("RGB", (64, 64), "#1a1a2e")
        ImageDraw.Draw(img).ellipse([10, 10, 54, 54], fill="#e94560")
        m = pystray.Menu(
            pystray.MenuItem("Show", lambda i, s: self.after(0, self._show)),
            pystray.MenuItem("Clip 1min", lambda i, s: self.after(0, lambda: self._clip(0, 60))),
            pystray.MenuItem("Clip 5min", lambda i, s: self.after(0, lambda: self._clip(0, 300))),
            pystray.MenuItem("Open Folder", lambda i, s: self.after(0, self._open_folder)),
            pystray.MenuItem("Quit", self._quit))
        self._tray_icon = pystray.Icon("ClipTool", img, "ClipTool", m)
        threading.Thread(target=self._tray_icon.run, daemon=True).start()

    # ── Helpers ───────────────────────────────────────────────────────────
    def _toast(self, msg):
        logger.info(msg)
        try: notification.notify(title="ClipTool", message=msg, timeout=3)
        except Exception: pass

    def _show(self): self.deiconify(); self.lift()
    def _hide(self): self.withdraw()

    def _quit(self, *a):
        self._save_config()
        self._stop()
        cleanup_temps()
        try: self._tray_icon.stop()
        except Exception: pass
        keyboard.unhook_all()
        self.destroy()
        sys.exit(0)


if __name__ == "__main__":
    App().mainloop()
