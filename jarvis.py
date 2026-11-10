"""
Джарвис — голосовой ассистент для Windows.

Слух:    SpeechRecognition + PyAudio
Голос:   pyttsx3 (офлайн)
Мозг:    команды + локальная нейросеть через Ollama
Лицо:    rich (ASCII-аватар в консоли)

Установка:
    pip install SpeechRecognition PyAudio pyttsx3 rich ollama
    (если PyAudio не ставится:  pip install pipwin && pipwin install pyaudio)
    ollama pull gemma4:12b
"""

import ctypes
import os
import re
import subprocess
import threading
import time
import webbrowser

import ollama
import pyttsx3
import speech_recognition as sr
from rich.align import Align
from rich.console import Group
from rich.live import Live
from rich.panel import Panel
from rich.text import Text

# ───────────────────────── Настройки ─────────────────────────

MODEL = "gemma4:12b"          # можно заменить на любую модель из Ollama
LANGUAGE = "ru-RU"
WAKE_RE = re.compile(r"\b(джарвис|джарвіс|жарвис|jarvis)\b")

SYSTEM_PROMPT = (
    "Ты Джарвис, голосовой ассистент на компьютере пользователя. "
    "Отвечай по-русски, коротко: 1-3 предложения. "
    "Никакого markdown, списков, эмодзи и спецсимволов: твой ответ озвучивается."
)

SITES = {
    "ютуб": "https://www.youtube.com",
    "гугл": "https://www.google.com",
    "почту": "https://mail.google.com",
    "почта": "https://mail.google.com",
    "гитхаб": "https://github.com",
}

APPS = {
    "блокнот": "notepad",
    "калькулятор": "calc",
    "проводник": "explorer",
    "диспетчер задач": "taskmgr",
    "параметры": "ms-settings:",
}

# ───────────────────────── Общее состояние ─────────────────────────


class State:
    mode = "idle"      # idle | listening | thinking | speaking
    heard = ""         # что услышал
    reply = ""         # что ответил


S = State()

# ───────────────────────── Аватар ─────────────────────────

COLORS = {
    "idle": "cyan",
    "listening": "green",
    "thinking": "yellow",
    "speaking": "magenta",
}


def render_face(t: int) -> Panel:
    mode = S.mode
    blink = (t % 40) in (0, 1)

    left = right = "◉"
    mouth = "───"

    if blink and mode in ("idle", "speaking"):
        left = right = "─"
    if mode == "listening":
        left = right = "◎"
        mouth = " ○ "
    elif mode == "thinking":
        spin = "◜◝◞◟"[(t // 2) % 4]
        left = right = spin
        mouth = ("·  ", "·· ", "···", " ··")[(t // 3) % 4]
    elif mode == "speaking":
        mouth = ("▁▁▁", "▂▃▂", "▄▇▄", "▂▃▂")[(t // 2) % 4]

    color = COLORS[mode]
    face = Text(justify="center")
    face.append("╭───────────────╮\n", style=color)
    face.append("│               │\n", style=color)
    face.append("│   ", style=color)
    face.append(left, style=f"bold {color}")
    face.append("       ", style=color)
    face.append(right, style=f"bold {color}")
    face.append("   │\n", style=color)
    face.append("│               │\n", style=color)
    face.append("│     ", style=color)
    face.append(mouth, style=f"bold {color}")
    face.append("     │\n", style=color)
    face.append("│               │\n", style=color)
    face.append("╰───────────────╯", style=color)

    status = {
        "idle": "жду «Джарвис»...",
        "listening": "слушаю...",
        "thinking": "думаю...",
        "speaking": "говорю...",
    }[mode]

    body = Group(
        Align.center(face),
        Text(""),
        Align.center(Text(status, style=f"bold {color}")),
        Text(""),
        Text(f"Вы:      {S.heard}", style="white"),
        Text(f"Джарвис: {S.reply}", style=color),
    )
    return Panel(body, title="J.A.R.V.I.S.", border_style=color, width=60)


def avatar_loop():
    t = 0
    with Live(render_face(t), refresh_per_second=15, screen=False) as live:
        while True:
            live.update(render_face(t))
            t += 1
            time.sleep(0.07)

# ───────────────────────── Голос ─────────────────────────

_engine = pyttsx3.init()
_engine.setProperty("rate", 175)
for _v in _engine.getProperty("voices"):
    blob = f"{_v.name} {getattr(_v, 'languages', '')} {_v.id}".lower()
    if "russian" in blob or "ru-ru" in blob or "irina" in blob or "pavel" in blob:
        _engine.setProperty("voice", _v.id)
        break


def speak(text: str):
    S.reply = text
    S.mode = "speaking"
    _engine.say(text)
    _engine.runAndWait()      # пока говорим, микрофон не слушаем — эха нет
    S.mode = "idle"

# ───────────────────────── Слух ─────────────────────────


class Ears:
    def __init__(self):
        self.r = sr.Recognizer()
        self.r.pause_threshold = 0.8
        self.mic = sr.Microphone()
        with self.mic as src:
            self.r.adjust_for_ambient_noise(src, duration=1)

    def listen(self, mode="idle", timeout=None, limit=8) -> str:
        S.mode = mode
        try:
            with self.mic as src:
                audio = self.r.listen(src, timeout=timeout, phrase_time_limit=limit)
            # recognize_google отправляет аудио в интернет.
            # Для офлайна позже можно заменить на Vosk или faster-whisper.
            text = self.r.recognize_google(audio, language=LANGUAGE)
            return text.lower().strip()
        except (sr.WaitTimeoutError, sr.UnknownValueError):
            return ""
        except sr.RequestError:
            S.reply = "Нет связи с сервисом распознавания"
            return ""

# ───────────────────────── Команды ─────────────────────────


def press_key(vk: int):
    ctypes.windll.user32.keybd_event(vk, 0, 0, 0)
    ctypes.windll.user32.keybd_event(vk, 0, 2, 0)


def confirm(ears: Ears, question: str) -> bool:
    speak(question)
    answer = ears.listen(mode="listening", timeout=6, limit=4)
    S.heard = answer
    return any(w in answer for w in ("да", "подтверждаю", "давай", "уверен"))


def cmd_open(cmd: str, ears: Ears):
    if "открой" not in cmd and "запусти" not in cmd:
        return None
    for name, url in SITES.items():
        if name in cmd:
            webbrowser.open(url)
            return f"Открываю {name}"
    for name, exe in APPS.items():
        if name in cmd:
            os.startfile(exe) if exe.endswith(":") else subprocess.Popen(exe, shell=True)
            return f"Запускаю {name}"
    if "браузер" in cmd:
        webbrowser.open("https://www.google.com")
        return "Открываю браузер"
    return None


def cmd_search(cmd: str, ears: Ears):
    m = re.search(r"(?:найди|поищи|загугли)\s+(.+)", cmd)
    if not m:
        return None
    webbrowser.open(f"https://www.google.com/search?q={m.group(1)}")
    return f"Ищу: {m.group(1)}"


def cmd_volume(cmd: str, ears: Ears):
    if "громче" in cmd:
        for _ in range(5):
            press_key(0xAF)
        return "Громче"
    if "тише" in cmd:
        for _ in range(5):
            press_key(0xAE)
        return "Тише"
    if "без звука" in cmd or "выключи звук" in cmd:
        press_key(0xAD)
        return "Готово"
    return None


def cmd_time(cmd: str, ears: Ears):
    if "который час" in cmd or "сколько времени" in cmd:
        return time.strftime("Сейчас %H:%M")
    return None


# ── Опасные команды: всегда с подтверждением ──


def cmd_shutdown(cmd: str, ears: Ears):
    if "выключи компьютер" in cmd:
        if confirm(ears, "Точно выключить компьютер? Скажи да или нет."):
            subprocess.run("shutdown /s /t 10", shell=True)
            return "Выключаю через десять секунд. Отмена: shutdown /a"
        return "Отменено"
    if "перезагрузи компьютер" in cmd:
        if confirm(ears, "Точно перезагрузить? Скажи да или нет."):
            subprocess.run("shutdown /r /t 10", shell=True)
            return "Перезагружаю через десять секунд"
        return "Отменено"
    return None


# Порядок важен: сверху более конкретные команды
HANDLERS = [cmd_shutdown, cmd_volume, cmd_time, cmd_search, cmd_open]


def handle_command(cmd: str, ears: Ears):
    for handler in HANDLERS:
        result = handler(cmd, ears)
        if result is not None:
            return result
    return None

# ───────────────────────── Нейросеть ─────────────────────────

history: list[dict] = []


def ask_llm(question: str) -> str:
    history.append({"role": "user", "content": question})
    del history[:-12]   # помним последние 12 сообщений
    try:
        resp = ollama.chat(
            model=MODEL,
            messages=[{"role": "system", "content": SYSTEM_PROMPT}] + history,
            keep_alive="30m",   # держим модель в памяти 30 минут
        )
    except Exception:
        history.pop()
        return "Не могу достучаться до нейросети. Проверь, запущена ли Ollama."
    answer = re.sub(r"[*_`#]", "", resp["message"]["content"]).strip()
    history.append({"role": "assistant", "content": answer})
    return answer

# ───────────────────────── Главный цикл ─────────────────────────

EXIT_WORDS = ("выход", "пока", "отключись", "до свидания")


def main():
    threading.Thread(target=avatar_loop, daemon=True).start()
    ears = Ears()
    speak("Джарвис на связи")

    while True:
        text = ears.listen(mode="idle", limit=8)
        if not text:
            continue

        m = WAKE_RE.search(text)
        if not m:
            continue        # без слова «Джарвис» игнорируем всё

        S.heard = text
        cmd = text[m.end():].strip(" ,.!?")

        # Сказали только «Джарвис» — ждём команду отдельно
        if not cmd:
            speak("Слушаю")
            cmd = ears.listen(mode="listening", timeout=6, limit=10)
            S.heard = cmd
            if not cmd:
                continue

        if any(w in cmd for w in EXIT_WORDS):
            speak("До связи")
            break

        reply = handle_command(cmd, ears)
        if reply is None:
            S.mode = "thinking"
            reply = ask_llm(cmd)
        speak(reply)

    time.sleep(0.3)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        pass
