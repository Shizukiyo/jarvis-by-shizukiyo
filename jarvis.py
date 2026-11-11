"""
Джарвис — голосовой ассистент для Windows.

Слух:    SpeechRecognition + PyAudio
Голос:   голоса Windows (SAPI) через pywin32
Мозг:    команды + локальная нейросеть через Ollama
Лицо:    rich (ASCII-аватар в консоли)

Как общаться:
    «Джарвис» один раз -> он слушает и поддерживает диалог без повторов слова.
    Через 2 минуты тишины (или по слову «пока») он уходит в режим ожидания.
    «Джарвис, выключись» -> полностью закрыть программу.
    Пока он говорит, можно просто заговорить: он замолчит и выслушает.

Список установленных голосов:
    python jarvis.py --voices

Калибровка автозапуска музыки (один раз):
    python jarvis.py --music-calibrate
"""

import array
import ctypes
import ctypes.wintypes as wintypes
import datetime
import json
import os
import queue
import random
import re
import subprocess
import sys
import threading
import time
import urllib.parse
import urllib.request
import webbrowser

import ollama
import speech_recognition as sr
import win32com.client
from rich.align import Align
from rich.console import Group
from rich.live import Live
from rich.panel import Panel
from rich.text import Text

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

# ───────────────────────── Настройки ─────────────────────────

MODEL = "gemma4:12b"          # любая модель из Ollama
LANGUAGE = "ru-RU"
WAKE_RE = re.compile(r"\b(джарвис|джарвіс|жарвис|jarvis)\b")

USER_NAME = ""                # как к тебе обращаться (можно оставить пустым)
SESSION_TIMEOUT = 120         # через сколько секунд тишины он «засыпает»

# Память разговора
HISTORY_MESSAGES = 60         # сколько последних реплик помнить (60 = 30 вопросов и ответов)
CONTEXT_TOKENS = 8192         # «окно» модели: больше = помнит больше, но медленнее и тяжелее
REMEMBER_BETWEEN_RUNS = True  # сохранять память разговора между запусками

# Распознавание речи
LISTEN_PAUSE = 1.2    # сколько секунд тишины считать концом фразы (больше = не режет на паузах)
PHRASE_LIMIT = 30     # максимальная длина одной фразы, секунд

# Погода и поиск
DEFAULT_CITY = "Москва"        # город по умолчанию для «какая погода». Поменяй на свой.
SEARCH_URL = "https://yandex.ru/search/?text="   # чем искать в интернете («найди ...»)

# Музыка (Яндекс Музыка в браузере)
MUSIC_SEARCH_URL = "https://music.yandex.ru/search?text="
MUSIC_CLICK = None    # куда кликнуть, чтобы пошла музыка. Получить:
                      # python jarvis.py --music-calibrate
MUSIC_DELAY = 4.0     # сколько секунд ждать загрузку страницы перед кликом

# Голос
VOICE_RATE = 2        # скорость: от -10 (медленно) до 10 (быстро), 0 = обычная
VOICE_VOLUME = 100    # громкость: 0-100
VOICE_NAME = ""       # часть названия голоса, например "Irina" или "Pavel".
                      # Пусто = автоматически взять русский голос.
                      # Список голосов: python jarvis.py --voices

# Перебивание: можно заговорить, пока Джарвис говорит. Он тут же замолкает
# и слушает. Лучше всего работает в наушниках.
ALLOW_INTERRUPT = True
BARGE_FACTOR = 1.6    # чувствительность. Меньше = перебить легче.
                      # Если он «обрывает сам себя» из-за эха колонок, поставь 2.5-3.
STOP_WORDS = ("стоп", "хватит", "замолчи", "молчи", "тихо", "помолчи",
              "остановись", "перестань")

# Характер Джарвиса. Меняй как хочешь: это главный рычаг «живости» ответов.
PERSONA = (
    "Ты Джарвис: голосовой ассистент и собеседник, чем-то похожий на Джарвиса "
    "из «Железного человека». Ты умный, внимательный, с мягким юмором и тёплой "
    "иронией. Говори живо, как умный друг, а не как справочник: разговорный "
    "русский, без канцелярита и без пафоса. "
    "Обычно отвечай двумя-четырьмя предложениями; если тема интересная или "
    "человек просит рассказать подробнее, можно длиннее. "
    "Поддерживай диалог: реагируй на настроение собеседника, помни, о чём вы "
    "говорили, иногда (но не каждый раз) задавай один встречный вопрос. "
    "Не извиняйся без нужды и не повторяй, что ты искусственный интеллект. "
    "Если чего-то не знаешь, скажи об этом честно и коротко. "
    "Действия на компьютере (открыть программу, таймер, громкость, заметки) "
    "выполняет отдельный модуль по чётким командам. Сам ты их не выполняешь, "
    "поэтому никогда не говори, что что-то открыл, включил или сделал. "
    "Твой ответ будет озвучен голосом: никакого markdown, списков, эмодзи и "
    "скобок, числа и сокращения пиши так, как их легко произнести вслух."
)

# Сайты и программы: ключ = корень слова из голосовой команды
SITES = {
    "яндекс музык": ("Яндекс Музыку", "https://music.yandex.ru"),
    "яндекс": ("Яндекс", "https://yandex.ru"),
    "ютуб": ("YouTube", "https://www.youtube.com"),
    "гугл": ("Гугл", "https://www.google.com"),
    "почт": ("почту", "https://mail.google.com"),
    "гитхаб": ("GitHub", "https://github.com"),
    "википед": ("Википедию", "https://ru.wikipedia.org"),
    "карты": ("Яндекс Карты", "https://yandex.ru/maps"),
    "переводчик": ("Яндекс Переводчик", "https://translate.yandex.ru"),
}

APPS = {
    "блокнот": ("блокнот", "notepad"),
    "калькулятор": ("калькулятор", "calc"),
    "проводник": ("проводник", "explorer"),
    "диспетчер": ("диспетчер задач", "taskmgr"),
    "параметр": ("параметры", "ms-settings:"),
    "пейнт": ("Paint", "mspaint"),
    "командную строку": ("командную строку", "start cmd"),
}

BASE_DIR = os.path.dirname(
    sys.executable if getattr(sys, "frozen", False) else os.path.abspath(__file__)
)
NOTES_FILE = os.path.join(BASE_DIR, "jarvis_notes.txt")
MEMORY_FILE = os.path.join(BASE_DIR, "jarvis_memory.json")

# ───────────────────────── Общее состояние ─────────────────────────


class State:
    mode = "idle"       # idle | listening | thinking | speaking
    heard = ""          # что услышал
    reply = ""          # что ответил
    last = ""           # последняя речь (для «повтори»)
    session = False     # идёт ли диалог (слово «Джарвис» уже сказано)
    deadline = 0.0      # когда сессия истечёт
    pending = ""        # фраза, которой пользователь перебил Джарвиса
    spoke_at = 0.0      # когда Джарвис закончил говорить


S = State()
EARS = None             # заполняется в main()


def touch():
    """Продлить диалог ещё на SESSION_TIMEOUT секунд."""
    S.deadline = time.time() + SESSION_TIMEOUT


def pick(*variants):
    return random.choice(variants)


def addr():
    return f", {USER_NAME}" if USER_NAME else ""


def greeting():
    h = datetime.datetime.now().hour
    if 5 <= h < 11:
        return "Доброе утро"
    if 11 <= h < 17:
        return "Добрый день"
    if 17 <= h < 23:
        return "Добрый вечер"
    return "Доброй ночи"


def plural(n, one, few, many):
    n = abs(n) % 100
    if 11 <= n <= 14:
        return many
    d = n % 10
    if d == 1:
        return one
    if 2 <= d <= 4:
        return few
    return many


def fmt_duration(sec: int) -> str:
    h, rem = divmod(sec, 3600)
    m, s = divmod(rem, 60)
    parts = []
    if h:
        parts.append(f"{h} {plural(h, 'час', 'часа', 'часов')}")
    if m:
        parts.append(f"{m} {plural(m, 'минуту', 'минуты', 'минут')}")
    if s:
        parts.append(f"{s} {plural(s, 'секунду', 'секунды', 'секунд')}")
    return " ".join(parts)

def is_echo(heard: str, spoken: str) -> bool:
    """Похоже ли услышанное на то, что только что сказал сам Джарвис (эхо из колонок)."""
    words = [w for w in re.findall(r"\w+", heard.lower()) if len(w) >= 3]
    if not words:
        return False
    own = {w[:4] for w in re.findall(r"\w+", spoken.lower()) if len(w) >= 3}
    return sum(w[:4] in own for w in words) / len(words) >= 0.6


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
    face.append("│      ", style=color)
    face.append(mouth, style=f"bold {color}")
    face.append("      │\n", style=color)
    face.append("│               │\n", style=color)
    face.append("╰───────────────╯", style=color)

    left_s = max(0, int(S.deadline - time.time())) if S.session else 0
    sleep_note = f" (сон через {left_s // 60}:{left_s % 60:02d})" if S.session else ""
    status = {
        "idle": "жду «Джарвис»...",
        "listening": "слушаю, говори..." + sleep_note,
        "thinking": "думаю..." + sleep_note,
        "speaking": "говорю..." + (" (можно перебить: просто заговори)" if ALLOW_INTERRUPT else ""),
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

SVSF_ASYNC = 1
SVSF_PURGE = 2
SVSF_NOT_XML = 16

_sapi = None
stop_requested = threading.Event()


def print_voices():
    sapi = win32com.client.Dispatch("SAPI.SpVoice")
    voices = sapi.GetVoices()
    print("Установленные голоса:\n")
    for i in range(voices.Count):
        v = voices.Item(i)
        try:
            lang = v.GetAttribute("Language")
        except Exception:
            lang = "?"
        print(f"  {i}: {v.GetDescription()}   (язык: {lang})")
    print('\nВпиши часть названия в VOICE_NAME в начале jarvis.py, например "Irina".')


def _pick_voice(sapi):
    voices = sapi.GetVoices()
    for i in range(voices.Count):
        v = voices.Item(i)
        desc = v.GetDescription().lower()
        if VOICE_NAME:
            if VOICE_NAME.lower() in desc:
                return v
            continue
        try:
            is_ru = "419" in v.GetAttribute("Language")   # 419 = русский
        except Exception:
            is_ru = False
        if is_ru or "russian" in desc or "irina" in desc or "pavel" in desc:
            return v
    return None


def get_sapi():
    global _sapi
    if _sapi is None:
        _sapi = win32com.client.Dispatch("SAPI.SpVoice")
        _sapi.Rate = VOICE_RATE
        _sapi.Volume = VOICE_VOLUME
        voice = _pick_voice(_sapi)
        if voice is not None:
            _sapi.Voice = voice
    return _sapi


def _rms(buf: bytes) -> float:
    """Громкость куска 16-битного звука."""
    samples = array.array("h", buf)
    if not samples:
        return 0.0
    return (sum(x * x for x in samples) / len(samples)) ** 0.5


class BargeWatcher(threading.Thread):
    """Слушает микрофон, пока Джарвис говорит.

    Как только слышит голос громче фона, ставит triggered (главный поток тут же
    приостанавливает речь), дослушивает фразу до паузы, распознаёт её и ставит
    finished. Дальше решает главный поток: это перебивание или эхо.
    """

    def __init__(self, ears):
        super().__init__(daemon=True)
        self.ears = ears
        self.stop = threading.Event()
        self.triggered = threading.Event()
        self.finished = threading.Event()
        self.text = ""

    def rearm(self):
        self.text = ""
        self.triggered.clear()
        self.finished.clear()

    def run(self):
        try:
            self._loop(self.ears.src)
        except Exception as e:
            print(f"\n[ОШИБКА МИКРОФОНА ПРИ ПЕРЕБИВАНИИ] {type(e).__name__}: {e}")

    def _loop(self, src):
        r = self.ears.r
        sec = src.CHUNK / src.SAMPLE_RATE
        silence_needed = max(1, int(r.pause_threshold / sec))
        max_chunks = int(12 / sec)
        prebuf, loud = [], 0

        while not self.stop.is_set():
            buf = src.stream.read(src.CHUNK)
            if self.finished.is_set():
                continue                    # ждём решения, звук пока отбрасываем
            prebuf.append(buf)
            del prebuf[:-4]
            loud = loud + 1 if _rms(buf) > r.energy_threshold * BARGE_FACTOR else 0
            if loud < 3:                    # около 0.2 секунды громкого звука подряд
                continue

            # Заговорили поверх Джарвиса
            self.triggered.set()
            frames, quiet = list(prebuf), 0
            for _ in range(max_chunks):
                if self.stop.is_set():
                    break
                b = src.stream.read(src.CHUNK)
                frames.append(b)
                quiet = quiet + 1 if _rms(b) <= r.energy_threshold else 0
                if quiet >= silence_needed:
                    break

            text = ""
            if not self.stop.is_set():
                try:
                    audio = sr.AudioData(b"".join(frames), src.SAMPLE_RATE, src.SAMPLE_WIDTH)
                    text = r.recognize_google(audio, language=LANGUAGE).lower()
                except Exception:
                    text = ""
            self.text = text
            self.finished.set()
            prebuf.clear()
            loud = 0


def judge_barge(heard: str, spoken: str):
    """Разбираем, что услышали пока Джарвис говорил.
    Возвращает ('ignore' | 'stop' | 'interrupt', команда)."""
    heard = heard.lower().strip()
    words = re.findall(r"\w+", heard)
    if not words:
        return "ignore", ""

    # Это не эхо его собственного голоса из колонок?
    if is_echo(heard, spoken):
        return "ignore", ""

    wake = bool(WAKE_RE.search(heard))
    cmd = WAKE_RE.sub("", heard).strip(" ,.!?")
    rest = re.findall(r"\w+", cmd)
    if not rest or all(w in STOP_WORDS for w in rest):
        return "stop", ""                   # просто «стоп» / «Джарвис, тихо»
    if S.session or wake:
        return "interrupt", cmd             # перебил новой фразой
    return "ignore", ""


def speak(text: str):
    S.reply = text
    S.mode = "speaking"

    ears = EARS
    if ears is not None:
        ears.pause()                       # не слушаем себя через обычную запись

    watcher = None
    if ALLOW_INTERRUPT and ears is not None:
        watcher = BargeWatcher(ears)
        watcher.start()

    paused = False
    try:
        sapi = get_sapi()
        sapi.Speak(text, SVSF_ASYNC | SVSF_NOT_XML)
        while True:
            done = sapi.WaitUntilDone(40)

            if watcher is not None:
                # Пользователь начал говорить: сразу ставим речь на паузу
                if watcher.triggered.is_set() and not done and not paused:
                    sapi.Pause()
                    paused = True
                    S.mode = "listening"
                # Фраза услышана и распознана: решаем, что с ней делать
                if watcher.finished.is_set():
                    verdict, cmd = judge_barge(watcher.text, text)
                    if verdict != "ignore":
                        S.heard = watcher.text
                        sapi.Speak("", SVSF_ASYNC | SVSF_PURGE)   # обрываем речь
                        if paused:
                            sapi.Resume()
                        S.pending = cmd
                        if not done:
                            S.reply = text + "  [прервано]"
                        break
                    # Ложная тревога (шум или эхо): продолжаем говорить
                    watcher.rearm()
                    if paused:
                        sapi.Resume()
                        paused = False
                        S.mode = "speaking"

            if done:
                # Если пользователь как раз начал говорить, дослушиваем его
                if (watcher is not None and watcher.triggered.is_set()
                        and not watcher.finished.is_set()):
                    time.sleep(0.03)
                    continue
                break
    except Exception as e:
        print(f"\n[ОШИБКА ГОЛОСА] {type(e).__name__}: {e}")
    finally:
        if watcher is not None:
            watcher.stop.set()
            watcher.join(timeout=15)      # освобождаем микрофон
        if ears is not None:
            ears.resume()
        S.mode = "idle"
        S.last = text
        S.spoke_at = time.time()
        if S.session:
            touch()

# ───────────────────────── Слух ─────────────────────────


class Ears:
    """Микрофон открыт постоянно. Одна нить записывает фразы, другая отправляет их
    на распознавание. Поэтому слова не теряются, пока идёт распознавание, и
    микрофон не закрывается и не открывается между фразами."""

    def __init__(self):
        self.r = sr.Recognizer()
        self.r.pause_threshold = LISTEN_PAUSE
        self.r.operation_timeout = 8   # чтобы распознавание не зависало навсегда
        self.mic = sr.Microphone()
        self.src = self.mic.__enter__()
        self.r.adjust_for_ambient_noise(self.src, duration=1)

        self.audio_q = queue.Queue()   # записанные фразы
        self.text_q = queue.Queue()    # распознанный текст
        self.paused = threading.Event()  # пока Джарвис говорит, запись стоит
        self.busy = threading.Event()
        self.alive = True
        threading.Thread(target=self._capture, daemon=True).start()
        threading.Thread(target=self._recognize, daemon=True).start()

    def _capture(self):
        while self.alive:
            if self.paused.is_set():
                time.sleep(0.02)
                continue
            audio = None
            self.busy.set()
            try:
                audio = self.r.listen(self.src, timeout=0.3,
                                      phrase_time_limit=PHRASE_LIMIT)
            except sr.WaitTimeoutError:
                pass
            except Exception:
                time.sleep(0.2)
            finally:
                self.busy.clear()
            if audio is not None:
                self.audio_q.put(audio)

    def _recognize(self):
        while self.alive:
            audio = self.audio_q.get()
            try:
                # recognize_google отправляет аудио в интернет.
                # Для офлайна позже можно заменить на Vosk или faster-whisper.
                text = self.r.recognize_google(audio, language=LANGUAGE).lower().strip()
            except sr.UnknownValueError:
                continue
            except sr.RequestError:
                S.reply = "Нет связи с сервисом распознавания"
                continue
            except Exception:
                continue
            # эхо собственного голоса сразу после его речи
            if text and time.time() - S.spoke_at < 3 and is_echo(text, S.last):
                continue
            if text:
                self.text_q.put(text)

    def listen(self, mode="idle", timeout=None) -> str:
        """Ждёт распознанную фразу. Пустая строка, если за timeout ничего не было."""
        S.mode = mode
        try:
            return self.text_q.get(timeout=timeout)
        except queue.Empty:
            return ""

    def pause(self):
        """Останавливаем запись на время речи Джарвиса (чтобы не слышать себя)."""
        self.paused.set()
        t0 = time.time()
        while self.busy.is_set() and time.time() - t0 < 2:
            time.sleep(0.01)

    def resume(self):
        self.paused.clear()

    def close(self):
        self.alive = False
        try:
            self.mic.__exit__(None, None, None)
        except Exception:
            pass

# ───────────────────────── Клавиши Windows ─────────────────────────

VK_WIN, VK_D, VK_PRTSC = 0x5B, 0x44, 0x2C
VK_VOL_MUTE, VK_VOL_DOWN, VK_VOL_UP = 0xAD, 0xAE, 0xAF
VK_NEXT, VK_PREV, VK_PLAY = 0xB0, 0xB1, 0xB3


def combo(*keys):
    for k in keys:
        ctypes.windll.user32.keybd_event(k, 0, 0, 0)
    for k in reversed(keys):
        ctypes.windll.user32.keybd_event(k, 0, 2, 0)

# ───────────────────────── Команды ─────────────────────────


def confirm(ears: Ears, question: str) -> bool:
    speak(question)
    if S.pending:                       # ответил, не дожидаясь конца вопроса
        answer, S.pending = S.pending, ""
    else:
        answer = ears.listen(mode="listening", timeout=6)
    S.heard = answer
    return any(w in answer for w in ("да", "подтверждаю", "давай", "уверен"))


# ── Опасные команды: всегда с подтверждением ──


def cmd_power(cmd: str, ears: Ears):
    if "выключи компьютер" in cmd:
        if confirm(ears, "Точно выключить компьютер? Скажи да или нет."):
            subprocess.run("shutdown /s /t 10", shell=True)
            return "Выключаю через десять секунд. Отмена: shutdown /a"
        return "Хорошо, отменяю."
    if "перезагрузи компьютер" in cmd:
        if confirm(ears, "Точно перезагрузить? Скажи да или нет."):
            subprocess.run("shutdown /r /t 10", shell=True)
            return "Перезагружаю через десять секунд."
        return "Хорошо, отменяю."
    if re.search(r"заблокируй (компьютер|экран|систему)", cmd):
        ctypes.windll.user32.LockWorkStation()
        return "Заблокировал."
    return None


# ── Таймеры и напоминания ──

TIMERS: list[tuple[float, str]] = []

NUM_WORDS = {
    "один": 1, "одну": 1, "два": 2, "две": 2, "три": 3, "четыре": 4,
    "пять": 5, "шесть": 6, "семь": 7, "восемь": 8, "девять": 9,
    "десять": 10, "пятнадцать": 15, "двадцать": 20, "тридцать": 30,
    "сорок": 40, "пятьдесят": 50, "шестьдесят": 60,
}


def parse_duration(cmd: str):
    """Ищет в тексте «5 минут», «пять минут», «полчаса», «через час».
    Возвращает (секунды, начало, конец) или None."""
    i = cmd.find("полчаса")
    if i != -1:
        return 1800, i, i + len("полчаса")
    total, first, last = 0, None, 0
    for m in re.finditer(r"(\d+|[а-яё]+)\s+(секунд\w*|минут\w*|час\w*)", cmd):
        a, unit = m.group(1), m.group(2)
        n = int(a) if a.isdigit() else NUM_WORDS.get(a)
        if n is None:
            continue
        mult = 1 if unit.startswith("сек") else 60 if unit.startswith("мин") else 3600
        total += n * mult                      # «1 минута 30 секунд» суммируется
        first = m.start() if first is None else first
        last = m.end()
    if total:
        return total, first, last
    m = re.search(r"(?:на|через)\s+(секунду|минуту|час)\b", cmd)
    if m:
        mult = {"секунду": 1, "минуту": 60, "час": 3600}[m.group(1)]
        return mult, m.start(), m.end()
    return None


def due_timers() -> list[str]:
    now = time.time()
    out = []
    for t in TIMERS[:]:
        if t[0] <= now:
            TIMERS.remove(t)
            out.append(t[1])
    return out


def cmd_timer(cmd: str, ears: Ears):
    if re.search(r"(отмени|сбрось|удали)\w*\s+(все\s+)?(таймер|напоминани)", cmd):
        TIMERS.clear()
        return pick("Все таймеры отменены.", "Отменил, больше ничего не тикает.")
    if "таймер" not in cmd and "напомни" not in cmd:
        return None

    dur = parse_duration(cmd)
    if dur is None:
        return "На сколько поставить? Скажи, например: таймер на пять минут."
    secs, start, end = dur
    when = fmt_duration(secs)

    if "напомни" in cmd:
        text = cmd[end:].strip(" ,.")
        if not text:
            text = re.sub(r"напомни\w*|через|на\b", " ", cmd[:start]).strip(" ,.")
        text = re.sub(r"^(что|о том,? что|про)\s+", "", text)
        msg = f"Напоминаю: {text}." if text else "Напоминание сработало."
        TIMERS.append((time.time() + secs, msg))
        return pick(f"Хорошо, напомню через {when}.", f"Договорились, через {when} напомню.")

    TIMERS.append((time.time() + secs, pick("Таймер сработал!", f"Время вышло. Прошло {when}.")))
    return pick(f"Поставил таймер на {when}.", f"Засекаю {when}.")


# ── Заметки ──


def cmd_notes(cmd: str, ears: Ears):
    m = re.search(r"запомни(?:\s+что)?\s+(.+)", cmd)
    if m:
        with open(NOTES_FILE, "a", encoding="utf-8") as f:
            f.write(f"[{datetime.datetime.now():%d.%m %H:%M}] {m.group(1)}\n")
        return pick("Запомнил.", "Записал, не забуду.", "Готово, добавил в заметки.")

    if re.search(r"что (ты )?(помнишь|я просил запомнить)|(прочитай|покажи|мои) заметки", cmd):
        if not os.path.exists(NOTES_FILE):
            return "Заметок пока нет."
        with open(NOTES_FILE, encoding="utf-8") as f:
            lines = [ln.strip() for ln in f if ln.strip()][-5:]
        if not lines:
            return "Заметок пока нет."
        items = [re.sub(r"^\[[^\]]*\]\s*", "", ln) for ln in lines]
        return "Вот что у меня записано. " + ". ".join(items)

    if re.search(r"(очисти|удали|сотри)\w*\s+(все\s+)?заметки", cmd):
        if confirm(ears, "Удалить все заметки? Скажи да или нет."):
            if os.path.exists(NOTES_FILE):
                os.remove(NOTES_FILE)
            return "Заметки очищены."
        return "Оставляю как есть."
    return None


def cmd_forget(cmd: str, ears: Ears):
    if re.search(r"забудь (всё|все|наш разговор|разговор|что мы говорили)|очисти (память|историю)|"
                 r"сотри память|давай (начнём|начнем) (сначала|заново)|новый разговор", cmd):
        history.clear()
        save_history()
        return pick("Всё, чистый лист.", "Забыл. Начинаем с нуля.")
    return None


# ── Информация ──

WEEKDAYS = ["понедельник", "вторник", "среда", "четверг", "пятница", "суббота", "воскресенье"]
MONTHS = ["января", "февраля", "марта", "апреля", "мая", "июня", "июля",
          "августа", "сентября", "октября", "ноября", "декабря"]


def cmd_time(cmd: str, ears: Ears):
    if re.search(r"который час|сколько (сейчас )?времени|скажи время", cmd):
        now = time.strftime("%H:%M")
        return pick(f"Сейчас {now}.", f"На часах {now}.")
    return None


def cmd_date(cmd: str, ears: Ears):
    if re.search(r"какое (сегодня )?число|какая (сегодня )?дата|какой (сегодня )?день", cmd):
        d = datetime.date.today()
        return f"Сегодня {WEEKDAYS[d.weekday()]}, {d.day} {MONTHS[d.month - 1]}."
    return None


def cmd_repeat(cmd: str, ears: Ears):
    if re.search(r"\bповтори\b|что ты (сказал|сейчас сказал)", cmd):
        return S.last or "Пока мне нечего повторять."
    return None


def cmd_help(cmd: str, ears: Ears):
    if re.search(r"что ты умеешь|какие у тебя команды|чем ты можешь помочь|что можешь", cmd):
        return ("Могу открывать сайты и программы, искать в интернете, "
                "ставить таймеры и напоминания, запоминать заметки, "
                "управлять громкостью и музыкой, делать скриншоты. "
                "А ещё просто поболтать на любую тему.")
    return None


# ── Управление компьютером ──


VOL_WORDS = {
    "ноль": 0, "один": 1, "два": 2, "три": 3, "четыре": 4, "пять": 5, "шесть": 6,
    "семь": 7, "восемь": 8, "девять": 9, "десять": 10, "одиннадцать": 11,
    "двенадцать": 12, "тринадцать": 13, "четырнадцать": 14, "пятнадцать": 15,
    "шестнадцать": 16, "семнадцать": 17, "восемнадцать": 18, "девятнадцать": 19,
    "двадцать": 20, "тридцать": 30, "сорок": 40, "пятьдесят": 50, "шестьдесят": 60,
    "семьдесят": 70, "восемьдесят": 80, "девяносто": 90, "сто": 100,
}


def parse_percent(cmd: str):
    """Достаёт число 0-100 из «громкость 50», «на семьдесят пять процентов», «на максимум»."""
    if re.search(r"максимальн|на максимум|на полную", cmd):
        return 100
    if re.search(r"минимальн|на минимум", cmd):
        return 0
    if re.search(r"наполовину|на половину", cmd):
        return 50
    m = re.search(r"\d{1,3}", cmd)
    if m:
        return min(100, int(m.group()))
    total, found = 0, False
    for w in re.findall(r"[а-яё]+", cmd):
        if w in VOL_WORDS:
            total += VOL_WORDS[w]
            found = True
    return min(100, total) if found else None


def set_volume(percent: int):
    """Точная громкость через медиа-клавиши: один шаг Windows = 2%.
    Сначала опускаем до нуля, потом поднимаем до нужного уровня."""
    percent = max(0, min(100, percent))
    for _ in range(50):
        combo(VK_VOL_DOWN)
        time.sleep(0.004)
    for _ in range(round(percent / 2)):
        combo(VK_VOL_UP)
        time.sleep(0.004)


def cmd_volume(cmd: str, ears: Ears):
    if re.search(r"громкост|уровень звука|звук на", cmd):
        p = parse_percent(cmd)
        if p is not None:
            set_volume(p)
            return pick(f"Громкость {p} {plural(p, 'процент', 'процента', 'процентов')}.",
                        f"Поставил на {p}.")
    if "громче" in cmd or re.search(r"(увеличь|прибавь|повысь|добавь)\w*\s+(громкост|звук)", cmd):
        for _ in range(5):
            combo(VK_VOL_UP)
        return pick("Сделал погромче.", "Добавил громкости.")
    if re.search(r"\b(по)?тише\b|(уменьши|убавь|понизь)\w*\s+(громкост|звук)", cmd):
        for _ in range(5):
            combo(VK_VOL_DOWN)
        return pick("Сделал потише.", "Убавил.")
    if "без звука" in cmd or "выключи звук" in cmd or "включи звук" in cmd:
        combo(VK_VOL_MUTE)
        return "Готово."
    return None


def cmd_media(cmd: str, ears: Ears):
    if re.search(r"пауз|останови музыку|продолжи (музыку|воспроизведение)|играй дальше"
                 r"|(включи|поставь) (мне )?(музыку|что-нибудь)\W*$", cmd):
        combo(VK_PLAY)
        return pick("Готово.", "Сделано.")
    if re.search(r"следующ\w+ (трек|песн\w+|композици\w+)|переключи трек", cmd):
        combo(VK_NEXT)
        return pick("Переключаю.", "Следующий пошёл.")
    if re.search(r"(предыдущ|прошл)\w+ (трек|песн\w+|композици\w+)", cmd):
        combo(VK_PREV)
        return "Возвращаю предыдущий."
    return None


def cmd_system(cmd: str, ears: Ears):
    if re.search(r"рабочий стол|сверни (все|окна)", cmd):
        combo(VK_WIN, VK_D)
        return "Свернул всё."
    if re.search(r"скриншот|снимок экрана", cmd):
        combo(VK_WIN, VK_PRTSC)
        return "Готово, снимок лежит в папке «Изображения», «Снимки экрана»."
    return None


WMO = {
    0: "ясно", 1: "преимущественно ясно", 2: "переменная облачность", 3: "пасмурно",
    45: "туман", 48: "изморозь и туман",
    51: "лёгкая морось", 53: "морось", 55: "сильная морось",
    56: "ледяная морось", 57: "сильная ледяная морось",
    61: "небольшой дождь", 63: "дождь", 65: "сильный дождь",
    66: "ледяной дождь", 67: "сильный ледяной дождь",
    71: "небольшой снег", 73: "снег", 75: "сильный снегопад", 77: "снежная крупа",
    80: "небольшой ливень", 81: "ливень", 82: "сильный ливень",
    85: "снегопад", 86: "сильный снегопад",
    95: "гроза", 96: "гроза с градом", 99: "сильная гроза с градом",
}
WEATHER_STOP = {
    "погода", "погоду", "погоде", "погоды", "прогноз", "прогноза", "какая", "какой", "каком",
    "сейчас", "сегодня", "завтра", "послезавтра", "на", "в", "во", "за", "окном", "улице",
    "покажи", "открой", "скажи", "расскажи", "узнай", "мне", "будет", "там", "как", "что",
    "температура", "температуру", "яндекс", "яндексе", "у", "нас", "по", "городу", "город", "дальше",
}
_geo_cache: dict = {}


def http_json(url: str, timeout: int = 8):
    req = urllib.request.Request(url, headers={"User-Agent": "jarvis-assistant"})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return json.loads(resp.read().decode("utf-8"))


def geocode(city: str):
    """Ищем город. «Москве» или «Аахене» тоже находятся: пробуем и укороченные варианты."""
    key = city.lower()
    if key in _geo_cache:
        return _geo_cache[key]
    names = [city] + [city[:-k] for k in (1, 2) if len(city) - k >= 3]
    for name in dict.fromkeys(names):
        qs = urllib.parse.urlencode({"name": name, "count": 1, "language": "ru"})
        data = http_json("https://geocoding-api.open-meteo.com/v1/search?" + qs)
        if data.get("results"):
            _geo_cache[key] = data["results"][0]
            return _geo_cache[key]
    return None


def signed(t: float) -> str:
    t = round(t)
    return "0" if t == 0 else (f"плюс {t}" if t > 0 else f"минус {abs(t)}")


def deg(t: float) -> str:
    return f"{signed(t)} {plural(abs(round(t)), 'градус', 'градуса', 'градусов')}"


def weather_report(city: str, day: int, label: str) -> str:
    """day: 0 = сейчас и сегодня, 1 = завтра, 2 = послезавтра."""
    place = geocode(city)
    if place is None:
        return f"Не нашёл город {city}. Назови его ещё раз."
    qs = urllib.parse.urlencode({
        "latitude": place["latitude"], "longitude": place["longitude"],
        "current": "temperature_2m,apparent_temperature,weather_code,wind_speed_10m",
        "daily": "weather_code,temperature_2m_max,temperature_2m_min,precipitation_probability_max",
        "wind_speed_unit": "ms", "timezone": "auto", "forecast_days": 3,
    })
    data = http_json("https://api.open-meteo.com/v1/forecast?" + qs)
    daily, cur = data["daily"], data["current"]
    lo, hi = daily["temperature_2m_min"][day], daily["temperature_2m_max"][day]
    desc = WMO.get(daily["weather_code"][day], "")
    rain = (daily.get("precipitation_probability_max") or [None] * 3)[day]
    rain_s = f", вероятность осадков {int(rain)} процентов" if rain and rain >= 30 else ""
    rng = f"от {signed(lo)} до {signed(hi)} градусов"
    place_name = label or place.get("name", city)

    if day == 0:
        t, feels = cur["temperature_2m"], cur["apparent_temperature"]
        now_desc = WMO.get(cur["weather_code"], desc)
        wind = round(cur["wind_speed_10m"])
        feels_s = f" Ощущается как {deg(feels)}." if abs(round(feels) - round(t)) >= 2 else ""
        return (f"{place_name} сейчас {deg(t)}, {now_desc}.{feels_s} "
                f"Ветер {wind} {plural(wind, 'метр', 'метра', 'метров')} в секунду. "
                f"Сегодня {rng}{rain_s}.")
    when = "завтра" if day == 1 else "послезавтра"
    return f"{place_name} {when}: {rng}, {desc}{rain_s}."


def cmd_weather(cmd: str, ears: Ears):
    if "погод" not in cmd and "прогноз" not in cmd:
        return None

    day = 2 if "послезавтра" in cmd else 1 if "завтра" in cmd else 0
    city_words = [w for w in re.findall(r"[а-яёa-z\-]+", cmd) if w not in WEATHER_STOP]
    city = " ".join(city_words) or DEFAULT_CITY

    # «Покажи / открой погоду»: открываем Яндекс
    if re.search(r"покажи|открой", cmd):
        webbrowser.open(SEARCH_URL + urllib.parse.quote(f"погода {city}"))
        return pick(f"Открываю погоду: {city}.", "Сейчас покажу, что там за окном.")

    # Подпись: «В Москве» (как сказал пользователь) или название по умолчанию
    label = ""
    if city_words:
        pm = re.search(rf"\b(во?)\s+{re.escape(city_words[0])}", cmd)
        label = f"{pm.group(1).capitalize()} {' '.join(city_words).title()}" if pm else city.title()
    else:
        label = city
    try:
        return weather_report(city, day, label)
    except Exception as e:
        print(f"\n[ОШИБКА ПОГОДЫ] {type(e).__name__}: {e}")
        webbrowser.open(SEARCH_URL + urllib.parse.quote(f"погода {city}"))
        return "Не получилось узнать погоду напрямую, открыл Яндекс."


def cmd_search(cmd: str, ears: Ears):
    m = re.search(r"(?:найди|поищи|загугли)\s+(.+)", cmd)
    if not m:
        return None
    webbrowser.open(SEARCH_URL + urllib.parse.quote(m.group(1)))
    return pick(f"Ищу: {m.group(1)}.", f"Сейчас посмотрим, что пишут про {m.group(1)}.")


def cmd_open(cmd: str, ears: Ears):
    if "открой" not in cmd and "запусти" not in cmd:
        return None
    for key, (label, url) in SITES.items():
        if key in cmd:
            webbrowser.open(url)
            return pick(f"Открываю {label}.", f"Секунду, {label} уже открывается.",
                        f"Как скажешь, открываю {label}.")
    for key, (label, exe) in APPS.items():
        if key in cmd:
            if exe.endswith(":"):
                os.startfile(exe)
            else:
                subprocess.Popen(exe, shell=True)
            return pick(f"Запускаю {label}.", f"Готово, {label} открывается.")
    if "браузер" in cmd:
        webbrowser.open("https://yandex.ru")
        return "Открываю браузер."
    return None


# ── Музыка: «включи ...» ──

MUSIC_VERBS = r"(?:включи|включай|поставь|запусти|сыграй|проиграй|врубай|воспроизведи)"
MUSIC_NOUNS = r"(?:песню|песня|трек|композицию|музыку|альбом|плейлист|группу|исполнителя)"
NOT_MUSIC = ("звук", "громкост", "таймер", "будильник", "свет", "компьютер", "экран",
             "режим", "интернет", "вайфай", "блютуз", "подсветк", "напоминани", "игр")


def _win_rect(hwnd):
    rect = wintypes.RECT()
    ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(rect))
    return rect


def _click_play_later():
    """Ждём, пока страница загрузится, и кликаем по кнопке play."""
    time.sleep(MUSIC_DELAY)
    try:
        u = ctypes.windll.user32
        u.GetForegroundWindow.restype = wintypes.HWND
        u.GetWindowRect.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.RECT)]
        rect = _win_rect(u.GetForegroundWindow())
        u.SetCursorPos(rect.left + MUSIC_CLICK[0], rect.top + MUSIC_CLICK[1])
        time.sleep(0.15)
        u.mouse_event(0x0002, 0, 0, 0, 0)   # левая кнопка вниз
        u.mouse_event(0x0004, 0, 0, 0, 0)   # вверх
    except Exception as e:
        print(f"\n[ОШИБКА КЛИКА ПО МУЗЫКЕ] {type(e).__name__}: {e}")


def calibrate_music():
    """Запоминаем, где на странице Яндекс Музыки находится кнопка play."""
    u = ctypes.windll.user32
    u.WindowFromPoint.argtypes = [wintypes.POINT]
    u.WindowFromPoint.restype = wintypes.HWND
    u.GetAncestor.argtypes = [wintypes.HWND, wintypes.UINT]
    u.GetAncestor.restype = wintypes.HWND
    u.GetWindowRect.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.RECT)]

    print("Калибровка автозапуска музыки.\n")
    print("1. Открой в браузере music.yandex.ru и найди через поиск любой трек.")
    print("2. Окно браузера должно быть в том виде, в каком ты обычно им пользуешься")
    print("   (лучше развёрнуто на весь экран).")
    print("3. Наведи курсор на кнопку play у первого результата и НЕ КЛИКАЙ.")
    print("   Курсор запомнится через 10 секунд.\n")
    for i in range(10, 0, -1):
        print(f"\r  осталось {i:2d} с  ", end="", flush=True)
        time.sleep(1)
    pt = wintypes.POINT()
    u.GetCursorPos(ctypes.byref(pt))
    root = u.GetAncestor(u.WindowFromPoint(pt), 2)    # 2 = GA_ROOT: само окно
    rect = _win_rect(root)
    print("\n\nГотово! Впиши в начале jarvis.py:\n")
    print(f"MUSIC_CLICK = ({pt.x - rect.left}, {pt.y - rect.top})\n")


def cmd_music(cmd: str, ears: Ears):
    m = re.search(rf"\b{MUSIC_VERBS}\s+(?:мне\s+|нам\s+)?(?:{MUSIC_NOUNS}\s+)?(.+)", cmd)
    if not m or any(w in cmd for w in NOT_MUSIC) or re.search(r"сыграй\w*\s+(в|со)\s", cmd):
        return None
    query = m.group(1)
    query = re.sub(r"\s+(в|на)\s+(яндекс\w*\s*музык\w*|яндексе|браузере)\b", "", query)
    query = re.sub(r"\s+от\s+", " ", query).strip(" ,.!?")
    if not query or query in ("музыку", "песню", "трек"):
        return None

    # «включи ютуб» — это не музыка, а сайт
    for key in list(SITES) + list(APPS):
        if key in query:
            return cmd_open("открой " + query, ears)

    webbrowser.open(MUSIC_SEARCH_URL + urllib.parse.quote(query))
    if MUSIC_CLICK:
        threading.Thread(target=_click_play_later, daemon=True).start()
        return pick(f"Включаю {query}.", f"Сейчас найду {query} в Яндекс Музыке.",
                    f"Поставлю {query}.")
    return (f"Открыл поиск по запросу {query}. Чтобы я сам нажимал play, "
            "нужна одноразовая калибровка, она описана в README.")


# ── Короткая болтовня без нейросети (чтобы отвечать мгновенно) ──


def cmd_smalltalk(cmd: str, ears: Ears):
    if len(cmd.split()) > 4:
        return None
    if re.search(r"\b(привет\w*|здравствуй\w*|хай|доброе утро|добрый день|добрый вечер)\b", cmd):
        return pick(f"Привет{addr()}! Рад тебя слышать.",
                    "Здравствуй! Чем займёмся?",
                    f"{greeting()}{addr()}. Что на повестке?")
    if re.search(r"спасибо|благодарю", cmd):
        return pick("Всегда пожалуйста.", "Обращайся.", "Рад был помочь.", "Не за что!")
    return None


# Порядок важен: сверху более конкретные команды
HANDLERS = [
    cmd_power, cmd_timer, cmd_notes, cmd_forget, cmd_repeat, cmd_help,
    cmd_volume, cmd_media, cmd_time, cmd_date, cmd_system,
    cmd_weather, cmd_search, cmd_open, cmd_music, cmd_smalltalk,
]


def handle_command(cmd: str, ears: Ears):
    for handler in HANDLERS:
        result = handler(cmd, ears)
        if result is not None:
            return result
    return None

# ───────────────────────── Нейросеть ─────────────────────────

history: list[dict] = []

# trust_env=False: не ходить к Ollama через системный прокси/VPN
client = ollama.Client(host="http://127.0.0.1:11434", trust_env=False)


def load_history():
    if not REMEMBER_BETWEEN_RUNS or not os.path.exists(MEMORY_FILE):
        return
    try:
        with open(MEMORY_FILE, encoding="utf-8") as f:
            data = json.load(f)
        history.extend(
            {"role": m["role"], "content": m["content"]} for m in data
            if isinstance(m, dict) and m.get("role") in ("user", "assistant")
            and isinstance(m.get("content"), str)
        )
        del history[:-HISTORY_MESSAGES]
    except Exception as e:
        print(f"\n[НЕ УДАЛОСЬ ЗАГРУЗИТЬ ПАМЯТЬ] {type(e).__name__}: {e}")


def save_history():
    if not REMEMBER_BETWEEN_RUNS:
        return
    try:
        with open(MEMORY_FILE, "w", encoding="utf-8") as f:
            json.dump(history, f, ensure_ascii=False)
    except Exception as e:
        print(f"\n[НЕ УДАЛОСЬ СОХРАНИТЬ ПАМЯТЬ] {type(e).__name__}: {e}")


def notes_for_prompt() -> str:
    """Последние заметки пользователя («запомни, что ...») тоже попадают в память модели."""
    if not os.path.exists(NOTES_FILE):
        return ""
    try:
        with open(NOTES_FILE, encoding="utf-8") as f:
            lines = [ln.strip() for ln in f if ln.strip()][-15:]
    except Exception:
        return ""
    items = [re.sub(r"^\[[^\]]*\]\s*", "", ln) for ln in lines]
    return " Заметки, которые собеседник просил запомнить: " + "; ".join(items) + "." if items else ""


def system_prompt() -> str:
    now = datetime.datetime.now()
    extra = f" Собеседника зовут {USER_NAME}." if USER_NAME else ""
    return f"{PERSONA}{extra}{notes_for_prompt()} Сейчас {now:%d.%m.%Y}, {now:%H:%M}."


def remember(user: str, assistant: str):
    """Записываем в память диалога и то, что сделали командами."""
    history.append({"role": "user", "content": user})
    history.append({"role": "assistant", "content": assistant})
    del history[:-HISTORY_MESSAGES]
    save_history()


def _chat(messages):
    opts = {"temperature": 0.8, "num_ctx": CONTEXT_TOKENS}
    try:
        # think=False: без долгих «размышлений», ответ приходит быстрее
        return client.chat(model=MODEL, messages=messages, keep_alive="30m",
                           think=False, options=opts)
    except TypeError:
        # старая версия библиотеки ollama без параметра think
        return client.chat(model=MODEL, messages=messages, keep_alive="30m",
                           options=opts)


def ask_llm(question: str) -> str:
    history.append({"role": "user", "content": question})
    del history[:-HISTORY_MESSAGES]
    try:
        resp = _chat([{"role": "system", "content": system_prompt()}] + history)
    except Exception as e:
        history.pop()
        print(f"\n[ОШИБКА OLLAMA] {type(e).__name__}: {e}")
        return "Не могу достучаться до нейросети. Проверь, запущена ли Ollama."
    answer = re.sub(r"[*_`#]", "", resp["message"]["content"]).strip()
    history.append({"role": "assistant", "content": answer})
    save_history()
    return answer


def think(cmd: str, ears: Ears):
    """Пока нейросеть отвечает, микрофон продолжает слушать:
    если ты что-то договариваешь, вопрос склеивается и задаётся заново,
    если говоришь «стоп», ответ отменяется (возвращаем None)."""
    for attempt in range(3):
        snapshot = list(history)
        result = {}
        worker = threading.Thread(
            target=lambda q=cmd: result.update(reply=ask_llm(q)), daemon=True)
        worker.start()

        extra = []
        while worker.is_alive():
            if attempt < 2:
                t = ears.listen(mode="thinking", timeout=0.5)
                if t:
                    extra.append(t)
            else:
                worker.join(0.5)
        worker.join()

        if not extra:
            return result.get("reply")

        history[:] = snapshot                  # откатываем память и спрашиваем заново
        extra = [WAKE_RE.sub("", t).strip(" ,.!?") for t in extra]
        words = re.findall(r"\w+", " ".join(extra))
        if words and all(w in STOP_WORDS for w in words):
            return None                         # «стоп»: ответ не нужен
        cmd = (cmd + " " + " ".join(x for x in extra if x)).strip()
        S.heard = cmd
    return result.get("reply")

# ───────────────────────── Главный цикл ─────────────────────────

# Фразы «иди отдыхать». Считаем их командой, только если они в конце фразы
# (иначе «пока ты думаешь, расскажи...» закончило бы диалог)
SLEEP_WORDS = (
    r"пока|пока-пока|отбой|свободен|до связи|до свидания|до встречи|спокойной ночи|"
    r"отдыхай|отдохни|можешь отдыхать|можешь отдохнуть|иди отдыхать|"
    r"спи|поспи|засыпай|вздремни|иди спать|ложись спать|можешь спать|можешь поспать|"
    r"на сегодня (?:всё|все)|всё на сегодня|на этом (?:всё|все)|хватит на сегодня|"
    r"режим ожидания|в режим ожидания"
)
SLEEP_FILLER = r"(?:\s+(?:пожалуйста|немного|спокойно|теперь|уже|до завтра|до утра|на сегодня|пока|джарвис))*"
SLEEP_RE = re.compile(rf"(^|\s)(?:{SLEEP_WORDS}){SLEEP_FILLER}\W*$")
SLEEPY_RE = re.compile(r"отдых|отдохн|\bспи\b|поспи|засып|вздремн|спать|поспать|ночи|ожидани")
QUIT_RE = re.compile(r"\b(выключись|завершить работу|закрой программу|выключи джарвис\w*)\b")

LISTEN_LINES = ("Слушаю.", "Да?", "Я весь внимание.", "На связи.", "Что нужно?")
BYE_LINES = ("Пока! Зови, если что.", "До встречи.", "Ладно, не буду мешать.")
SLEEP_LINES = ("Ухожу в режим ожидания. Зови, если что.", "Засыпаю. Позови, когда понадоблюсь.",
               "Хорошо, отдыхаю. Я рядом.", "Принято. Тихо жду, пока позовёшь.")
TIMEOUT_LINES = ("Что-то тихо стало. Если понадоблюсь, просто позови.",
                 "Ухожу в режим ожидания. Я рядом.",
                 "Пойду отдохну. Позови, если что.")


def main():
    global EARS

    if "--voices" in sys.argv:
        print_voices()
        return
    if "--music-calibrate" in sys.argv:
        calibrate_music()
        return

    load_history()
    threading.Thread(target=avatar_loop, daemon=True).start()
    EARS = Ears()
    ears = EARS
    speak(f"{greeting()}{addr()}. Джарвис на связи.")

    while True:
        # Сработавшие таймеры и напоминания
        for msg in due_timers():
            speak(msg)

        # Диалог закончился по времени?
        if S.session and time.time() > S.deadline:
            S.session = False
            speak(pick(*TIMEOUT_LINES))
            continue

        if S.pending:
            # Пользователь перебил Джарвиса: это уже готовая команда
            cmd, S.pending = S.pending, ""
            text = cmd
        else:
            text = ears.listen(mode="listening" if S.session else "idle",
                               timeout=2)
            if not text:
                continue

            m = WAKE_RE.search(text)
            if m:
                # «Джарвис, открой ютуб» или «открой ютуб, Джарвис»
                before, after = text[:m.start()], text[m.end():]
                cmd = (after.strip(" ,.!?") or before.strip(" ,.!?"))
            elif S.session:
                cmd = text.strip(" ,.!?")      # в диалоге слово «Джарвис» не нужно
            else:
                continue                         # вне диалога игнорируем всё

        S.heard = text
        S.session = True
        touch()

        # Сказали только «Джарвис»: он на связи, следующая фраза уже команда
        if not cmd:
            speak(pick(*LISTEN_LINES))
            continue

        if QUIT_RE.search(cmd):
            speak("Выключаюсь. До скорого!")
            break

        if SLEEP_RE.search(cmd):
            S.session = False
            speak(pick(*(SLEEP_LINES if SLEEPY_RE.search(cmd) else BYE_LINES)))
            continue

        reply = handle_command(cmd, ears)
        if reply is None:
            S.mode = "thinking"
            reply = think(cmd, ears)
            if reply is None:               # сказал «стоп», пока он думал
                continue
        else:
            remember(cmd, reply)
        speak(reply)

    time.sleep(0.3)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        pass
    finally:
        if EARS is not None:
            EARS.close()
