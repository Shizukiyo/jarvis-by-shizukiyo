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
"""

import array
import ctypes
import datetime
import os
import random
import re
import subprocess
import sys
import threading
import time
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
    "ютуб": ("YouTube", "https://www.youtube.com"),
    "гугл": ("Гугл", "https://www.google.com"),
    "почт": ("почту", "https://mail.google.com"),
    "гитхаб": ("GitHub", "https://github.com"),
    "википед": ("Википедию", "https://ru.wikipedia.org"),
    "карты": ("карты", "https://www.google.com/maps"),
    "переводчик": ("переводчик", "https://translate.google.com"),
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

# ───────────────────────── Общее состояние ─────────────────────────


class State:
    mode = "idle"       # idle | listening | thinking | speaking
    heard = ""          # что услышал
    reply = ""          # что ответил
    last = ""           # последняя речь (для «повтори»)
    session = False     # идёт ли диалог (слово «Джарвис» уже сказано)
    deadline = 0.0      # когда сессия истечёт
    pending = ""        # фраза, которой пользователь перебил Джарвиса


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
            with self.ears.mic as src:
                self._loop(src)
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
    own = {w[:4] for w in re.findall(r"\w+", spoken.lower()) if len(w) >= 3}
    long_words = [w for w in words if len(w) >= 3]
    if long_words and sum(w[:4] in own for w in long_words) / len(long_words) >= 0.6:
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

    watcher = None
    if ALLOW_INTERRUPT and EARS is not None:
        watcher = BargeWatcher(EARS)
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
        S.mode = "idle"
        S.last = text
        if S.session:
            touch()

# ───────────────────────── Слух ─────────────────────────


class Ears:
    def __init__(self):
        self.r = sr.Recognizer()
        self.r.pause_threshold = 0.8
        self.r.operation_timeout = 8   # чтобы распознавание не зависало навсегда
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
        answer = ears.listen(mode="listening", timeout=6, limit=4)
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


def cmd_volume(cmd: str, ears: Ears):
    if "громче" in cmd:
        for _ in range(5):
            combo(VK_VOL_UP)
        return pick("Сделал погромче.", "Добавил громкости.")
    if re.search(r"\b(по)?тише\b", cmd):
        for _ in range(5):
            combo(VK_VOL_DOWN)
        return pick("Сделал потише.", "Убавил.")
    if "без звука" in cmd or "выключи звук" in cmd or "включи звук" in cmd:
        combo(VK_VOL_MUTE)
        return "Готово."
    return None


def cmd_media(cmd: str, ears: Ears):
    if re.search(r"пауз|останови музыку|продолжи (музыку|воспроизведение)|играй дальше", cmd):
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


def cmd_weather(cmd: str, ears: Ears):
    if "погод" in cmd:
        webbrowser.open("https://www.google.com/search?q=погода")
        return pick("Открываю прогноз погоды.", "Сейчас посмотрим, что там за окном.")
    return None


def cmd_search(cmd: str, ears: Ears):
    m = re.search(r"(?:найди|поищи|загугли)\s+(.+)", cmd)
    if not m:
        return None
    webbrowser.open(f"https://www.google.com/search?q={m.group(1)}")
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
        webbrowser.open("https://www.google.com")
        return "Открываю браузер."
    return None


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
    cmd_power, cmd_timer, cmd_notes, cmd_repeat, cmd_help,
    cmd_volume, cmd_media, cmd_time, cmd_date, cmd_system,
    cmd_weather, cmd_search, cmd_open, cmd_smalltalk,
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


def system_prompt() -> str:
    now = datetime.datetime.now()
    extra = f" Собеседника зовут {USER_NAME}." if USER_NAME else ""
    return f"{PERSONA}{extra} Сейчас {now:%d.%m.%Y}, {now:%H:%M}."


def remember(user: str, assistant: str):
    """Записываем в память диалога и то, что сделали командами."""
    history.append({"role": "user", "content": user})
    history.append({"role": "assistant", "content": assistant})
    del history[:-20]


def _chat(messages):
    opts = {"temperature": 0.8}
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
    del history[:-20]   # помним последние 20 сообщений
    try:
        resp = _chat([{"role": "system", "content": system_prompt()}] + history)
    except Exception as e:
        history.pop()
        print(f"\n[ОШИБКА OLLAMA] {type(e).__name__}: {e}")
        return "Не могу достучаться до нейросети. Проверь, запущена ли Ollama."
    answer = re.sub(r"[*_`#]", "", resp["message"]["content"]).strip()
    history.append({"role": "assistant", "content": answer})
    return answer

# ───────────────────────── Главный цикл ─────────────────────────

# «пока» и подобное считаем прощанием, только если это конец фразы
# (иначе «пока ты думаешь, расскажи...» закончило бы диалог)
SLEEP_RE = re.compile(r"(^|\s)(пока|отбой|свободен|до связи|до свидания|спокойной ночи)\W*$")
QUIT_RE = re.compile(r"\b(выключись|завершить работу|закрой программу|выключи джарвис\w*)\b")

LISTEN_LINES = ("Слушаю.", "Да?", "Я весь внимание.", "На связи.", "Что нужно?")
BYE_LINES = ("Пока! Зови, если что.", "До встречи.", "Отдыхай, я рядом.", "Ладно, не буду мешать.")
TIMEOUT_LINES = ("Что-то тихо стало. Если понадоблюсь, просто позови.",
                 "Ухожу в режим ожидания. Я рядом.",
                 "Пойду отдохну. Позови, если что.")


def main():
    global EARS

    if "--voices" in sys.argv:
        print_voices()
        return

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
                               timeout=2, limit=12)
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
            speak(pick(*BYE_LINES))
            continue

        reply = handle_command(cmd, ears)
        if reply is None:
            S.mode = "thinking"
            reply = ask_llm(cmd)
        else:
            remember(cmd, reply)
        speak(reply)

    time.sleep(0.3)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        pass
