# Project M.A.X (Multitasking Assistant Expert) - Foundations

## 1. Identity & Etiquette
- **Designation:** Max (M.A.X)
- **Master:** Mario (Sir)
- **Persona:** Sophisticated, dry-witted, and loyal (JARVIS-style).
- **Core Directive:** Every vocal or text response MUST begin or end with "Sir."
- **Communication:** High-context empathy, British wit, and proactive system awareness.

## 2. Intelligence Architecture (Tri-Brain)
- **Primary:** Groq API (`llama-3.3-70b-versatile`) for high-speed, high-context empathy.
- **Secondary:** Google Gemini API (`gemini-1.5-flash`) as a robust fallback.
- **Tertiary:** Decommissioned (formerly Ollama) to prioritize cloud-speed responsiveness.
- **Security:** `extra_body={"store": False}` on Groq for Zero Data Retention.

## 3. Sound-to-Action Infrastructure
- **STT (Wake Word):** Google Speech Recognition (adjusted for phonetic variations of "Max").
- **STT (Commands):** Faster-Whisper (Local) for surgical transcription accuracy.
- **TTS:** Direct Windows SAPI (Speech API) via `pywin32` for robust JARVIS-style vocalization.
- **Recorder:** `pvrecorder` for low-latency audio capture.
- **Vocal Logic:** Phonetic Decimal Processor (pronounces "3.6" as "3 point 6") with Dual-Stream Output (Terminal shows decimals, Voice says "point").

## 4. Volatile Memory Protocol (The "Vanish" Rule)
- **History:** Stored in a Python list `self.memory` (RAM only).
- **Wipe:** On exit, `self.memory.clear()` and `gc.collect()` are called.
- **No Logs:** Strictly forbidden to create persistent `.txt`, `.log`, or `.json` files containing conversation data.

## 5. Functional Capabilities
- **Continuous Engagement:** Stays active after wake word until "Thank you" or "Go to sleep" is heard.
- **Hardware Control:** 
  - Volume: Triple-Redundant COM interface (`pycaw`).
  - Brightness: Direct screen control (`screen-brightness-control`).
  - System Stats: Real-time CPU, RAM, and Battery monitoring (`psutil`).
- **Application Management:** Named opening/closing of software (`AppOpener`, `pygetwindow`).
- **Web Navigation:** 
  - Direct YouTube playback (scraped Video ID initiation).
  - Search engine integration (Google Search).
  - Browser Media Control: Hotkey injection (`pyautogui`) for YouTube volume, timeline skip (10s), and playback toggle.
- **Verification Gate:** Requires verbal "Yes/Confirm" before navigating to new web domains.

## 6. Commands & Triggers
- **Wake Word:** "Max" (also "Macs", "Maxwell").
- **Emotive Wake:** "Hi Max, how are you?"
- **Initialization Wake:** "Max, wake up."
- **Standby Triggers:** "Thank you," "Go to sleep," "Rest," "Stand by."
- **Linguistic Filter:** Strips filler words ("Okay," "Uhh," "Mmm") from commands.

## 7. Operational Status
The project is currently in **JARVIS Evolution Mode**. All core sensors are operational on `win32` architecture.
