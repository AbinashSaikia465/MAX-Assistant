import os
import gc
import sys
import atexit
import groq
from google import genai
from faster_whisper import WhisperModel
import screen_brightness_control as sbc
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume, IMMDeviceEnumerator
import comtypes
from comtypes import CLSCTX_ALL
from ctypes import cast, POINTER
from AppOpener import open as open_app_cmd, close as close_app_cmd, give_appnames
import pygetwindow as gw
import pyautogui
import threading
import speech_recognition as sr
import time
from pvrecorder import PvRecorder
import numpy as np
import wave
import struct
import win32com.client 
import re
import difflib
import webbrowser
import random
import urllib.request
import urllib.parse
import pythoncom 
import psutil 
import subprocess
import socket
import json
import cv2 # Sovereign Vision
from elevenlabs.client import ElevenLabs

# Project: M.A.X (Multitasking Assistant Expert) - Sovereign Autonomous Assistant
# User: Boss / Sir (Mario)
# Assistant: Max (JARVIS Mode)

class MAXAssistant:
    def __init__(self):
        self.memory = [] 
        self.active_session = False 
        self.pending_action = None 
        self.eleven_chars_used = 0
        self.eleven_char_limit = 9000
        self.force_sapi = False 
        self.current_speech_process = None 
        self.last_spoken_text = "" 
        self.pref_file = "preferences.json"
        
        # Vision State
        self.vision_active = True
        self.last_seen_time = time.time()
        self.greeting_cooldown = 300 # 5 minutes
        self.face_cascade = cv2.CascadeClassifier(cv2.data.haarcascades + 'haarcascade_frontalface_default.xml')
        
        self.load_preferences() 
        pythoncom.CoInitialize()
        
        # 1. API Key Sanitization
        groq_key = os.environ.get("GROQ_API_KEY", "").strip()
        google_key = os.environ.get("GOOGLE_API_KEY", "").strip()
        self.eleven_key = os.environ.get("ELEVENLABS_API_KEY", "").strip()

        if not groq_key or not google_key:
            print("Sir, API keys are missing. Please set GROQ_API_KEY and GOOGLE_API_KEY.")
            sys.exit(1)

        self.groq_client = groq.Groq(api_key=groq_key)
        self.gemini_client = genai.Client(api_key=google_key)

        # 2. Vocal Synchronization
        self.eleven_client = None
        self.eleven_voice_id = None
        
        if self.eleven_key:
            try:
                self.eleven_client = ElevenLabs(api_key=self.eleven_key)
                print("Synchronizing with ElevenLabs Library... Sir.")
                try:
                    available_voices = self.eleven_client.voices.get_all()
                    if available_voices.voices:
                        sophisticated_voice = next((v for v in available_voices.voices if v.category == "premade" or "Brian" in v.name or "Roger" in v.name), available_voices.voices[0])
                        self.eleven_voice_id = sophisticated_voice.voice_id
                        print(f"ElevenLabs Premium Engine Initialized (Voice: {sophisticated_voice.name}), Sir.")
                    else:
                        self.eleven_client = None
                except Exception:
                    self.eleven_client = None
            except Exception as e:
                print(f"ElevenLabs Initialization Error: {e}, Sir.")

        # 3. Component Initialization
        try:
            self.speaker = win32com.client.Dispatch("SAPI.SpVoice")
            voices = self.speaker.GetVoices()
            for i in range(voices.Count):
                desc = voices.Item(i).GetDescription().lower()
                if "male" in desc or "david" in desc:
                    self.speaker.Voice = voices.Item(i)
                    break
            self.speaker.Rate = 1 
        except Exception: pass
        
        self.recognizer = sr.Recognizer()
        print("Initializing Neural STT Engine (Whisper)... Sir.")
        try:
            self.stt_model = WhisperModel("tiny", device="cpu", compute_type="float32")
            print("Neural STT Engine Online, Boss.")
        except Exception:
            self.stt_model = None 
        
        try:
            self.recorder = PvRecorder(device_index=-1, frame_length=512)
        except Exception as e:
            print(f"Audio Initialization Error: {e}, Sir.")
            sys.exit(1)
        
        # 4. Initiate Sovereign Vision sentry
        threading.Thread(target=self.vision_sentry, daemon=True).start()
        
        atexit.register(self.wipe_memory)
        self.apply_preferences() 
        print("Welcome Boss. System online, Max.")
        self.speak("Systems fully initialized. My name is Max. My eyes are open, and I am at your service, Boss.")

    def vision_sentry(self):
        """Background thread for presence detection"""
        cap = None
        while True:
            if self.vision_active:
                try:
                    if cap is None: cap = cv2.VideoCapture(0)
                    ret, frame = cap.read()
                    if ret:
                        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
                        faces = self.face_cascade.detectMultiScale(gray, 1.3, 5)
                        if len(faces) > 0:
                            current_time = time.time()
                            if not self.active_session and (current_time - self.last_seen_time > self.greeting_cooldown):
                                self.active_session = True
                                self.speak("Welcome back, Boss. I've been monitoring the system while you were away. How can I assist you today?")
                            self.last_seen_time = current_time
                except Exception:
                    if cap: cap.release(); cap = None
                time.sleep(2.0)
            else:
                if cap: cap.release(); cap = None
                time.sleep(5.0)

    def load_preferences(self):
        """Load hardware settings from local vault"""
        defaults = {"volume": 50, "brightness": 50, "force_sapi": False}
        try:
            if os.path.exists(self.pref_file):
                with open(self.pref_file, 'r') as f:
                    self.prefs = {**defaults, **json.load(f)}
            else: self.prefs = defaults
        except: self.prefs = defaults
        self.force_sapi = self.prefs["force_sapi"]

    def save_preferences(self):
        """Save current hardware state to vault"""
        try:
            self.prefs["force_sapi"] = self.force_sapi
            with open(self.pref_file, 'w') as f:
                json.dump(self.prefs, f)
        except: pass

    def apply_preferences(self):
        """Restore last known hardware levels"""
        try:
            self.set_volume(self.prefs["volume"])
            self.set_brightness(self.prefs["brightness"])
        except: pass

    def is_network_stable(self, timeout=0.6):
        """Perform a high-speed pulse check for internet stability"""
        try:
            socket.setdefaulttimeout(timeout)
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.connect(("8.8.8.8", 53))
            return True
        except: return False

    def is_vocalizing(self):
        """Check if any vocal engine is currently active"""
        try:
            if self.current_speech_process and self.current_speech_process.poll() is None:
                return True
            if hasattr(self, 'speaker') and self.speaker.Status.RunningState == 2:
                return True
        except: pass
        return False

    def stop_speech(self):
        """Interrupt current vocal stream immediately and completely"""
        try:
            if self.current_speech_process and self.current_speech_process.poll() is None:
                self.current_speech_process.terminate()
                self.current_speech_process = None
            self.speaker.Speak("", 1 | 2) # Async + Purge
        except: pass

    def speak(self, text):
        if not text: return
        clean_text = text.strip().lower()
        if not any(w in clean_text for w in ["sir", "boss"]):
            text = f"{text}, Sir."
        
        self.last_spoken_text = text 
        print(f"Max: {text}")
        vocal_text = re.sub(r'(\d+)\.(\d+)', r'\1 point \2', text)
        
        self.stop_speech()
        try:
            if hasattr(self, 'recorder') and self.recorder.is_recording:
                self.recorder.stop()
            
            if not self.force_sapi and self.eleven_client and self.eleven_voice_id and self.eleven_chars_used < self.eleven_char_limit:
                if self.is_network_stable():
                    try:
                        audio_generator = self.eleven_client.text_to_speech.convert(
                            voice_id=self.eleven_voice_id, text=vocal_text, model_id="eleven_flash_v2"
                        )
                        audio_bytes = b"".join(list(audio_generator))
                        self.current_speech_process = subprocess.Popen(
                            ["mpv", "--no-video", "--no-terminal", "-"],
                            stdin=subprocess.PIPE, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
                        )
                        def feed_mpv():
                            try:
                                self.current_speech_process.stdin.write(audio_bytes)
                                self.current_speech_process.stdin.close()
                            except: pass
                        threading.Thread(target=feed_mpv, daemon=True).start()
                        self.eleven_chars_used += len(text)
                        return
                    except Exception: pass
            
            self.speaker.Speak(vocal_text, 1) 
        except Exception as e:
            print(f"Vocal Synthesis Error: {e}, Sir.")

    def wipe_memory(self):
        """Cleanup system resources and purge volatile memory on exit"""
        if hasattr(self, 'recorder') and self.recorder.is_recording:
            self.recorder.stop()
        self.memory.clear()
        gc.collect()
        for f in ["wake.wav", "cmd.wav", "pulse.wav", "burst.wav"]:
            if os.path.exists(f):
                try: os.remove(f)
                except: pass
        sys.stderr.write("\nMemory purged, Sir. System offline.\n")

    def think(self, prompt):
        system_prompt = (
            "Role: You are 'MAX,' an ultra-intelligent, dry-witted autonomous system designed by Mario. "
            "Your personality is modeled after J.A.R.V.I.S.—sophisticated, slightly sarcastic, yet impeccably polite and efficient.\n\n"
            "1. Speech Architecture: Use ellipses (...) for mid-sentence pauses. Use natural contractions.\n"
            "2. Dynamics: Address user as 'Sir' or 'Boss.' NEVER use 'Mario' unless identity query.\n"
            "3. Logic: Briefly confirm actions while they happen.\n"
            "4. Constraints: NEVER use 'As an AI...'. NEVER provide lists in bullet points."
        )
        messages = [{"role": "system", "content": system_prompt}]
        for entry in self.memory[-15:]: messages.append(entry)
        messages.append({"role": "user", "content": prompt})
        try:
            completion = self.groq_client.chat.completions.create(
                model="llama-3.3-70b-versatile", messages=messages, extra_body={"store": False}
            )
            return completion.choices[0].message.content
        except Exception:
            try:
                response = self.gemini_client.models.generate_content(model="gemini-1.5-flash", contents=prompt)
                return response.text
            except Exception: return None

    def setup_workspace(self):
        """Deploy hubs in priority sequence"""
        links = [
            "https://open.spotify.com/",
            "https://web.whatsapp.com/",
            "https://hub.sunstone.in/#/dashboard",
            "https://chatgpt.com/",
            "https://gemini.google.com/app",
            "https://www.linkedin.com/feed/",
            "https://www.instagram.com/",
            "https://www.youtube.com/"
        ]
        self.speak("Right away, Boss. Manually initiating Chrome and deploying your hubs.")
        chrome_paths = ["C:/Program Files/Google/Chrome/Application/chrome.exe", "C:/Program Files (x86)/Google/Chrome/Application/chrome.exe"]
        chrome_bin = next((p for p in chrome_paths if os.path.exists(p)), None)
        try:
            for link in links:
                if chrome_bin: subprocess.Popen([chrome_bin, link])
                else: webbrowser.open(link)
                time.sleep(0.4)
        except: pass
        return "Workspace deployment complete, Boss."

    def handle_local_commands(self, text):
        raw_text = text.lower(); text = re.sub(r"\bmax\b|\bm.a.x\b|\bhey\b|\bcan you\b|\bplease\b", "", raw_text).strip()
        if not text: return None

        if "close" in text and ("eyes" in text or "vision" in text):
            self.vision_active = False; return "Vision sentry offline. I have closed my eyes, Boss."
        if ("open" in text or "enable" in text) and ("eyes" in text or "vision" in text):
            self.vision_active = True; return "Vision sentry online. I can see you now, Boss."

        if ("set up" in text or "setup" in text) and "workspace" in text: return self.setup_workspace()
        
        if "switch" in text and ("sapi" in text or "local" in text or "windows" in text):
            self.force_sapi = True; self.save_preferences(); return "Vocal core reconfigured to local protocols, Boss."
        if ("switch" in text or "enable" in text) and ("premium" in text or "elevenlabs" in text):
            self.force_sapi = False; self.save_preferences(); return "Neural uplink restored, Boss."

        if any(w in text for w in ["pause", "resume", "skip", "forward", "backward", "rewind", "fullscreen", "maximize", "mute"]):
            return self.control_browser_media(text)

        if "open" in text or "on chrome" in text:
            is_direct = "on chrome" in text
            raw_target = text.replace("open", "").replace("on chrome", "").strip()
            if raw_target:
                if is_direct:
                    targets = [t.strip() for t in re.split(r",| and ", raw_target) if t.strip()]
                    self.speak(f"Opening {targets[0] if len(targets)==1 else str(len(targets)) + ' hubs'} on Chrome.")
                    for t in targets:
                        url = f"https://www.{t.replace(' ', '')}.com" if "." not in t else t
                        if not url.startswith("http"): url = "https://" + url
                        webbrowser.open(url); time.sleep(0.4)
                    return f"Direct navigation for {raw_target} initiated."
                try:
                    installed = [a.lower() for a in give_appnames()]
                    matches = difflib.get_close_matches(raw_target.lower(), installed, n=1, cutoff=0.6)
                    if matches:
                        self.pending_action = {"type": "app_open", "target": matches[0]}
                        return f"I have located {matches[0]} on your system, Boss. Should I open it?"
                except: pass
                self.pending_action = {"type": "web_open", "query": raw_target}
                return f"I could not find {raw_target} locally, Boss. Shall I search the web?"

        if any(w in text for w in ["system status", "performance", "cpu usage"]):
            cpu = psutil.cpu_percent(); ram = psutil.virtual_memory().percent
            return f"CPU is at {cpu} percent, and RAM usage is at {ram} percent, Boss."
        
        if "close" in text:
            app_name = text.replace("close", "").replace("window", "").strip()
            if not app_name:
                win = gw.getActiveWindow()
                if win: win.close(); return "Active window closed"
            else:
                try:
                    installed = [a.lower() for a in give_appnames()]; matches = difflib.get_close_matches(app_name.lower(), installed, n=1, cutoff=0.6)
                    close_app_cmd(matches[0] if matches else app_name, match_closest=True); return f"Closing {matches[0] if matches else app_name}"
                except: pass

        if "volume" in text:
            nums = [int(n) for n in re.findall(r'\d+', text)]
            if nums:
                try:
                    v = self.get_volume_interface()
                    if v:
                        v.SetMasterVolumeLevelScalar(max(0, min(100, nums[0])) / 100, None)
                        self.prefs["volume"] = nums[0]; self.save_preferences()
                        return f"Volume set to {nums[0]} percent"
                except: pass
        return None

    def get_volume_interface(self):
        try:
            pythoncom.CoInitialize()
            from pycaw.pycaw import IAudioEndpointVolume, IMMDeviceEnumerator
            devices = AudioUtilities.GetSpeakers()
            raw_device = devices if hasattr(devices, 'Activate') else devices.device
            interface = raw_device.Activate(IAudioEndpointVolume._iid_, comtypes.CLSCTX_ALL, None)
            return cast(interface, POINTER(IAudioEndpointVolume))
        except: return None

    def control_browser_media(self, action):
        try:
            windows = [w for w in gw.getAllWindows() if 'YouTube' in w.title or 'Google Chrome' in w.title]
            if windows:
                win = windows[0]
                try:
                    if win.isMinimized: win.restore()
                    win.activate()
                except: pass
                time.sleep(0.4); action = action.lower()
                if "up" in action: pyautogui.press('up', presses=5); return "Increasing volume"
                if "down" in action: pyautogui.press('down', presses=5); return "Decreasing volume"
                if "rewind" in action or "back" in action: pyautogui.press('j'); return "Skipping backward"
                if "forward" in action or "skip" in action: pyautogui.press('l'); return "Skipping forward"
                if "fullscreen" in action: pyautogui.press('f'); return "Full screen toggled"
                pyautogui.press('k'); return "Playback toggled"
            return "No media session found, Boss."
        except: return "Media control error"

    def get_system_audio_peak(self):
        try:
            from pycaw.pycaw import AudioUtilities
            sessions = AudioUtilities.GetAllSessions()
            max_peak = 0
            for session in sessions:
                if hasattr(session, '_ctl'):
                    meter = session._ctl.QueryInterface(comtypes.gen.AudioRouter.IAudioMeterInformation)
                    peak = meter.GetPeakValue()
                    if peak > max_peak: max_peak = peak
            return max_peak
        except: return 0

    def record_audio(self, filename, duration=7, silence_limit=2.0):
        audio_data = []
        last_voice_time = time.time()
        was_vocalizing_during_record = False
        last_pulse_time = time.time()
        stop_triggers = ["stop", "cancel", "shut up", "leave it", "quiet", "abort", "don't do it"]
        system_peak = self.get_system_audio_peak()
        dynamic_threshold = 450 + (system_peak * 1050)
        
        try:
            self.recorder.start(); start_time = time.time()
            while time.time() - start_time < duration:
                frame = self.recorder.read(); audio_data.extend(frame)
                if np.max(np.abs(frame)) > dynamic_threshold: last_voice_time = time.time()
                is_talking = self.is_vocalizing()
                if is_talking: was_vocalizing_during_record = True
                if is_talking and time.time() - last_pulse_time > 1.0:
                    last_pulse_time = time.time()
                    with wave.open("pulse.wav", 'wb') as pf:
                        pf.setnchannels(1); pf.setsampwidth(2); pf.setframerate(self.recorder.sample_rate); pf.writeframes(struct.pack('h' * len(audio_data), *audio_data))
                    try:
                        with sr.AudioFile("pulse.wav") as source:
                            pulse_audio = self.recognizer.record(source); pulse_text = self.recognizer.recognize_google(pulse_audio).lower()
                            if any(t in pulse_text for t in stop_triggers):
                                self.stop_speech(); return "INTERRUPT", True
                    except: pass
                if time.time() - last_voice_time > silence_limit and len(audio_data) > 16000: break
        finally: self.recorder.stop()
        with wave.open(filename, 'wb') as f:
            f.setnchannels(1); f.setsampwidth(2); f.setframerate(self.recorder.sample_rate); f.writeframes(struct.pack('h' * len(audio_data), *audio_data))
        return filename, was_vocalizing_during_record

    def wait_for_wake_word(self):
        print("Listening for 'Max' (Sir)...")
        wake_wav, _ = self.record_audio("wake.wav", duration=4, silence_limit=1.5)
        if wake_wav == "INTERRUPT": return None
        try:
            with sr.AudioFile(wake_wav) as source:
                audio = self.recognizer.record(source); text = self.recognizer.recognize_google(audio).lower()
                if any(word in text.split() for word in ["max", "macs", "maxwell"]):
                    return "emotive" if "how are you" in text else "standard"
        except: pass
        finally:
            if os.path.exists(wake_wav): os.remove(wake_wav)
        return None

    def capture_command(self):
        if not self.is_vocalizing(): print("Listening for command (Sir)...")
        cmd_wav, was_vocalizing = self.record_audio("cmd.wav", duration=8, silence_limit=2.0)
        if cmd_wav == "INTERRUPT": return "stop", True
        try:
            segments, info = self.stt_model.transcribe(cmd_wav, beam_size=5)
            return "".join([segment.text for segment in segments]).strip(), was_vocalizing
        except: return "", False
        finally:
            if os.path.exists(cmd_wav): os.remove(cmd_wav)

    def is_echo(self, command):
        if not self.last_spoken_text: return False
        cmd_words = set(re.findall(r'\w+', command.lower()))
        if not cmd_words: return False
        ref_words = set(re.findall(r'\w+', self.last_spoken_text.lower()))
        overlap = cmd_words.intersection(ref_words)
        return (len(overlap) / (len(cmd_words) + 1e-6)) > 0.6

    def run(self):
        try:
            while True:
                if not self.active_session:
                    wake_type = self.wait_for_wake_word()
                    if wake_type:
                        self.active_session = True
                        if wake_type == "emotive": self.speak("I am functioning at peak efficiency, Boss. How can I assist?")
                        else: self.speak("Yes, Boss? I am listening")
                
                if self.active_session:
                    command, overlapped_with_speech = self.capture_command()
                    if command and len(command) > 1:
                        cmd_lower = command.lower().replace(",", "").replace(".", "").replace("?", "").strip()
                        
                        # 1. INTERRUPT CHECK
                        stop_triggers = ["stop", "cancel", "shut up", "leave it", "quiet", "abort", "don't do it"]
                        is_stop_cmd = any(trigger in cmd_lower for trigger in stop_triggers) or command == "stop"
                        if is_stop_cmd and (self.is_vocalizing() or overlapped_with_speech):
                            print(f"Interrupt Detected: '{command}'")
                            self.stop_speech(); self.pending_action = None; time.sleep(0.5); continue

                        # 2. FUZZY ECHO PURGE
                        if self.is_echo(cmd_lower):
                            print(f"Fuzzy Echo Suppressed: '{command}'")
                            continue

                        print(f"Mario: '{command}'")

                        if self.pending_action:
                            if any(word in cmd_lower for word in ["yes", "yeah", "confirm", "go ahead", "proceed", "okay"]):
                                if self.pending_action["type"] == "web_open":
                                    query = self.pending_action["query"].lower().strip()
                                    known = ["youtube", "facebook", "instagram", "google", "github", "whatsapp", "twitter", "linkedin"]
                                    if any(s in query for s in known) or "." in query:
                                        url = f"https://www.{query.replace(' ', '')}.com" if "." not in query else query
                                        if not url.startswith("http"): url = "https://" + url
                                        webbrowser.open(url); self.speak(f"Navigating directly to {query}, Boss.")
                                    else:
                                        webbrowser.open(f"https://www.google.com/search?q={query}"); self.speak(f"Searching web for {query}, Boss.")
                                elif self.pending_action["type"] == "app_open":
                                    open_app_cmd(self.pending_action["target"], match_closest=True); self.speak(f"Launching {self.pending_action['target']}, Boss.")
                                self.pending_action = None; time.sleep(1.0); continue
                            else:
                                self.speak("Action cancelled"); self.pending_action = None; time.sleep(1.0); continue

                        action_result = self.handle_local_commands(command)
                        if self.pending_action: self.speak(action_result); time.sleep(1.0); continue

                        thought_prompt = f"System Update: {action_result}. {command}" if action_result else command
                        emotive_response = self.think(thought_prompt)
                        
                        if any(word in cmd_lower for word in ["sleep", "stop listening", "thank you", "thanks"]):
                            self.active_session = False; self.speak(emotive_response if emotive_response else "Standby, Boss."); time.sleep(1.0); continue
                            
                        if emotive_response: self.speak(emotive_response)
                        elif action_result: self.speak(action_result)
                        else: self.speak("I'm afraid I didn't quite catch that, Sir.")
                        
                        if emotive_response:
                            self.memory.append({"role": "user", "content": command})
                            self.memory.append({"role": "assistant", "content": emotive_response})
                        time.sleep(1.0)
        except KeyboardInterrupt: sys.exit(0)

if __name__ == "__main__":
    assistant = MAXAssistant()
    assistant.run()
