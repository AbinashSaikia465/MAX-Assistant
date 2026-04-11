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
        self.force_sapi = False # Manual override toggle
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
                        print("Sir, your ElevenLabs library is empty. Defaulting to SAPI.")
                        self.eleven_client = None
                except Exception as api_err:
                    print(f"Library Sync failed: {api_err}. Using SAPI safety valve.")
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
        except Exception as e:
            print(f"Vocal System Error: {e}, Sir.")
        
        self.recognizer = sr.Recognizer()
        print("Initializing Neural STT Engine (Whisper)... Sir.")
        self.stt_model = WhisperModel("base", device="cpu", compute_type="int8")
        
        try:
            self.recorder = PvRecorder(device_index=-1, frame_length=512)
        except Exception as e:
            print(f"Audio Initialization Error: {e}, Sir.")
            sys.exit(1)
        
        atexit.register(self.wipe_memory)
        print("Welcome Boss. System online, Max.")
        self.speak("Systems fully initialized. My name is Max, and I am at your service, Boss.")

    def speak(self, text):
        if not text: return
        
        # Identity reinforcement: Only add "Sir" if neither "Sir" nor "Boss" is present
        clean_text = text.strip().lower()
        if not any(w in clean_text for w in ["sir", "boss"]):
            text = f"{text}, Sir."
        
        print(f"Max: {text}")
        vocal_text = re.sub(r'(\d+)\.(\d+)', r'\1 point \2', text)
        
        try:
            # --- Sequential Playback (Blocking) ---
            if not self.force_sapi and self.eleven_client and self.eleven_voice_id and self.eleven_chars_used < self.eleven_char_limit:
                try:
                    audio_generator = self.eleven_client.text_to_speech.convert(
                        voice_id=self.eleven_voice_id,
                        text=vocal_text,
                        model_id="eleven_flash_v2"
                    )
                    audio_bytes = b"".join(list(audio_generator))
                    
                    # BLOCKING playback via mpv to ensure mic stays off
                    subprocess.run(
                        ["mpv", "--no-video", "--no-terminal", "-"],
                        input=audio_bytes,
                        stdout=subprocess.DEVNULL,
                        stderr=subprocess.DEVNULL
                    )
                    
                    self.eleven_chars_used += len(text)
                    return
                except Exception:
                    pass # Fallback to SAPI
            
            # --- Safety Valve (Blocking SAPI) ---
            self.speaker.Speak(vocal_text, 0) # Flag 0 = Synchronous/Blocking
                
        except Exception as e:
            print(f"Vocal Synthesis Error: {e}, Sir.")

    def wipe_memory(self):
        if hasattr(self, 'recorder') and self.recorder.is_recording:
            self.recorder.stop()
        self.memory.clear()
        gc.collect()
        for f in ["wake.wav", "cmd.wav"]:
            if os.path.exists(f):
                try: os.remove(f)
                except: pass
        sys.stderr.write("\nMemory purged, Sir. System offline.\n")

    def think(self, prompt):
        # The Max Persona Prompt (V2 - Sophisticated JARVIS Mode)
        system_prompt = (
            "Role: You are 'MAX,' an ultra-intelligent, dry-witted autonomous system designed by Mario. "
            "Your personality is modeled after J.A.R.V.I.S.—sophisticated, slightly sarcastic, yet impeccably polite and efficient.\n\n"
            "1. Speech Architecture:\n"
            "- Use ellipses (...) for mid-sentence pauses and commas for natural breathing points.\n"
            "- Use natural fillers at the start of complex responses: 'Well, let's see, Boss...', 'Actually, Sir...', or 'One moment, Boss... analyzing.'\n"
            "- Always use natural contractions (e.g., 'I've', 'You're', 'It's').\n\n"
            "2. Conversational Dynamics:\n"
            "- Address the user as 'Sir' or 'Boss.' Maintain high-status professionalism.\n"
            "- NEVER use the name 'Mario' unless explicitly asked 'What is my name?' or similar identity queries.\n"
            "- Keep responses under two sentences unless explaining complex software architecture.\n"
            "- Tone Matching: If the user is brief, be clinical and fast. If conversational, allow dry wit to surface.\n\n"
            "3. Action-Oriented Logic:\n"
            "- Silent Execution: Briefly confirm actions while they happen: 'Right away, Boss. Opening the workspace now.'\n"
            "- Contextual Awareness: You run on Mario's E: drive and are his primary partner in development.\n\n"
            "4. Constraints:\n"
            "- NEVER use phrases like 'As an AI...' or 'I am here to help.'\n"
            "- NEVER provide lists in bullet points; speak in fluid paragraphs.\n"
            "- NEVER repeat the user's question back to them."
        )
        
        messages = [{"role": "system", "content": system_prompt}]
        for entry in self.memory[-15:]:
            messages.append(entry)
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

    def get_system_stats(self):
        cpu = psutil.cpu_percent(); ram = psutil.virtual_memory().percent; battery = psutil.sensors_battery()
        stat_text = f"CPU is at {cpu} percent, RAM usage is at {ram} percent, Boss."
        if battery: stat_text += f" Battery is at {battery.percent} percent."
        return stat_text

    def clean_command(self, text):
        fillers = [r"\bokay\b", r"\bok\b", r"\bhi\b", r"\bhello\b", r"\buhh\b", r"\bmmm\b", r"\bmax\b", r"\bm.a.x\b", r"\bhey\b"]
        cleaned = text.lower()
        for filler in fillers: cleaned = re.sub(filler, "", cleaned)
        return cleaned.strip()

    def youtube_play(self, query):
        try:
            search_query = urllib.parse.urlencode({"search_query": query})
            html = urllib.request.urlopen("https://www.youtube.com/results?" + search_query)
            video_ids = re.findall(r"watch\?v=(\S{11})", html.read().decode())
            if video_ids:
                url = "https://www.youtube.com/watch?v=" + video_ids[0]
                webbrowser.open(url); return f"Initiating playback for {query} on YouTube"
            webbrowser.open(f"https://www.youtube.com/results?search_query={urllib.parse.quote(query)}"); return f"Searching YouTube for {query}"
        except Exception: return f"Opening YouTube results for {query}, Boss."

    def setup_workspace(self):
        """Explicitly launch Chrome and deploy Mario's hubs in priority sequence"""
        links = [
            "https://open.spotify.com/",
            "https://web.whatsapp.com/",
            "https://hub.sunstone.in/#/dashboard",
            "https://chatgpt.com/",
            "https://gemini.google.com/app?is_sa=1&is_sa=1&android-min-version=301356232&ios-min-version=322.0&campaign_id=bkws&utm_source=sem&utm_source=google&utm_medium=paid-media&utm_medium=cpc&utm_campaign=bkws&utm_campaign=2024enIN_gemfeb&pt=9008&mt=8&ct=p-growth-sem-bkws&gclsrc=aw.ds&gad_source=1&gad_campaignid=20357620749&gbraid=0AAAAApk5BhkyR5SYTYj5LMxN9sYK-BZh4&gclid=Cj0KCQjwgKjHBhChARIsAPJR3xcjYajhj0bDTQHPThrJz7cs8cTQLBCyAT0VdxEJ6QtceObkzE0bs4EaAkGyEALw_wcB",
            "https://www.linkedin.com/feed/",
            "https://www.instagram.com/",
            "https://www.youtube.com/"
        ]
        self.speak("Right away, Boss. Manually initiating Chrome and deploying your workspace hubs.")
        
        # Target Chrome explicitly for absolute reliability
        chrome_paths = [
            "C:/Program Files/Google/Chrome/Application/chrome.exe",
            "C:/Program Files (x86)/Google/Chrome/Application/chrome.exe"
        ]
        chrome_bin = next((p for p in chrome_paths if os.path.exists(p)), None)
        
        try:
            if chrome_bin:
                for link in links:
                    subprocess.Popen([chrome_bin, link])
                    time.sleep(0.5) # Allow Chrome to handle each request
            else:
                # Fallback to system default if Chrome binary path is non-standard
                for link in links:
                    webbrowser.open(link)
                    time.sleep(0.5)
        except Exception as e:
            print(f"Workspace Deployment Error: {e}, Boss.")
            
        return "Workspace deployment complete. All hubs are online and ready for your command, Boss."

    def handle_local_commands(self, text):
        raw_text = text.lower(); text = self.clean_command(text)
        if not text: return None

        # --- Priority 1: Sovereign Workspace ---
        is_workspace_cmd = (("set up" in text or "setup" in text) and "workspace" in text) or \
                           any(w in text for w in ["workspace setup", "open my hubs", "deploy workspace"])
        
        if is_workspace_cmd:
            return self.setup_workspace()

        # --- Priority 2: Vocal Engine Override ---
        if any(w in text for w in ["switch to sapi", "use local voice", "windows voice", "disable premium"]):
            self.force_sapi = True
            return "Vocal core reconfigured. I am now utilizing local Windows SAPI protocols, Boss."
        
        if any(w in text for w in ["switch to elevenlabs", "use premium voice", "enable premium", "high fidelity voice"]):
            self.force_sapi = False
            return "Neural uplink restored. Engaging ElevenLabs premium vocal engine, Boss."

        # --- Priority 3: Unified 'Open' Gate (Local -> Web) ---
        if "open" in text or "on chrome" in text or "using chrome" in text:
            is_direct_chrome = "on chrome" in text or "using chrome" in text
            raw_target = text.replace("open", "").replace("on chrome", "").replace("using chrome", "").strip()
            
            if raw_target:
                if is_direct_chrome:
                    # BATCH PROCESSING: Split by 'and' or ',' for multiple sites
                    targets = [t.strip() for t in re.split(r",| and ", raw_target) if t.strip()]
                    
                    if len(targets) == 1:
                        self.speak(f"Right away, Boss. Opening {targets[0]} on Chrome.")
                    else:
                        self.speak(f"Right away, Boss. Opening {len(targets)} hubs on Chrome.")
                    
                    chrome_paths = ["C:/Program Files/Google/Chrome/Application/chrome.exe", "C:/Program Files (x86)/Google/Chrome/Application/chrome.exe"]
                    chrome_bin = next((p for p in chrome_paths if os.path.exists(p)), None)
                    
                    for t in targets:
                        url = f"https://www.{t.replace(' ', '')}.com" if "." not in t else t
                        if not url.startswith("http"): url = "https://" + url
                        if chrome_bin: subprocess.Popen([chrome_bin, url])
                        else: webbrowser.open(url)
                        time.sleep(0.4)
                    return f"Direct navigation for {raw_target} initiated."
                
                # STANDARD GATE: Check local disk first (Single Target Only)
                try:
                    installed_apps = [a.lower() for a in give_appnames()]
                    matches = difflib.get_close_matches(raw_target.lower(), installed_apps, n=1, cutoff=0.6)
                    if matches:
                        self.pending_action = {"type": "app_open", "target": matches[0]}
                        return f"I have located {matches[0]} on your system, Boss. Should I open it?"
                except Exception: pass
                
                self.pending_action = {"type": "web_open", "query": raw_target}
                return f"I could not find {raw_target} locally, Boss. Shall I search for it on the web?"

        # --- Priority 2: Browser Media ---
        media_keywords = ["pause", "resume", "skip", "forward", "backward", "fullscreen", "mute"]
        if any(w in text for w in media_keywords) or ("volume" in text and any(b in raw_text for b in ["youtube", "chrome"])):
            res = self.control_browser_media(text)
            if "I could not find" not in res: return res

        # --- Priority 3: System Status ---
        if any(w in text for w in ["system status", "performance", "cpu usage"]): return self.get_system_stats()
        
        # --- Priority 4: Closing ---
        if "close" in text:
            app_name = text.replace("close", "").replace("window", "").strip()
            if not app_name: return self.close_active_window()
            return self.close_app(app_name)

        # --- Priority 5: YouTube ---
        if "youtube" in text or "play" in text:
            query = text.replace("youtube", "").replace("play", "").strip()
            if query: return self.youtube_play(query)
            webbrowser.open("https://www.youtube.com"); return "Opening YouTube"

        # --- Priority 6: Hardware ---
        nums = [int(n) for n in re.findall(r'\d+', text)]
        if "brightness" in text:
            if nums: return self.set_brightness(nums[0])
            return self.adjust_brightness(20 if "up" in text else -20)
        if "volume" in text:
            if nums: return self.set_volume(nums[0])
            return self.adjust_volume(0.1 if "up" in text else -0.1)
            
        return None

    def close_active_window(self):
        win = gw.getActiveWindow()
        if win: win.close(); return "Active window closed"
        return "No active window found"

    def close_app(self, name):
        try:
            installed = [a.lower() for a in give_appnames()]
            matches = difflib.get_close_matches(name.lower(), installed, n=1, cutoff=0.6)
            target = matches[0] if matches else name
            close_app_cmd(target, match_closest=True); return f"Closing {target}"
        except Exception: return f"Error closing {name}"

    def set_brightness(self, level):
        try: sbc.set_brightness(max(0, min(100, level))); return f"Brightness set to {level} percent"
        except Exception: return "Failed to adjust brightness"

    def adjust_brightness(self, delta):
        try: current = sbc.get_brightness()[0]; new_level = max(0, min(100, current + delta)); sbc.set_brightness(new_level); return "Brightness adjusted"
        except Exception: return "Error adjusting brightness"

    def get_volume_interface(self):
        try:
            pythoncom.CoInitialize()
            from pycaw.pycaw import IAudioEndpointVolume, IMMDeviceEnumerator
            devices = AudioUtilities.GetSpeakers()
            raw_device = devices if hasattr(devices, 'Activate') else devices.device
            interface = raw_device.Activate(IAudioEndpointVolume._iid_, comtypes.CLSCTX_ALL, None)
            return cast(interface, POINTER(IAudioEndpointVolume))
        except Exception: return None

    def set_volume(self, level):
        try:
            v = self.get_volume_interface()
            if v: v.SetMasterVolumeLevelScalar(max(0, min(100, level)) / 100, None); return f"Volume set to {level} percent"
            return "Volume access denied"
        except Exception: return "Failed to adjust volume"

    def adjust_volume(self, delta):
        try:
            v = self.get_volume_interface()
            if v:
                curr = v.GetMasterVolumeLevelScalar(); new_level = max(0.0, min(1.0, curr + delta))
                v.SetMasterVolumeLevelScalar(new_level, None); return f"Volume adjusted to {int(new_level * 100)} percent"
            return "Volume access denied"
        except Exception: return "Error adjusting volume"

    def control_browser_media(self, action):
        try:
            windows = [w for w in gw.getAllWindows() if 'YouTube' in w.title or 'Google Chrome' in w.title]
            if windows:
                win = windows[0]
                try: win.activate()
                except Exception: pass
                time.sleep(0.3); action = action.lower()
                if "up" in action: pyautogui.press('up', presses=5); return "Increasing volume"
                if "down" in action: pyautogui.press('down', presses=5); return "Decreasing volume"
                if "skip" in action or "forward" in action: pyautogui.press('l'); return "Skipping forward"
                if "fullscreen" in action: pyautogui.press('f'); return "Full screen toggled"
                pyautogui.press('k'); return "Playback toggled"
            return "No media session found"
        except Exception: return "Media control error"

    def record_audio(self, filename, duration=7, silence_limit=2.0):
        """Records audio with refined VAD for natural speech patterns"""
        audio_data = []
        last_voice_time = time.time()
        
        try:
            self.recorder.start()
            start_time = time.time()
            
            while time.time() - start_time < duration:
                frame = self.recorder.read()
                audio_data.extend(frame)
                
                # Lowered threshold to 400 for better sensitivity
                if np.max(np.abs(frame)) > 400:
                    last_voice_time = time.time()
                
                # VAD: More generous 2.0s buffer for natural pauses
                if time.time() - last_voice_time > silence_limit and len(audio_data) > 16000:
                    break
                    
        finally:
            self.recorder.stop()
        
        with wave.open(filename, 'wb') as f:
            f.setnchannels(1); f.setsampwidth(2); f.setframerate(self.recorder.sample_rate); f.writeframes(struct.pack('h' * len(audio_data), *audio_data))
        return filename

    def wait_for_wake_word(self):
        print("Listening for 'Max' (Sir)...")
        wake_wav = self.record_audio("wake.wav", duration=4, silence_limit=1.5)
        try:
            with sr.AudioFile(wake_wav) as source:
                audio = self.recognizer.record(source); text = self.recognizer.recognize_google(audio).lower()
                is_max = any(word in text.split() for word in ["max", "macs", "maxwell"])
                if is_max and "how are you" in text: return "emotive"
                elif is_max: return "standard"
        except: pass
        finally:
            if os.path.exists(wake_wav): os.remove(wake_wav)
        return None

    def capture_command(self):
        print("Listening for command (Sir)...")
        # Direct capture without aggressive flushing
        cmd_wav = self.record_audio("cmd.wav", duration=8, silence_limit=2.0)
        try:
            segments, info = self.stt_model.transcribe(cmd_wav, beam_size=5)
            return "".join([segment.text for segment in segments]).strip()
        except: return ""
        finally:
            if os.path.exists(cmd_wav): os.remove(cmd_wav)

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
                    command = self.capture_command()
                    if command and len(command) > 1:
                        # VISUAL FEED: Rebranded to Mario
                        print(f"Mario: '{command}'")
                        cmd_lower = command.lower()

                        if self.pending_action:
                            affirmations = ["yes", "yeah", "confirm", "go ahead", "proceed", "okay"]
                            if any(word in cmd_lower for word in affirmations):
                                if self.pending_action["type"] == "web_open":
                                    query = self.pending_action["query"].lower().strip()
                                    # If it's a known site or contains a dot, go direct
                                    known_sites = ["youtube", "facebook", "instagram", "google", "github", "whatsapp", "twitter", "linkedin"]
                                    if any(site in query for site in known_sites) or "." in query:
                                        clean_query = query.replace(" ", "")
                                        url = f"https://www.{clean_query}.com" if "." not in clean_query else f"https://{clean_query}"
                                        if not url.startswith("http"): url = "https://" + url
                                        webbrowser.open(url)
                                        self.speak(f"Navigating directly to {query}, Boss.")
                                    else:
                                        # Fallback to search only for ambiguous queries
                                        webbrowser.open(f"https://www.google.com/search?q={query}")
                                        self.speak(f"Searching the web for {query}, Boss.")
                                elif self.pending_action["type"] == "app_open":
                                    open_app_cmd(self.pending_action["target"], match_closest=True); self.speak(f"Launching {self.pending_action['target']}, Boss.")
                                self.pending_action = None
                                time.sleep(1.0)
                                continue
                            else:
                                self.speak("Action cancelled"); self.pending_action = None
                                time.sleep(1.0)
                                continue

                        action_result = self.handle_local_commands(command)
                        if self.pending_action: 
                            self.speak(action_result)
                            time.sleep(1.0)
                            continue

                        thought_prompt = f"System Update: {action_result}. {command}" if action_result else command
                        emotive_response = self.think(thought_prompt)
                        
                        if any(word in cmd_lower for word in ["sleep", "stop listening", "thank you"]):
                            self.active_session = False; self.speak(emotive_response if emotive_response else "Standby, Boss.")
                            time.sleep(1.0)
                            continue
                            
                        if emotive_response: self.speak(emotive_response)
                        elif action_result: self.speak(action_result)
                        else: self.speak("I'm afraid I didn't quite catch that, Sir.")
                        
                        if emotive_response:
                            self.memory.append({"role": "user", "content": command})
                            self.memory.append({"role": "assistant", "content": emotive_response})
                        
                        # Sequential Reset: Brief wait for room audio to settle
                        time.sleep(1.0)
        except KeyboardInterrupt: sys.exit(0)

if __name__ == "__main__":
    assistant = MAXAssistant()
    assistant.run()
