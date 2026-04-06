import os
import gc
import sys
import atexit
import groq
from google import genai
from faster_whisper import WhisperModel
import screen_brightness_control as sbc
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
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

# Define required COM CLSID if not available in pycaw
CLSID_MMDeviceEnumerator = comtypes.GUID('{BCDE0395-E52F-467C-8E3D-C4579291692E}')

# Project: M.A.X (Multitasking Assistant Expert) - Sovereign Autonomous Assistant
# User: Mario (Sir)
# Assistant: Max (JARVIS Mode)

class MAXAssistant:
    def __init__(self):
        self.memory = [] 
        self.active_session = False 
        self.pending_action = None 
        pythoncom.CoInitialize()
        
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
            print(f"Vocal System Initialization Error: {e}, Sir.")
            sys.exit(1)
        
        groq_key = os.environ.get("GROQ_API_KEY")
        google_key = os.environ.get("GOOGLE_API_KEY")
        if not groq_key or not google_key:
            print("Sir, API keys are missing. Please set GROQ_API_KEY and GOOGLE_API_KEY.")
            sys.exit(1)

        self.groq_client = groq.Groq(api_key=groq_key)
        self.gemini_client = genai.Client(api_key=google_key)
        
        self.recognizer = sr.Recognizer()
        print("Initializing Neural STT Engine (Whisper)... Sir.")
        self.stt_model = WhisperModel("base", device="cpu", compute_type="int8")
        
        try:
            self.recorder = PvRecorder(device_index=-1, frame_length=512)
        except Exception as e:
            print(f"Audio Initialization Error: {e}, Sir.")
            sys.exit(1)
        
        atexit.register(self.wipe_memory)
        print("Welcome Mario. System online, Max.")
        self.speak("Systems fully initialized. My name is Max, and I am at your service, Sir.")

    def speak(self, text):
        if not text: return
        
        # Identity reinforcement for Max (avoiding duplicates)
        clean_text = text.strip().lower().rstrip(".,!?;")
        if not clean_text.startswith("sir") and not clean_text.endswith("sir"):
            text = f"{text}, Sir."
        
        # Display the original text with decimals
        print(f"Max: {text}")
        
        # Phonetic Decimal Processor for VOCAL stream only
        vocal_text = re.sub(r'(\d+)\.(\d+)', r'\1 point \2', text)
        
        try:
            if hasattr(self, 'recorder') and self.recorder.is_recording:
                self.recorder.stop()
            
            sentences = [s.strip() for s in vocal_text.split('.') if s.strip()]
            for sentence in sentences:
                self.speaker.Speak(sentence)
                time.sleep(0.4)
                
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
        # The Max Persona Prompt
        system_prompt = (
            "You are Max (Multitasking Assistant Expert), a highly advanced AI assistant like JARVIS. "
            "Your user is Mario, whom you always address as Sir. "
            "You are sophisticated, loyal, and possess a dry, British wit. "
            "Acknowledge the user's emotions or situation. "
            "Keep your responses concise but natural. Every response must begin or end with 'Sir.' "
            "Always refer to yourself as Max. Never use the acronym M.A.X in your responses unless specifically explaining its meaning."
        )
        
        messages = [{"role": "system", "content": system_prompt}]
        for entry in self.memory[-15:]:
            messages.append(entry)
        messages.append({"role": "user", "content": prompt})
        
        try:
            completion = self.groq_client.chat.completions.create(
                model="llama-3.3-70b-versatile", 
                messages=messages, 
                extra_body={"store": False}
            )
            return completion.choices[0].message.content
        except Exception:
            try:
                response = self.gemini_client.models.generate_content(model="gemini-1.5-flash", contents=prompt)
                return response.text
            except Exception: return None

    def get_system_stats(self):
        cpu = psutil.cpu_percent(); ram = psutil.virtual_memory().percent; battery = psutil.sensors_battery()
        stat_text = f"CPU is at {cpu} percent, and RAM usage is at {ram} percent, Sir."
        if battery: stat_text += f" Battery is at {battery.percent} percent."
        return stat_text

    def clean_command(self, text):
        # Identity focused clean
        fillers = [r"\bokay\b", r"\bok\b", r"\bhi\b", r"\bhello\b", r"\buhh\b", r"\bmmm\b", r"\bmax\b", r"\bm.a.x\b", r"\bhey\b", r"\bcan you\b", r"\bplease\b", r"\bjust\b"]
        cleaned = text.lower()
        for filler in fillers: cleaned = re.sub(filler, "", cleaned)
        return cleaned.strip()

    def youtube_play(self, query):
        try:
            print(f"[YouTube] Direct stream search: {query}")
            search_query = urllib.parse.urlencode({"search_query": query})
            html = urllib.request.urlopen("https://www.youtube.com/results?" + search_query)
            video_ids = re.findall(r"watch\?v=(\S{11})", html.read().decode())
            if video_ids:
                url = "https://www.youtube.com/watch?v=" + video_ids[0]
                webbrowser.open(url); return f"Initiating playback for {query} on YouTube"
            else:
                webbrowser.open(f"https://www.youtube.com/results?search_query={urllib.parse.quote(query)}"); return f"Searching YouTube for {query}"
        except Exception: return f"Opening YouTube results for {query}, Sir."

    def handle_local_commands(self, text):
        raw_text = text.lower()
        text = self.clean_command(text)
        if not text: return None

        # --- Priority 1: Browser Media Control (Intercepts YouTube/Chrome commands) ---
        media_keywords = ["pause", "resume", "skip", "forward", "backward", "rewind", "fullscreen", "full screen", "maximize", "mute", "unmute"]
        is_media_control = any(w in text for w in media_keywords)
        is_browser_volume = "volume" in text and any(b in raw_text for b in ["youtube", "chrome", "browser"])
        is_browser_play = text.strip() == "play" or (text.startswith("play") and any(b in raw_text for b in ["youtube", "chrome"]))

        if is_media_control or is_browser_volume or is_browser_play:
            # Avoid intercepting "play [song]" unless youtube/chrome is specifically mentioned
            if not (text.startswith("play ") and len(text.split()) > 2 and "youtube" not in raw_text):
                res = self.control_browser_media(text)
                if "I could not find" not in res: return res

        # --- Priority 2: System Status ---
        if any(w in text for w in ["system status", "cpu", "ram", "battery", "performance"]): return self.get_system_stats()
        
        # --- Priority 3: Web navigation on Chrome ---
        if "on chrome" in text or "using chrome" in text:
            target = text.replace("open", "").replace("on chrome", "").replace("using chrome", "").strip()
            if target:
                self.pending_action = {"type": "web_open", "query": target}
                return f"I have located {target} on the web. Should I open it?"

        # --- Priority 4: Closing actions ---
        if any(w in text for w in ["close", "exit", "terminate", "kill"]):
            if "window" in text and "tab" not in text and not any(a in text for a in ["chrome", "browser", "youtube"]): return self.close_active_window()
            elif "tab" in text: return self.close_active_tab()
            else:
                app_name = text.replace("close", "").replace("exit", "").replace("terminate", "").replace("kill", "").replace("window", "").strip()
                if app_name: return self.close_app(app_name)

        # --- Priority 5: YouTube / Music Initiation ---
        if "youtube" in text or "play" in text:
            if text == "open youtube" or text == "youtube":
                webbrowser.open("https://www.youtube.com"); return "Opening YouTube"
            
            live_streams = {
                "lofi": "https://www.youtube.com/watch?v=jfKfPfyJRdk", 
                "jazz": "https://www.youtube.com/watch?v=5yx6BWbL1E4", 
                "synthwave": "https://www.youtube.com/watch?v=4xDzrJKXOOY", 
                "classical": "https://www.youtube.com/watch?v=mIYzp5rcTvU"
            }
            if "music" in text and not any(word in text for word in ["search", "find", "google"]):
                selection = "lofi"
                for genre in live_streams:
                    if genre in text: selection = genre; break
                if "random" in text or "some" in text: selection = random.choice(list(live_streams.keys()))
                webbrowser.open(live_streams[selection]); return f"Initiating {selection} stream for you, Sir."
            
            query = text.replace("youtube", "").replace("play", "").replace("on", "").replace("search", "").replace("open", "").strip()
            if query: return self.youtube_play(query)
            else: webbrowser.open("https://www.youtube.com"); return "Opening YouTube"

        # --- Priority 6: Web Search ---
        elif "google" in text or "search" in text:
            query = text.replace("google", "").replace("search", "").replace("for", "").strip()
            if query: webbrowser.open(f"https://www.google.com/search?q={query}"); return f"Searching the web for {query}"

        # --- Priority 7: Brightness and Volume (System) ---
        nums = [int(n) for n in re.findall(r'\d+', text)]
        if "brightness" in text:
            if nums: return self.set_brightness(nums[0])
            elif any(w in text for w in ["increase", "up", "more", "higher"]): return self.adjust_brightness(delta=20)
            elif any(w in text for w in ["decrease", "down", "less", "lower"]): return self.adjust_brightness(delta=-20)
        
        elif any(w in text for w in ["volume", "sound", "mute", "unmute"]):
            if "mute" in text and "unmute" not in text: return self.set_volume(0)
            elif "unmute" in text: return self.set_volume(20)
            elif nums: return self.set_volume(nums[0])
            elif any(w in text for w in ["increase", "up", "more", "higher"]): return self.adjust_volume(delta=0.1)
            elif any(w in text for w in ["decrease", "down", "less", "lower"]): return self.adjust_volume(delta=-0.1)
            
        # --- Priority 8: App opening ---
        elif "open" in text: return self.open_app(text.replace("open", "").strip())
        
        # --- Priority 9: Time ---
        elif "time" in text: return f"The current time is {time.strftime('%I:%M %p')}"
        
        return None

    def close_active_window(self):
        active_window = gw.getActiveWindow()
        if active_window: active_window.close(); return "The active window has been closed"
        return "I could not find an active window to close"

    def close_active_tab(self):
        pyautogui.hotkey('ctrl', 'w'); return "Tab closed"

    def close_app(self, name):
        try:
            installed_apps = [a.lower() for a in give_appnames()]; matches = difflib.get_close_matches(name.lower(), installed_apps, n=1, cutoff=0.6)
            target = matches[0] if matches else name
            close_app_cmd(target, match_closest=True); return f"Closing {target}"
        except Exception: return f"Error closing {name}"

    def set_brightness(self, level):
        try: level = max(0, min(100, level)); sbc.set_brightness(level); return f"Brightness set to {level} percent"
        except Exception: return "Failed to adjust brightness"

    def adjust_brightness(self, delta):
        try: current = sbc.get_brightness()[0]; new_level = max(0, min(100, current + delta)); sbc.set_brightness(new_level); return f"Brightness adjusted to {new_level} percent"
        except Exception: return "Error adjusting brightness"

    def get_volume_interface(self):
        try:
            pythoncom.CoInitialize()
            from pycaw.pycaw import IAudioEndpointVolume, IMMDeviceEnumerator
            
            # Use AudioUtilities to get speakers endpoint
            devices = AudioUtilities.GetSpeakers()
            
            # Handle potential wrapper class (AudioDevice) in some pycaw versions
            if hasattr(devices, 'Activate'):
                raw_device = devices
            elif hasattr(devices, 'device'):
                raw_device = devices.device
            else:
                # Direct creation if utility fails to provide IMMDevice
                enumerator = comtypes.CoCreateInstance(
                    CLSID_MMDeviceEnumerator,
                    IMMDeviceEnumerator,
                    comtypes.CLSCTX_ALL)
                raw_device = enumerator.GetDefaultAudioEndpoint(0, 1)
            
            interface = raw_device.Activate(
                IAudioEndpointVolume._iid_, comtypes.CLSCTX_ALL, None)
            return cast(interface, POINTER(IAudioEndpointVolume))
        except Exception as e:
            print(f"System Volume Access Error: {e}, Sir.")
            return None

    def set_volume(self, level):
        try:
            level = max(0, min(100, level))
            volume = self.get_volume_interface()
            if volume:
                volume.SetMasterVolumeLevelScalar(level / 100, None)
                return f"Volume set to {level} percent"
            return "I'm sorry, Sir, but I cannot access the system volume controls at the moment."
        except Exception as e:
            return f"Failed to adjust the volume: {str(e)}"

    def adjust_volume(self, delta):
        try:
            volume = self.get_volume_interface()
            if volume:
                current = volume.GetMasterVolumeLevelScalar()
                new_level = max(0.0, min(1.0, current + delta))
                volume.SetMasterVolumeLevelScalar(new_level, None)
                return f"Volume adjusted to {int(new_level * 100)} percent"
            return "I'm sorry, Sir, but I cannot access the system volume controls."
        except Exception as e:
            return f"Error adjusting volume: {str(e)}"

    def open_app(self, name):
        try:
            installed_apps = [a.lower() for a in give_appnames()]; matches = difflib.get_close_matches(name.lower(), installed_apps, n=1, cutoff=0.6)
            if not matches: return "I’m sorry, Sir, I couldn’t find that application."
            open_app_cmd(matches[0], match_closest=True); return f"Opening {matches[0]}"
        except Exception: return "I’m sorry, Sir, I encountered an error while attempting to open the application."

    def control_browser_media(self, action):
        try:
            # Look for YouTube or Chrome windows
            windows = [w for w in gw.getAllWindows() if 'YouTube' in w.title or 'Google Chrome' in w.title]
            
            if windows:
                win = windows[0]
                try:
                    win.activate()
                except Exception:
                    pass # Already active or minimized
                
                time.sleep(0.3) # Ensure focus
                
                action = action.lower()
                if any(w in action for w in ["increase", "up", "volume up", "more volume"]): 
                    pyautogui.press('up', presses=5); return "Increasing playback volume"
                elif any(w in action for w in ["decrease", "down", "volume down", "less volume"]): 
                    pyautogui.press('down', presses=5); return "Decreasing playback volume"
                elif any(w in action for w in ["skip", "forward", "ahead", "next"]): 
                    pyautogui.press('l'); return "Skipping forward"
                elif any(w in action for w in ["rewind", "back", "previous", "return"]): 
                    pyautogui.press('j'); return "Rewinding"
                elif any(w in action for w in ["pause", "stop", "resume", "play", "toggle"]): 
                    pyautogui.press('k'); return "Playback toggled"
                elif any(w in action for w in ["full screen", "fullscreen", "maximize"]): 
                    pyautogui.press('f'); return "Full screen toggled"
                elif any(w in action for w in ["mute", "unmute", "silence"]): 
                    pyautogui.press('m'); return "Mute toggled"
            
            return "I could not find an active YouTube or Chrome session to control, Sir."
        except Exception as e:
            return f"Error during browser media control: {str(e)}, Sir."

    def record_audio(self, filename, duration=5):
        audio_data = []
        try:
            self.recorder.start(); start_time = time.time()
            while time.time() - start_time < duration:
                frame = self.recorder.read(); audio_data.extend(frame)
        finally: self.recorder.stop()
        with wave.open(filename, 'wb') as f:
            f.setnchannels(1); f.setsampwidth(2); f.setframerate(self.recorder.sample_rate); f.writeframes(struct.pack('h' * len(audio_data), *audio_data))
        return filename

    def wait_for_wake_word(self):
        print("Listening for 'Max' (Sir)...")
        wake_wav = self.record_audio("wake.wav", duration=4)
        try:
            with sr.AudioFile(wake_wav) as source:
                audio = self.recognizer.record(source); text = self.recognizer.recognize_google(audio).lower()
                print(f"Heard: '{text}'")
                is_max = any(word in text.split() for word in ["max", "macs", "maxwell", "make"])
                if is_max and "how are you" in text: return "emotive"
                elif is_max and "wake up" in text: return "initialize"
                elif is_max: return "standard"
        except: pass
        finally:
            if os.path.exists(wake_wav): os.remove(wake_wav)
        return None

    def capture_command(self):
        print("Listening for command (Sir)...")
        cmd_wav = self.record_audio("cmd.wav", duration=6)
        try:
            segments, info = self.stt_model.transcribe(cmd_wav, beam_size=5)
            command = "".join([segment.text for segment in segments]).strip(); return command
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
                        if wake_type == "emotive": self.speak("I am functioning at peak efficiency, Sir. Thank you for asking. How can I assist you today?")
                        elif wake_type == "initialize": self.speak("Systems initialized. I am online and ready, Sir")
                        else: self.speak("Yes, Sir? I am listening")
                
                if self.active_session:
                    command = self.capture_command()
                    if command and len(command) > 1:
                        print(f"Command Captured: '{command}'")
                        if self.pending_action:
                            if any(word in command.lower() for word in ["yes", "yeah", "sure", "do it", "confirm", "open it"]):
                                if self.pending_action["type"] == "web_open":
                                    query = self.pending_action["query"].lower().replace(" ", "")
                                    url = f"https://www.{query}.com" if "." not in query else f"https://{query}"
                                    webbrowser.open(url); self.speak(f"Navigating directly to {query}")
                                self.pending_action = None; continue
                            else:
                                self.speak("Action cancelled"); self.pending_action = None; continue
                        action_result = self.handle_local_commands(command)
                        thought_prompt = command
                        if action_result:
                            thought_prompt = f"System Update: I have successfully performed this action: {action_result}. Now respond to the user command: {command}"
                        emotive_response = self.think(thought_prompt)
                        standby_triggers = ["go to sleep", "stop listening", "rest", "stand by", "thank you", "thanks"]
                        if any(word in command.lower() for word in standby_triggers):
                            self.active_session = False
                            self.speak(emotive_response if emotive_response else "As you wish, Sir. Returning to standby")
                            continue
                        if emotive_response and "unstable" not in emotive_response:
                            self.speak(emotive_response)
                        elif action_result:
                            self.speak(action_result)
                        else:
                            self.speak("I'm afraid I didn't quite catch that, Sir.")
                        if emotive_response:
                            self.memory.append({"role": "user", "content": command})
                            self.memory.append({"role": "assistant", "content": emotive_response})
        except KeyboardInterrupt: sys.exit(0)

if __name__ == "__main__":
    assistant = MAXAssistant()
    assistant.run()
