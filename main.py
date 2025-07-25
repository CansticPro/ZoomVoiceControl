import keyboard
import pyautogui
import speech_recognition as sr
import threading
import time

# For real volume control
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL

# Get system volume interface
devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))

# ----- Zoom Control Functions -----
def hold_mute():
    keyboard.press('alt+a')
    print("🔇 Holding Mute for 10s...")
    time.sleep(10)
    keyboard.release('alt+a')
    print("🔈 Released Mute")

def increase_volume():
    current_volume = volume.GetMasterVolumeLevelScalar()
    new_volume = min(current_volume + 0.1, 1.0)
    volume.SetMasterVolumeLevelScalar(new_volume, None)
    print(f"🔊 Volume increased to {int(new_volume * 100)}%")
    time.sleep(0.2)

def decrease_volume():
    current_volume = volume.GetMasterVolumeLevelScalar()
    new_volume = max(current_volume - 0.1, 0.0)
    volume.SetMasterVolumeLevelScalar(new_volume, None)
    print(f"🔉 Volume decreased to {int(new_volume * 100)}%")
    time.sleep(0.2)

def toggle_mute():
    keyboard.press_and_release('alt+a')
    print("🔇 Toggled Mute/Unmute")
    time.sleep(0.2)

def toggle_camera():
    keyboard.press_and_release('alt+v')
    print("🎥 Toggled Camera")
    time.sleep(0.2)

def toggle_hand():
    keyboard.press_and_release('alt+y')
    print("✋ Toggled Raise/Lower Hand")
    time.sleep(0.2)

def toggle_chat():
    keyboard.press_and_release('alt+h')
    print("💬 Toggled Chat")
    time.sleep(0.2)

def toggle_participants():
    keyboard.press_and_release('alt+u')
    print("👥 Toggled Participants")
    time.sleep(0.2)

def leave_meeting():
    keyboard.press_and_release('alt+q')
    time.sleep(0.5)
    keyboard.press_and_release('enter')
    print("🚪 Left Meeting")

def double_click_screen():
    pyautogui.doubleClick()
    print("🖱️ Double-clicked Screen")
    time.sleep(0.2)

def press_escape():
    keyboard.press_and_release('esc')
    print("⎋ Escape Key Pressed")
    time.sleep(0.2)

# ----- Voice Command Handler -----
def handle_command(command):
    command = command.lower()

    if any(word in command for word in ['mute', 'unmute']):
        toggle_mute()
    elif any(word in command for word in ['volume up', 'increase volume', 'increase', 'louder']):
        increase_volume()
    elif any(word in command for word in ['volume down', 'decrease volume', 'decrease', 'quieter', 'lower']):
        decrease_volume()
    elif any(word in command for word in ['hold mute', 'hold']):
        hold_mute()
    elif any(word in command for word in ['camera', 'video']):
        toggle_camera()
    elif any(word in command for word in ['raise hand', 'hand up', 'hand']):
        toggle_hand()
    elif any(word in command for word in ['lower hand', 'unraise']):
        toggle_hand()
    elif any(word in command for word in ['chat', 'message']):
        toggle_chat()
    elif any(word in command for word in ['participants', 'people']):
        toggle_participants()
    elif any(word in command for word in ['leave', 'exit']):
        leave_meeting()
    elif any(word in command for word in ['time', 'escape']):
        press_escape()
    elif any(word in command for word in ['full screen', 'fullscreen']):
        double_click_screen()
    else:
        print(f"❓ Unrecognized command: {command}")

# ----- Voice Recognition Listener -----
def listen_voice_commands():
    recognizer = sr.Recognizer()
    mic = sr.Microphone()

    with mic:
        recognizer.adjust_for_ambient_noise(mic)
        print("🎙️ Voice control started. Say a command...")

    while True:
        try:
            with mic:
                audio = recognizer.listen(mic, timeout=5)
            command = recognizer.recognize_google(audio)
            print(f"🗣️ Heard: {command}")
            handle_command(command)
        except sr.WaitTimeoutError:
            continue
        except sr.UnknownValueError:
            print("❌ Couldn’t understand you.")
        except Exception as e:
            print(f"⚠️ Error: {e}")

# Start listener in background
threading.Thread(target=listen_voice_commands, daemon=True).start()

# Keep script running
while True:
    time.sleep(1)
