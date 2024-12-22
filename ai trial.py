import ollama
import vosk
import pyttsx3
import webbrowser
import os
import ctypes
import json
import queue
import sounddevice as sd
import numpy as np
import noisereduce as nr
from datetime import datetime

# Initialize text-to-speech engine
engine = pyttsx3.init()
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)  # Set to the second voice
engine.setProperty('rate', 210)

# Constants for minimizing and restoring console window
SW_MINIMIZE, SW_RESTORE = 6, 9
kernel32, user32 = ctypes.windll.kernel32, ctypes.windll.user32

# Global states and memory
is_processing_command, has_minimized, is_conversation_paused = False, False, False
conversation_memory = []
conversation_file = "conversation_history.json"

# Load Vosk Model
model_path = "vosk-model-en-in-0.5"  # Change this to your actual Vosk model path
if not os.path.exists(model_path):
    raise FileNotFoundError(f"Please download the Vosk model and place it in the '{model_path}' folder.")
model = vosk.Model(model_path)
recognizer_queue = queue.Queue()

# Load conversation history
def load_conversation():
    global conversation_memory
    if os.path.exists(conversation_file):
        try:
            with open(conversation_file, "r") as file:
                conversation_memory = json.load(file)
                print("Previous conversation history loaded.")
        except Exception as e:
            print(f"Error loading conversation history: {e}")

# Save conversation history
def save_conversation():
    global conversation_memory
    try:
        with open(conversation_file, "w") as file:
            json.dump(conversation_memory, file, indent=4)
    except Exception as e:
        print(f"Error saving conversation history: {e}")

# Console window management
def minimize_console():
    hwnd = kernel32.GetConsoleWindow()
    if hwnd:
        user32.ShowWindow(hwnd, SW_MINIMIZE)

def restore_console():
    hwnd = kernel32.GetConsoleWindow()
    if hwnd:
        user32.ShowWindow(hwnd, SW_RESTORE)
        user32.SetForegroundWindow(hwnd)

# Speak text aloud
def speak(text):
    print(f"[Assistant]: {text}")
    engine.say(text)
    engine.runAndWait()

# Handle conversation responses
def respond_to_conversation(command):
    global conversation_memory
    try:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        conversation_memory.append({"timestamp": timestamp, "role": "user", "content": command})
        response = ollama.chat(model="llama3.2:3b", messages=conversation_memory)
        conversation_memory.append({"timestamp": timestamp, "role": "assistant", "content": response.message.content})
        speak(response.message.content)
        save_conversation()
    except Exception as e:
        speak("Error with conversation model.")

# Perform a Google search
def perform_google_search(query):
    webbrowser.open(f"https://www.google.com/search?q={query}")

# Resume paused conversation
def resume_conversation():
    global is_conversation_paused
    is_conversation_paused = False
    speak("Conversation resumed.")

# Process user commands
def process_command(command):
    global is_processing_command, has_minimized, is_conversation_paused
    try:
        if "exit" in command or "quit" in command or "goodbye" in command:
            speak("Goodbye! Saving conversation history.")
            save_conversation()
            os._exit(0)

        elif "friday" in command:
            restore_console()
            query = command.replace("friday", "").strip()
            speak("Yes, how can I assist you?" if not query else f"Searching for: {query}")
            if query:
                perform_google_search(query)

        elif ("minimize" in command or "minimise" in command) and not has_minimized:
            has_minimized = True
            minimize_console()
            speak("Window minimized.")

        elif "google" in command:
            query = command.replace("google", "").strip()
            speak("Provide search terms." if not query else f"Searching Google for: {query}")
            if query:
                perform_google_search(query)

        elif "help" in command:
            speak("I can search, minimize, or chat. Just ask!")

        elif "pause" in command or "stop" in command:
            is_conversation_paused = True
            speak("Conversation paused.")

        elif "unpause" in command or "resume" in command:
            resume_conversation()

        else:
            respond_to_conversation(command)

    except Exception as e:
        print(f"Error processing command: {e}")
        speak("I encountered an error.")

# Apply noise reduction to the incoming audio stream
def reduce_noise(indata):
    try:
        # Apply noise reduction only if the input is valid (ensure no NaN or infinity)
        indata = np.nan_to_num(indata)  # Replace NaN or infinite values with 0
        reduced_noise = nr.reduce_noise(y=indata, sr=16000)
        return reduced_noise
    except Exception as e:
        print(f"Error in noise reduction: {e}")
        return indata  # Return original if error occurs

# Vosk audio callback for live recognition
def vosk_callback(indata, frames, time, status):
    if status:
        print(f"Stream status: {status}")  # If there's an issue with the stream, print status
    # Process the input data, applying noise reduction
    reduced_noise = reduce_noise(indata)
    recognizer_queue.put(bytes(reduced_noise))  # Add reduced noise data to the queue

# Listen for commands using Vosk
def listen_for_commands():
    rec = vosk.KaldiRecognizer(model, 16000)  # Initialize recognizer for the model
    with sd.InputStream(samplerate=16000, channels=1, dtype="int16", blocksize=512, callback=vosk_callback):
        print("Listening for commands...")
        while True:
            try:
                data = recognizer_queue.get()  # Get audio data from the queue
                if rec.AcceptWaveform(data):  # Process recognized audio
                    result = json.loads(rec.Result())  # Parse the recognized result
                    text = result.get("text", "").strip()  # Extract text from result
                    if text:
                        print(f"Recognized: {text}")
                        process_command(text)  # Process the command
            except Exception as e:
                print(f"Error in processing audio: {e}")

# Main function
def main():
    try:
        load_conversation()
        speak("Voice assistant running. Say 'FRIDAY' to start.")
        listen_for_commands()
    except Exception as e:
        print(f"Error in main function: {e}")
        speak(f"Error: {e}")

if __name__ == "__main__":
    main()
