import ollama  # Import Ollama package for conversational model
import speech_recognition as sr
import pyttsx3
import webbrowser
import os
import ctypes
import sounddevice as sd
import queue
import json
import vosk

# Initialize speech synthesizer
engine = pyttsx3.init()

# Set voice properties
voices = engine.getProperty("voices")
engine.setProperty("voice", voices[0].id)  # Change the index for different voices
engine.setProperty("rate", 210)  # Set speech rate

# Flags and constants
is_processing_command = False
has_minimized = False
SW_MINIMIZE = 6
SW_RESTORE = 9
kernel32 = ctypes.windll.kernel32
user32 = ctypes.windll.user32

# Audio queue for Vosk
audio_queue = queue.Queue()

# Globals for selected model
speech_to_text_model = None
vosk_model = None
vosk_recognizer = None

# Initialize Vosk model
VOSK_MODEL_PATH = "./vosk-model-small-en-us-0.15"  # Update to your Vosk model path

# Function to minimize and restore console

def minimize_console_window():
    hwnd = kernel32.GetConsoleWindow()
    if hwnd:
        user32.ShowWindow(hwnd, SW_MINIMIZE)

def restore_console_window():
    hwnd = kernel32.GetConsoleWindow()
    if hwnd:
        user32.ShowWindow(hwnd, SW_RESTORE)
        user32.SetForegroundWindow(hwnd)

# Speech functions

def speak(text):
    engine.say(text)
    engine.runAndWait()

def process_command(command):
    global is_processing_command, has_minimized

    if "friday" in command:
        restore_console_window()
        query = command.replace("friday", "").strip()
        if not query:
            speak("Yes, how can I assist you?")
        else:
            perform_google_search(query)

    elif "exit" in command:
        speak("Terminating the application. Goodbye!")
        os._exit(0)

    elif ("minimize" in command or "minimise" in command) and not has_minimized:
        has_minimized = True
        minimize_console_window()
        speak("The window has been minimized. Let me know if you need anything else.")

    elif "google" in command:
        query = command.replace("google", "").strip()
        if query:
            perform_google_search(query)
        else:
            speak("Please provide the search terms.")

    elif "what can you do" in command or "help" in command:
        speak("I can perform Google searches, minimize the window, or help with many other things. Just ask!")

    else:
        respond_to_conversation(command)

def respond_to_conversation(command):
    try:
        response = ollama.chat(model="llama3.2:3b", messages=[{"role": "user", "content": command}])
        bot_response = response.message.content
        if len(bot_response.split('.')) > 1:
            bot_response = bot_response.split('.')[0] + "."
        bot_response = bot_response.strip()
        print(f"Model Response: {bot_response}")
        speak(bot_response)
    except Exception as e:
        print(f"Error with Ollama response: {e}")
        speak("Sorry, there was an issue with the conversation model.")

def perform_google_search(query):
    if query:
        search_url = f"https://www.google.com/search?q={query}"
        webbrowser.open(search_url)
        speak(f"Performing Google search for: {query}")
    else:
        speak("No search term provided. Please try again.")

# Speech-to-text functions

def initialize_speech_to_text():
    global speech_to_text_model, vosk_model, vosk_recognizer

    try:
        # Test internet connectivity and SpeechRecognition
        recognizer = sr.Recognizer()
        with sr.Microphone() as source:
            print("Testing online SpeechRecognition...")
            recognizer.adjust_for_ambient_noise(source)
            recognizer.listen(source, timeout=3)  # Timeout if no audio
            speech_to_text_model = "speechrecognition"
            print("Online SpeechRecognition is available.")
            return
    except Exception as e:
        print(f"Online SpeechRecognition failed: {e}")

    # Fallback to Vosk if offline
    try:
        print("Loading Vosk model...")
        vosk_model = vosk.Model(VOSK_MODEL_PATH)
        vosk_recognizer = vosk.KaldiRecognizer(vosk_model, 16000)
        speech_to_text_model = "vosk"
        print("Vosk is ready for offline speech recognition.")
    except Exception as e:
        print(f"Failed to load Vosk model: {e}")
        speak("No speech recognition model is available. Exiting.")
        os._exit(1)

def listen_with_speechrecognition():
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening with SpeechRecognition...")
        audio = recognizer.listen(source)
        command = recognizer.recognize_google(audio).lower().strip()
        print(f"Recognized: {command}")
        return command

def listen_with_vosk():
    with sd.RawInputStream(samplerate=16000, blocksize=8000, dtype="int16", channels=1, callback=lambda indata, frames, time, status: audio_queue.put(bytes(indata))):
        print("Listening with Vosk...")
        while True:
            data = audio_queue.get()
            if vosk_recognizer.AcceptWaveform(data):
                result = json.loads(vosk_recognizer.Result())
                command = result.get("text", "").lower().strip()
                print(f"Recognized: {command}")
                return command

def listen_for_commands():
    global is_processing_command

    while True:
        try:
            if speech_to_text_model == "speechrecognition":
                command = listen_with_speechrecognition()
            elif speech_to_text_model == "vosk":
                command = listen_with_vosk()
            else:
                raise Exception("No valid speech-to-text model selected.")

            if command and not is_processing_command:
                is_processing_command = True
                process_command(command)
                is_processing_command = False

        except Exception as e:
            print(f"Error during command processing: {e}")

# Main entry point

def main():
    try:
        print("Initializing speech-to-text model...")
        initialize_speech_to_text()
        print("Voice command application is running. Say 'FRIDAY' to start.")
        speak("Voice command application is running. Say 'FRIDAY' to start.")
        listen_for_commands()
    except Exception as e:
        print(f"An error occurred: {e}")
        speak(f"An error occurred: {e}")

if __name__ == "__main__":
    main()