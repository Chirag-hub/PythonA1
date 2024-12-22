import ollama
import speech_recognition as sr
import pyttsx3
import webbrowser
import os
import ctypes
import json
from datetime import datetime

# Initialize text-to-speech engine
engine = pyttsx3.init()
voices = engine.getProperty('voices')

# List available voices
for index, voice in enumerate(voices):
    print(f"Voice {index}: {voice.name}")

# Set to a specific voice (change index as needed)
engine.setProperty('voice', voices[1].id)  # Example: Selecting the second voice
engine.setProperty('rate', 210)

# Constants for minimizing and restoring console window
SW_MINIMIZE, SW_RESTORE = 6, 9
kernel32, user32 = ctypes.windll.kernel32, ctypes.windll.user32

# Global states and memory
is_processing_command, has_minimized, is_conversation_paused = False, False, False
conversation_memory = []
conversation_file = "conversation_history.json"

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
    print(text)
    engine.say(text)
    engine.runAndWait()

# Search history for a topic on a specific day
def search_history_for_day(topic, day):
    global conversation_memory
    results = []
    for entry in conversation_memory:
        entry_date = datetime.strptime(entry["timestamp"], "%Y-%m-%d %H:%M:%S")
        if topic.lower() in entry["content"].lower() and day.lower() == entry_date.strftime("%A").lower():
            results.append(entry)

    if results:
        speak(f"Here are the discussions about '{topic}' on {day}:")
        for result in results:
            speak(result["content"])
    else:
        speak(f"No discussions found about '{topic}' on {day}.")

# Process user commands
def process_command(command):
    global is_processing_command, has_minimized, is_conversation_paused

    if "friday" in command:
        restore_console()
        query = command.replace("friday", "").strip()
        if "search history" in query:
            query_parts = query.replace("search history", "").strip().split(" on ")
            topic = query_parts[0].strip() if len(query_parts) > 0 else ""
            day = query_parts[1].strip() if len(query_parts) > 1 else ""
            if topic and day:
                search_history_for_day(topic, day)
            elif topic:
                speak("Please specify a day to search for the topic.")
            else:
                speak("Please specify a topic and a day to search in history.")
        else:
            speak("Yes, how can I assist you?" if not query else f"Searching for: {query}")
            if query: perform_google_search(query)

    elif "exit" in command:
        speak("Goodbye!")
        save_conversation()
        os._exit(0)

    elif ("minimize" in command or "minimise" in command) and not has_minimized:
        has_minimized = True
        minimize_console()
        speak("Window minimized.")

    elif "google" in command:
        query = command.replace("google", "").strip()
        speak("Provide search terms." if not query else f"Searching Google for: {query}")
        if query: perform_google_search(query)

    elif "help" in command:
        speak("I can search, minimize, or chat. Just ask!")

    elif "pause" in command or "stop" in command:
        is_conversation_paused = True
        speak("Conversation paused.")

    else:
        respond_to_conversation(command)

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
    listen_for_commands()

# Listen for user commands
def listen_for_commands():
    global is_processing_command, is_conversation_paused
    recognizer = sr.Recognizer()

    with sr.Microphone() as source:
        print("Listening for commands...")
        recognizer.adjust_for_ambient_noise(source)

        while True:
            try:
                command = recognizer.recognize_google(recognizer.listen(source)).lower().strip()
                print(f"Recognized: {command}")
                if not is_processing_command and not is_conversation_paused:
                    is_processing_command = True
                    process_command(command)
                    is_processing_command = False
                elif "unpause" in command or "resume" in command:
                    resume_conversation()
            except sr.UnknownValueError:
                print("Could not understand.")
            except sr.RequestError:
                print("Recognition service error.")

# Main function
def main():
    try:
        load_conversation()
        speak("Voice assistant running. Say 'FRIDAY' to start.")
        listen_for_commands()
    except Exception as e:
        speak(f"Error: {e}")

if __name__ == "__main__":
    main()
