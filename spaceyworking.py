import ollama
import speech_recognition as sr
import pyttsx3
import webbrowser
import os
import ctypes
import spacy
import re

# Initialize NLP model
nlp = spacy.load("en_core_web_sm")

# Initialize speech synthesizer
engine = pyttsx3.init()
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)
engine.setProperty('rate', 210)

SW_MINIMIZE, SW_RESTORE = 6, 9
kernel32, user32 = ctypes.windll.kernel32, ctypes.windll.user32

is_processing_command, has_minimized = False, False

# Function to minimize the console window
def minimize_console():
    hwnd = kernel32.GetConsoleWindow()
    if hwnd:
        user32.ShowWindow(hwnd, SW_MINIMIZE)

# Function to restore the console window
def restore_console():
    hwnd = kernel32.GetConsoleWindow()
    if hwnd:
        user32.ShowWindow(hwnd, SW_RESTORE)
        user32.SetForegroundWindow(hwnd)

# Text-to-Speech Function
def speak(text):
    """Convert text to speech, ensuring it doesn't contain unwanted characters."""
    clean_text = sanitize_text(text)  # Clean the text before speaking
    engine.say(clean_text)
    engine.runAndWait()

# Function to sanitize input (remove unwanted characters)
def sanitize_text(text):
    """Remove non-alphabetic characters and symbols from the input."""
    # Removing unwanted special characters (e.g., '*' or others)
    sanitized_text = re.sub(r"[^a-zA-Z0-9\s,.!?]", "", text)  # Keep letters, numbers, and basic punctuation
    return sanitized_text

# Function to parse the user's command
def parse_command(command):
    doc = nlp(command)
    intent = "general"
    entities = []

    if "google" in command or "search" in command:
        intent = "search"
        entities = [ent.text for ent in doc.ents]
    elif "minimize" in command or "minimise" in command:
        intent = "minimize"
    elif "exit" in command:
        intent = "exit"
    elif "help" in command:
        intent = "help"

    return intent, entities

# Function to process a command
def process_command(command):
    global is_processing_command, has_minimized

    intent, entities = parse_command(command)

    if intent == "search":
        query = " ".join(entities) if entities else command.replace("google", "").replace("search", "").strip()
        if query:
            perform_google_search(query)
        else:
            speak("What would you like to search for?")

    elif intent == "minimize" and not has_minimized:
        has_minimized = True
        minimize_console()
        speak("Window minimized. Let me know if you need anything else.")

    elif intent == "exit":
        speak("Goodbye!")
        os._exit(0)

    elif intent == "help":
        speak("I can perform searches, minimize windows, or answer questions. Just ask!")

    else:
        respond_to_conversation(command)

# Function to respond to conversational inputs
def respond_to_conversation(command):
    try:
        print("Processing user command:", command)  # Debug print
        response = ollama.chat(model="llama3.2:3b", messages=[{"role": "user", "content": command}])
        reply = response['message']['content']  # Fix to access the correct content in the response
        print("Response from Ollama:", reply)  # Debug print
        speak(reply)
    except Exception as e:
        speak(f"Sorry, I encountered an issue: {e}")

# Function to perform Google search
def perform_google_search(query):
    webbrowser.open(f"https://www.google.com/search?q={query}")
    speak(f"Searching Google for: {query}")

# Function to listen for commands
def listen_for_commands():
    global is_processing_command
    recognizer = sr.Recognizer()

    with sr.Microphone() as source:
        recognizer.adjust_for_ambient_noise(source, duration=1)  # Adjust for ambient noise

        while True:
            try:
                audio = recognizer.listen(source, timeout=5, phrase_time_limit=10)
                command = recognizer.recognize_google(audio).lower().strip()

                if "friday" in command:
                    restore_console()
                    command = command.replace("friday", "").strip()

                    if not is_processing_command:
                        is_processing_command = True
                        process_command(command)  # Process commands synchronously
                        is_processing_command = False

            except sr.UnknownValueError:
                # Ignore background noise
                pass
            except sr.RequestError as e:
                speak(f"Speech recognition service error: {e}")
            #except Exception as e:
             #   speak(f"An unexpected error occurred: {e}")

# Main function to start the assistant
def main():
    try:
        print("Voice assistant running. Say 'FRIDAY' to activate.")
        speak("Voice assistant running. Say 'FRIDAY' to activate.")
        listen_for_commands()  # Start listening synchronously
    except Exception as e:
        speak(f"An error occurred: {e}")

if __name__ == "__main__":
    main()  # Run the program synchronously