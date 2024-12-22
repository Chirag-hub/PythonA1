import ollama  # Import Ollama package for conversational model
import speech_recognition as sr
import pyttsx3
import webbrowser
import os
import ctypes

# Initialize speech synthesizer
engine = pyttsx3.init()

# Set voice properties (you can change the voice and properties)
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)  # Change the index for different voices
engine.setProperty('rate', 210)  # Set speech rate

# Flag to prevent multiple command processing at once
is_processing_command = False
has_minimized = False

# Constants for minimizing and bringing console window to front
SW_MINIMIZE = 6
SW_RESTORE = 9
kernel32 = ctypes.windll.kernel32
user32 = ctypes.windll.user32

def minimize_console_window():
    """Minimize the console window."""
    hwnd = kernel32.GetConsoleWindow()
    if hwnd:
        user32.ShowWindow(hwnd, SW_MINIMIZE)

def restore_console_window():
    """Restore the console window if minimized and bring it to the front."""
    hwnd = kernel32.GetConsoleWindow()
    if hwnd:
        user32.ShowWindow(hwnd, SW_RESTORE)
        user32.SetForegroundWindow(hwnd)

def speak(text):
    """Use pyttsx3 to speak the provided text."""
    engine.say(text)
    engine.runAndWait()

def process_command(command):
    """Process the recognized command with conversational awareness."""
    global is_processing_command, has_minimized

    if "friday" in command:
        restore_console_window()  # Restore and bring the console to the front
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
        query = command.replace("google", "").strip()  # Remove "google" from the command
        if query:
            perform_google_search(query)
        else:
            speak("Please provide the search terms.")

    elif "what can you do" in command or "help" in command:
        speak("I can perform Google searches, minimize the window, or help with many other things. Just ask!")

    else:
        # Conversational response using Ollama's Llama 3.2:3B model
        respond_to_conversation(command)

def respond_to_conversation(command):
    """Use Ollama's conversational model (Llama 3.2:3B) to respond to user input."""
    try:
        # Use Ollama to get a response with the Llama 3.2:3B model
        response = ollama.chat(model="llama3.2:3b", messages=[{"role": "user", "content": command}])

        # Extract the response text from the response object
        bot_response = response.message.content

        # If the response is too long, return only the first meaningful sentence
        if len(bot_response.split('.')) > 1:
            bot_response = bot_response.split('.')[0] + "."

        # Ensure the response is concise and stripped of extra whitespace
        bot_response = bot_response.strip()

        # Print the response on the terminal
        print(f"Model Response: {bot_response}")

        # Speak the response
        speak(bot_response)

    except Exception as e:
        print(f"Error with Ollama response: {e}")
        speak("Sorry, there was an issue with the conversation model.")

def perform_google_search(query):
    """Perform a Google search."""
    if query:
        search_url = f"https://www.google.com/search?q={query}"
        webbrowser.open(search_url)
        speak(f"Performing Google search for: {query}")
    else:
        speak("No search term provided. Please try again.")

def listen_for_commands():
    """Listen for voice commands continuously."""
    global is_processing_command
    recognizer = sr.Recognizer()

    with sr.Microphone() as source:
        print("Listening for commands...")
        recognizer.adjust_for_ambient_noise(source)

        while True:
            try:
                audio = recognizer.listen(source)
                print("Recognizing...")
                command = recognizer.recognize_google(audio).lower().strip()
                print(f"Recognized: {command}")

                if not is_processing_command:
                    is_processing_command = True
                    process_command(command)
                    is_processing_command = False
            except sr.UnknownValueError:
                print("Sorry, I could not understand that.")
            except sr.RequestError:
                print("Sorry, there was an issue with the speech recognition service.")

def main():
    """Main entry point of the application."""
    try:
        print("Voice command application is running. Say 'FRIDAY' to start.")
        speak("Voice command application is running. Say 'FRIDAY' to start.")
        listen_for_commands()
    except Exception as e:
        print(f"An error occurred: {e}")
        speak(f"An error occurred: {e}")

# Ensure this block is in place
if __name__ == "__main__":
    main()
