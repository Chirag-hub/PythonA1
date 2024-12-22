import vosk
import json
import queue
import sounddevice as sd
import os

# Ensure the Vosk model path is valid
model_path = "D:\\AI dev\\vosk-model-en-in-0.5"
if not os.path.exists(model_path):
    raise FileNotFoundError(f"Model not found at {model_path}")

# Load the Vosk model
model = vosk.Model(model_path)
recognizer_queue = queue.Queue()

# Audio callback function
def vosk_callback(indata, frames, time, status):
    if status:
        print(f"Stream status warning: {status}")
    try:
        # Convert numpy.ndarray to bytes and put into the queue
        recognizer_queue.put(indata.tobytes())
    except Exception as e:
        print(f"Error in callback: {e}")

# Function to listen for commands
def listen_for_commands():
    rec = vosk.KaldiRecognizer(model, 16000)  # Initialize the recognizer
    try:
        with sd.InputStream(samplerate=16000, channels=1, dtype="int16", callback=vosk_callback):
            print("Listening for commands...")
            while True:
                try:
                    # Get data from the queue
                    data = recognizer_queue.get(timeout=1)
                    if rec.AcceptWaveform(data):  # Process complete speech
                        result = json.loads(rec.Result())
                        command = result.get("text", "").strip()
                        if command:
                            print(f"Recognized: {command}")
                        else:
                            print("No command recognized.")
                    else:
                        print(f"Partial recognition: {rec.PartialResult()}")
                except queue.Empty:
                    pass  # Continue waiting for data
    except Exception as e:
        print(f"Error in recognition loop: {e}")

# Main function
if __name__ == "__main__":
    try:
        listen_for_commands()
    except Exception as e:
        print(f"Fatal error: {e}")
