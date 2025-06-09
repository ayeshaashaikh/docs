# import json
# import threading
# import time
# import tkinter as tk

# import pyaudio
# import vosk
# import win32com.client as win32


# def resource_path(relative_path):
#     """
#     Get the absolute path to a resource, whether running in a PyInstaller bundle or normally.
#     """
#     try:
#         base_path = sys._MEIPASS  # PyInstaller temporary folder
#     except Exception:
#         base_path = os.path.abspath(".")
#     return os.path.join(base_path, relative_path)
# # Initialize PowerPoint and ensure it’s visible.
# powerpoint = win32.Dispatch("PowerPoint.Application")
# powerpoint.Visible = True  

# # Global flag and overlay instance for subtitles.
# subtitle_active = False
# overlay_instance = None
# audio_thread = None
# status_thread = None

# def create_overlay(parent=None):
#     """
#     Create an overlay window.
#     If a parent is provided, use a Toplevel so that the overlay
#     is managed by the main application’s mainloop.
#     """
#     if parent:
#         overlay = tk.Toplevel(parent)
#     else:
#         overlay = tk.Tk()
#     overlay.title("Subtitles")
#     overlay.geometry("800x50")
#     overlay.attributes("-topmost", True)
#     overlay.configure(bg="black")
#     overlay.overrideredirect(True)
#     overlay.attributes("-alpha", 0.8)
    
#     # Create a label that fills the overlay.
#     label = tk.Label(overlay, text="", font=("Helvetica", 24), fg="white", bg="black", anchor="center")
#     label.pack(expand=True, fill='both')
#     return overlay, label

# def listen_and_display_subtitles(label):
#     model = vosk.Model(r"C:\Users\Acer\Documents\VOS\model\vosk-model-small-en-us-0.15")
#     recognizer = vosk.KaldiRecognizer(model, 16000)

#     p = pyaudio.PyAudio()
#     stream = p.open(format=pyaudio.paInt16, channels=1, rate=16000,
#                     input=True, frames_per_buffer=8000)
#     stream.start_stream()

#     print("Listening for voice input...")
#     while subtitle_active:
#         try:
#             # Read a chunk from the stream; using exception_on_overflow=False
#             data = stream.read(4000, exception_on_overflow=False)
#         except Exception as e:
#             print("Error reading audio stream:", e)
#             break
#         if recognizer.AcceptWaveform(data):
#             result = recognizer.Result()
#             try:
#                 text = json.loads(result).get('text', '').capitalize()
#             except Exception as e:
#                 text = ""
#             if text:
#                 print(f"Recognized Text: {text}")
#                 # Update the label in a thread-safe manner using after().
#                 label.after(0, lambda: label.config(text=text))
#         # A very short sleep yields CPU time.
#         time.sleep(0.01)
#     stream.stop_stream()
#     stream.close()
#     p.terminate()
#     print("Stopped listening for subtitles.")

# def check_presentation_status(overlay):
#     while subtitle_active:
#         try:
#             if powerpoint.SlideShowWindows.Count > 0:
#                 overlay.after(0, overlay.deiconify)
#             else:
#                 overlay.after(0, overlay.withdraw)
#         except Exception as e:
#             print("Error checking presentation status:", e)
#         time.sleep(1)
#     print("Stopped checking presentation status.")

# def start_subtitle(parent=None):
#     """
#     Start the subtitle functionality.
#     If a parent is provided, create a Toplevel attached to it.
#     Do not call a separate mainloop if the parent exists.
#     """
#     global subtitle_active, overlay_instance, audio_thread, status_thread
#     subtitle_active = True

#     # Create the overlay in the main thread.
#     overlay, label = create_overlay(parent)
#     overlay_instance = overlay
#     overlay.update_idletasks()

#     # Position the overlay to span the full width of the screen and increased height.
#     screen_width = overlay.winfo_screenwidth()
#     screen_height = overlay.winfo_screenheight()
#     overlay_height = 100        # Increased height
#     overlay_width = screen_width  # Full screen width
#     x_position = 0                # Start at the left edge.
#     y_position = screen_height - overlay_height - 50
#     overlay.geometry(f"{overlay_width}x{overlay_height}+{x_position}+{y_position}")
    
#     # Update label wrap length so that long text wraps within the screen width.
#     label.config(wraplength=screen_width)

#     # Start background threads for audio listening and presentation status.
#     audio_thread = threading.Thread(target=listen_and_display_subtitles, args=(label,), daemon=True)
#     audio_thread.start()

#     status_thread = threading.Thread(target=check_presentation_status, args=(overlay,), daemon=True)
#     status_thread.start()

#     # Only call mainloop if no parent is provided (i.e. standalone use).
#     if not parent:
#         overlay.mainloop()

# def stop_subtitle():
#     global subtitle_active, overlay_instance
#     subtitle_active = False
#     if overlay_instance:
#         # Schedule destruction of the overlay on the main thread.
#         overlay_instance.after(0, overlay_instance.destroy)
#         overlay_instance = None

# For testing purposes, you could uncomment the following lines:
# start_subtitle()
import json
import os
import sys
import threading
import time
import tkinter as tk
from queue import Empty, Queue

import pyaudio
import vosk
import win32com.client as win32


def resource_path(relative_path):
    """Get absolute path to resource."""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# Initialize PowerPoint with error handling
powerpoint = None
try:
    powerpoint = win32.Dispatch("PowerPoint.Application")
    powerpoint.Visible = True
    print("✓ PowerPoint connected")
except Exception as e:
    print(f"⚠ PowerPoint not available: {e}")

# Global variables
subtitle_active = False
overlay_instance = None
audio_thread = None
display_queue = Queue()

class UltraFastSubtitles:
    def __init__(self):
        self.model = None
        self.recognizer = None
        self.stream = None
        self.p = None
        self.last_text = ""
        
    def initialize_audio(self):
        """Initialize audio with optimized settings for speed."""
        try:
            # Use small model for maximum speed
            model_path = r"C:\Users\Acer\Documents\VOS\model\vosk-model-small-en-us-0.15"
            if not os.path.exists(model_path):
                print("⚠ Small model not found, trying large model...")
                model_path = r"C:\Users\Acer\Documents\VOS\model\vosk-model-en-us-0.22"
            
            self.model = vosk.Model(model_path)
            print(f"✓ Model loaded: {model_path}")
            
            # Ultra-fast recognizer settings
            self.recognizer = vosk.KaldiRecognizer(self.model, 16000)
            self.recognizer.SetWords(False)  # No word timestamps for speed
            
            # Optimized audio stream
            self.p = pyaudio.PyAudio()
            
            # Find best input device
            default_device = self.p.get_default_input_device_info()
            device_index = default_device['index']
            
            self.stream = self.p.open(
                format=pyaudio.paInt16,
                channels=1,
                rate=16000,
                input=True,
                frames_per_buffer=512,  # Smallest buffer for instant response
                input_device_index=device_index
            )
            self.stream.start_stream()
            print("🚀 Audio stream started - MAXIMUM SPEED MODE")
            return True
            
        except Exception as e:
            print(f"❌ Audio initialization failed: {e}")
            return False
    
    def process_audio_ultra_fast(self):
        """Ultra-fast audio processing with zero delays."""
        print("🎤 Listening... Speak NOW!")
        
        buffer = b''
        chunk_size = 256  # Even smaller chunks for instant response
        
        while subtitle_active:
            try:
                # Read tiny chunks immediately
                data = self.stream.read(chunk_size, exception_on_overflow=False)
                buffer += data
                
                # Process when we have enough data
                if len(buffer) >= 1024:
                    if self.recognizer.AcceptWaveform(buffer):
                        # Final result - immediate processing
                        result = self.recognizer.Result()
                        self.handle_final_result(result)
                    else:
                        # Partial result - show immediately
                        partial = self.recognizer.PartialResult()
                        self.handle_partial_result(partial)
                    
                    buffer = b''  # Clear buffer
                
            except Exception as e:
                print(f"Audio error: {e}")
                continue
        
        self.cleanup_audio()
    
    def handle_final_result(self, result):
        """Handle final recognition result with minimal processing."""
        try:
            data = json.loads(result)
            text = data.get('text', '').strip()
            
            if text and text != self.last_text:
                # Minimal processing for maximum speed
                display_text = text.capitalize()
                self.last_text = text
                
                # Queue for immediate display
                display_queue.put(('final', display_text))
                print(f"✓ {display_text}")
                
        except json.JSONDecodeError:
            pass
    
    def handle_partial_result(self, partial_result):
        """Handle partial results for real-time display."""
        try:
            data = json.loads(partial_result)
            text = data.get('partial', '').strip()
            
            if text:
                # Show partial immediately with minimal processing
                display_text = text.capitalize()
                display_queue.put(('partial', display_text))
                
        except json.JSONDecodeError:
            pass
    
    def cleanup_audio(self):
        """Clean up audio resources."""
        try:
            if self.stream:
                self.stream.stop_stream()
                self.stream.close()
            if self.p:
                self.p.terminate()
            print("⏹ Audio cleanup complete")
        except Exception as e:
            print(f"Cleanup error: {e}")

def create_ultra_fast_overlay(parent=None):
    """Create optimized overlay for maximum display speed."""
    if parent:
        overlay = tk.Toplevel(parent)
    else:
        overlay = tk.Tk()
    
    overlay.title("Ultra-Fast Subtitles")
    overlay.attributes("-topmost", True)
    overlay.configure(bg="black")
    overlay.overrideredirect(True)
    overlay.attributes("-alpha", 0.9)
    
    # Position overlay
    screen_width = overlay.winfo_screenwidth()
    screen_height = overlay.winfo_screenheight()
    
    overlay_width = min(screen_width - 100, 1200)
    overlay_height = 80
    x_pos = (screen_width - overlay_width) // 2
    y_pos = screen_height - overlay_height - 50
    
    overlay.geometry(f"{overlay_width}x{overlay_height}+{x_pos}+{y_pos}")
    
    # Optimized label for instant updates
    label = tk.Label(
        overlay,
        text="Ready to speak...",
        font=("Segoe UI", 28, "bold"),
        fg="white",
        bg="black",
        anchor="center",
        wraplength=overlay_width - 20,
        justify='center'
    )
    label.pack(expand=True, fill='both', padx=10, pady=10)
    
    return overlay, label

def update_display_ultra_fast(label):
    """Ultra-fast display updates with zero delay."""
    while subtitle_active:
        try:
            # Get display updates immediately
            msg_type, text = display_queue.get(timeout=0.01)
            
            if msg_type == 'final':
                # Final text in white
                label.config(text=text, fg="white")
            elif msg_type == 'partial':
                # Partial text in cyan for real-time feedback
                label.config(text=text, fg="cyan")
            
            # Force immediate display update
            label.update_idletasks()
            
        except Empty:
            continue
        except Exception as e:
            print(f"Display error: {e}")
            continue

def start_ultra_fast_subtitles(parent=None):
    """Start ultra-fast real-time subtitles with zero delay."""
    global subtitle_active, overlay_instance, audio_thread
    
    print("🚀 STARTING ULTRA-FAST MODE - ZERO DELAY!")
    subtitle_active = True
    
    # Create overlay
    overlay, label = create_ultra_fast_overlay(parent)
    overlay_instance = overlay
    
    # Initialize audio processor
    processor = UltraFastSubtitles()
    if not processor.initialize_audio():
        print("❌ Failed to initialize audio")
        return
    
    # Start audio processing thread
    audio_thread = threading.Thread(
        target=processor.process_audio_ultra_fast,
        daemon=True
    )
    audio_thread.start()
    
    # Start display update thread
    display_thread = threading.Thread(
        target=update_display_ultra_fast,
        args=(label,),
        daemon=True
    )
    display_thread.start()
    
    print("✓ All systems running - SPEAK NOW!")
    
    # Keep overlay visible
    overlay.deiconify()
    
    if not parent:
        try:
            overlay.mainloop()
        except KeyboardInterrupt:
            stop_ultra_fast_subtitles()

def stop_ultra_fast_subtitles():
    """Stop ultra-fast subtitles."""
    global subtitle_active, overlay_instance
    
    print("🛑 Stopping ultra-fast subtitles...")
    subtitle_active = False
    
    if overlay_instance:
        try:
            overlay_instance.destroy()
        except:
            pass
        overlay_instance = None
    
    print("✓ Stopped")

def test_audio_devices():
    """Test and display available audio devices."""
    p = pyaudio.PyAudio()
    print("\n🎤 Available Audio Input Devices:")
    for i in range(p.get_device_count()):
        info = p.get_device_info_by_index(i)
        if info['maxInputChannels'] > 0:
            print(f"  Device {i}: {info['name']}")
    p.terminate()

if __name__ == "__main__":
    print("=" * 50)
    print("🚀 ULTRA-FAST REAL-TIME SUBTITLES")
    print("=" * 50)
    print("📢 ZERO DELAY - INSTANT DISPLAY")
    print("🎯 Optimized for maximum speed")
    print("💨 Words appear as you speak!")
    print("=" * 50)
    
    try:
        start_ultra_fast_subtitles()
    except KeyboardInterrupt:
        print("\n🛑 Interrupted by user")
        stop_ultra_fast_subtitles()
    except Exception as e:
        print(f"❌ Error: {e}")
        stop_ultra_fast_subtitles()