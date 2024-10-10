import tkinter as tk
import vlc
import os

# Create a simple class to handle the Picture-in-Picture window
class PiPVideoPlayer:
    def __init__(self, video_path):
        self.root = tk.Tk()
        self.root.title("PiP Video Player")
        
        # Set the window to be always on top and resizable
        self.root.attributes("-topmost", True)
        self.root.resizable(True, True)

        # Set window to be frameless for a clean PiP look (optional)
        self.root.overrideredirect(True)

        # Create a canvas to host the video
        self.canvas = tk.Canvas(self.root)
        self.canvas.pack(fill=tk.BOTH, expand=True)

        # Create an instance of VLC player
        self.instance = vlc.Instance()
        self.player = self.instance.media_player_new()

        # Set the media (video) file
        self.media = self.instance.media_new(video_path)
        self.player.set_media(self.media)

        # Set the video output to the canvas window handle
        self.player.set_hwnd(self.canvas.winfo_id())

        # Bind mouse events to move or resize the window
        self.canvas.bind("<ButtonPress-1>", self.start_move)
        self.canvas.bind("<B1-Motion>", self.on_move)

        # Bind the Escape key to close the window
        self.root.bind("<Escape>", self.close)

        # Play the video
        self.player.play()

        # Start the Tkinter event loop
        self.root.mainloop()

    def start_move(self, event):
        """Save mouse position when dragging starts."""
        self.start_x = event.x
        self.start_y = event.y

    def on_move(self, event):
        """Move the window by dragging."""
        x = self.root.winfo_x() + event.x - self.start_x
        y = self.root.winfo_y() + event.y - self.start_y
        self.root.geometry(f"+{x}+{y}")

    def close(self, event=None):
        """Close the PiP window when Escape is pressed or right-clicked."""
        self.player.stop()
        self.root.quit()

# Check if the video file exists
video_path = "qj0sopzku6kc1.mp4"
if os.path.exists(video_path):
    player = PiPVideoPlayer(video_path)
else:
    print(f"Video file '{video_path}' not found.")
