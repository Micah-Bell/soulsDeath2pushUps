import tkinter as tk # gui
import psutil # process and system utilites
import time
import threading


class TrackerGUI:

    def __init__(self):
        """Create GUI Menu"""

        # State Data
        self.session_deaths = 0
        self.current_game = None

        # Colors
        self.bg_color = "#27170d"
        self.fg_color = "#bd6707"
        
        # Display Menu
        self.root = tk.Tk()
        self.root.title("Souls Death Tracker")
        self.root.geometry("600x200")
        self.root.configure(bg=self.bg_color)

        # Clean Up
        self.root.protocol(
            "WM_DELETE_WINDOW",
            self.on_close
        )

        #--------------------#
        #    Title Label     #
        #--------------------#
        title = tk.Label(
            self.root,
            text="-~-~-~-  Souls Death Tracker -~-~--~-",
            font=("Adobe Garamond", 24, "bold"),
            bg=self.bg_color,
            fg=self.fg_color
        )
        title.pack(pady=10)

        #--------------------#
        #    Status Label    #
        #--------------------#
        self.selected_label = tk.Label(
            self.root,
            text="Searching for Souls Game...",
            font=("Adobe Garamond", 16),
            bg=self.bg_color,
            fg=self.fg_color
        )
        self.selected_label.pack(pady=10)

        #--------------------#
        #   Death # Label    #
        #--------------------#
        self.death_label = tk.Label(
            self.root,
            text="Runtime Deaths: 0",
            font=("Adobe Garamond", 18),
            bg=self.bg_color,
            fg=self.fg_color
        )
        self.death_label.pack(pady=10)

        # Tracker Ref
        self.tracker = None

    #--------------------#
    #   Start Tracking   #
    #--------------------#
    def start_tracking(self):
        from death_tracker import DeathTracker

        self.tracker = DeathTracker(
            game=self.current_game,
            gui=self
        )

        self.tracker.start()
        self.update_ui()


    #--------------------#
    #   Stop Tracking    #
    #--------------------#
    def stop_tracking(self):

        if self.tracker:
            self.tracker.stop()
            self.tracker = None

            self.current_game = None
            self.session_deaths = 0
            self.update_ui()


    #--------------------#
    #     Update GUI     #
    #--------------------#
    def update_ui(self):

        if self.current_game:
            status = f"Tracking: {self.current_game}"
        else:
            status = "Searching for Souls Game..."

        self.selected_label.config(text=status)

        self.death_label.config(
            text=f"Session Deaths: {self.session_deaths}"
        )


    #--------------------#
    #   Runtime Deaths   #
    #--------------------#
    def update_death_count(self, count):

        self.session_deaths = count
        self.root.after(0, self.update_ui)


    #--------------------#
    #   Close Clean Up   #
    #--------------------#
    def on_close(self):

        if self.tracker:
            self.tracker.stop()

        self.root.destroy()


    #--------------------#
    #      Run GUI       #
    #--------------------#
    def run(self):
        self.root.mainloop()