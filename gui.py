import tkinter as tk # gui
import psutil # process and system utilites
import time
import threading

from death_tracker import DeathTracker


class TrackerGUI:

    def __init__(self):
        """Create GUI Menu"""
        
        # Display Menu
        self.root = tk.Tk()
        self.root.title("Souls Death Tracker")
        self.root.geometry("600x200")

        # Clean Up
        self.root.protocol(
            "WM_DELETE_WINDOW",
            self.on_close
        )

        # Tracker ref
        self.tracker = None

        # State
        self.selected_game = None
        self.game_buttons = {}

        #--------------------#
        #       Title        #
        #--------------------#
        title = tk.Label(
            self.root,
            text="-~-~-~-  Souls Death Tracker -~-~--~-",
            font=("Adobe Garamond", 24, "bold")
        )
        title.pack(pady=10)

        #--------------------#
        #       Games        #
        #--------------------#
        games = [
            "Demon's Souls",
            "Nightreign",
            "Dark Souls",
            "Dark Souls II",
            "Dark Souls III",
            "Bloodborne",
            "Sekiro",
            "Elden Ring"
        ]

        #--------------------#
        #   Selected Label   #
        #--------------------#
        self.selected_label = tk.Label(
            self.root,
            text="Searching for Souls Game...",
            font=("Adobe Garamond", 16)
        )
        self.selected_label.pack(pady=10)

        #--------------------#
        #   Death # Label    #
        #--------------------#
        self.death_label = tk.Label(
            self.root,
            text="Runtime Deaths: 0",
            font=("Adobe Garamond", 18)
        )
        self.death_label.pack(pady=10)


    #--------------------#
    #   Start Tracking   #
    #--------------------#
    def start_tracking(self):
        
        self.tracker = DeathTracker(
            game=self.selected_game,
            gui=self
        )

        self.tracker.start()

        self.selected_label.config(
            text=f"Tracking: {self.selected_game}"
        )


    #--------------------#
    #   Stop Tracking    #
    #--------------------#
    def stop_tracking(self):

        if self.tracker:
            self.tracker.stop()
            self.tracker = None

            # Raise all game buttons back up
            # for btn in self.game_buttons.values():
            #     btn.config(relief="raised")

            # Clear the selected game
            self.selected_game = None

        self.selected_label.config(
            text="Searching for Souls Game..."
        )


    #--------------------#
    #   Runtime Deaths   #
    #--------------------#
    def update_death_count(self, count):
        self.death_label.config(
            text=f"Runtime Deaths: {count}"
        )

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