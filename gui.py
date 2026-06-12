import tkinter as tk
import psutil
import time
import threading

from death_tracker import DeathTracker


class TrackerGUI:

    GAME_PROCESSES = {
        "eldenring.exe": "Elden Ring",
        "DarkSoulsIII.exe": "Dark Souls III",
        "DarkSoulsII.exe": "Dark Souls II",
        "DarkSoulsRemastered.exe": "Dark Souls",
        "sekiro.exe": "Sekiro",
        "demonssouls.exe": "Demon's Souls",
    }


    def __init__(self):
        """Create GUI Menu"""
        
        # Display Menu
        self.root = tk.Tk()
        self.root.title("Souls Death Tracker")
        self.root.geometry("700x600")

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

        # Background Watcher
        self.watcher_running = True
        self.auto_detection_enabled = True

        threading.Thread(
            target=self.watch_for_games,
            daemon=True
        ).start()

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
            "Dark Souls",
            "Dark Souls II",
            "Dark Souls III",
            "Bloodborne",
            "Sekiro",
            "Elden Ring"
        ]

        # Game Buttons
        # button_frame = tk.Frame(self.root)
        # button_frame.pack(pady=10)

        # for game in games:
        #     btn = tk.Button(
        #         button_frame,
        #         text=game,
        #         width =15,
        #         command=lambda g=game: self.select_game(g)
        #     )
        #     btn.pack(pady=2)

        #     self.game_buttons[game] = btn

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
        #      Buttons       #
        #--------------------#
        control_frame = tk.Frame(self.root)
        control_frame.pack(pady=10)

        # Start Button
        start_btn = tk.Button(
            control_frame,
            text="Start Tracking",
            command=self.start_tracking,
            width=15
        )
        start_btn.pack(side="left", padx=5)

        # Stop Button
        stop_btn = tk.Button(
            control_frame,
            text="Stop",
            command=self.stop_tracking,
            width=15
        )
        stop_btn.pack(side="left", padx=5)

        # Auto Button
        auto_btn = tk.Button(
            control_frame,
            text="Enable Auto Detect",
            command=self.enable_auto_detection,
            width=18
        )
        auto_btn.pack(side="left", padx=5)


    #--------------------#
    #   Game Selection   #
    #--------------------#
    def select_game(self, game):
        """Update the currently selected game"""

        self.selected_game = game

        # self.selected_label.config(
        #     text=f"Selected: {game}"
        # )

        # Reset button styles
        # if game in self.game_buttons:
        #     for btn in self.game_buttons.values():
        #         btn.config(relief="raised")

        #     # Highlight selected
        #     self.game_buttons[game].config(relief="sunken")


    #--------------------#
    #    Active Game     #
    #--------------------#
    def detect_running_game(self):
        """Return the Running Souls Game """

        for process in psutil.process_iter(['name']):
            try:
                name = process.info['name']

                if name in self.GAME_PROCESSES:
                    return self.GAME_PROCESSES[name]
                
            except (
                psutil.NoSuchProcess,
                psutil.AccessDenied,
                psutil.ZombieProcess
            ):
                continue

        return None
    

    #--------------------#
    #      Watcher       #
    #--------------------#
    def watch_for_games(self):
        """Watches for games in the background"""

        while self.watcher_running:
            game = self.detect_running_game()

            # Nothing Running    
            if game is None:
                self.root.after(
                    0,
                    lambda: self.selected_label.config(
                        text="Searching for Souls Game..."
                    )
                )

            # Game is running and not tracking yet
            elif (
                self.auto_detection_enabled
                and self.tracker is None
            ):
                print(f"{game} detected.")

                self.root.after(
                    0,
                    lambda g=game: self.auto_start_game(g)
                )

            # Tracking but game changed / Stop
            elif (
                self.tracker is not None
                and game != self.selected_game
            ):
                print("Game closed or changed.")

                self.root.after(
                    0,
                    self.stop_tracking
                )

            time.sleep(5)


    #--------------------#
    #     Auto Start     #
    #--------------------#
    def auto_start_game(self, game):

        print(f"{game} detected!")

        self.select_game(game)

        self.tracker = DeathTracker(
            game=game,
            gui=self
        )

        self.tracker.start()

        self.selected_label.config(
            text=f"Tracking: {game}"
        )


    #--------------------#
    #   Start Tracking   #
    #--------------------#
    def start_tracking(self):

        if self.selected_game is None:
            self.selected_label.config(
                text="Please select a game first."
            )
            return
        
        if self.tracker is not None:
            return
        
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

        # Disabled automatic restarting
        self.auto_detection_enabled = False

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
    #    Auto Detect     #
    #--------------------#
    def enable_auto_detection(self):
        self.auto_detection_enabled = True

        self.selected_label.config(
            text="Auto-detection enabled."
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

        self.watcher_running = False

        if self.tracker:
            self.tracker.stop()

        self.root.destroy()


    #--------------------#
    #      Run GUI       #
    #--------------------#
    def run(self):
        self.root.mainloop()