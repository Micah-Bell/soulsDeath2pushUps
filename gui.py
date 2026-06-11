import tkinter as tk
from death_tracker import DeathTracker


class TrackerGUI:

    def __init__(self):
        """Create GUI Menu"""
        
        # Display Menu
        self.root = tk.Tk()
        self.root.title("Souls Death Tracker")
        self.root.geometry("700x600")

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
            font=("Adobe Garamond", 16, "bold")
        )
        title.pack(pady=10)

        #--------------------#
        #      Buttons       #
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

        button_frame = tk.Frame(self.root)
        button_frame.pack(pady=10)

        for game in games:
            btn = tk.Button(
                button_frame,
                text=game,
                width =15,
                command=lambda g=game: self.select_game(g)
            )
            btn.pack(pady=2)

            self.game_buttons[game] = btn

        #--------------------#
        #   Selected Label   #
        #--------------------#
        self.selected_label = tk.Label(
            self.root,
            text="Selected: None",
            font=("Adobe Garamond", 12)
        )
        self.selected_label.pack(pady=10)

        #--------------------#
        #   Death Counter    #
        #--------------------#
        self.death_label = tk.Label(
            self.root,
            text="Runtime Deaths: 0",
            font=("Adobe Garamond", 14)
        )
        self.death_label.pack(pady=10)

        #--------------------#
        #      Buttons       #
        #--------------------#
        control_frame = tk.Frame(self.root)
        control_frame.pack(pady=10)

        start_btn = tk.Button(
            control_frame,
            text="Start Tracking",
            command=self.start_tracking,
            width=15
        )
        start_btn.pack(side="left", padx=5)

        stop_btn = tk.Button(
            control_frame,
            text="Stop",
            command=self.stop_tracking,
            width=15
        )
        stop_btn.pack(side="left", padx=5)


    #--------------------#
    #   Game Selection   #
    #--------------------#
    def select_game(self, game):

        # Reset button styles
        for btn in self.game_buttons.values():
            btn.config(relief="raised")

        # Highlight selected
        self.game_buttons[game].config(relief="sunken")

        self.select_game = game
        self.selected_label.config(
            text=f"Selection : {game}"
        )

    #--------------------#
    #   Start Tracking   #
    #--------------------#
    def start_tracking(self):

        if self.selected_game is None:
            self.selected_label.config(text="Please select a game first.")
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
        if self.tracker:
            self.tracker.stop()

        self.selected_label.config(
            text=f"Stopped: {self.selected_game}"
        )


    #--------------------#
    #   Runtime Deaths   #
    #--------------------#
    def update_death_count(self, count):
        self.death_label.config(
            text=f"Runtime Deaths: {count}"
        )


    #--------------------#
    #   Update Death     #
    #--------------------#
    def update_death_count(self, count):
        self.death_label.config(
            text=f"Runtime Deaths: {count}"
        )


    #--------------------#
    #      Run GUI       #
    #--------------------#
    def run(self):
        self.root.mainloop()