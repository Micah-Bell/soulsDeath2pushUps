import os
os.environ["FLAGS_use_mlkdnn"] = "0"
os.environ["KMP_DUPLICATE_LIB_OK"] = "TRUE"

from game_monitor import GameMonitor
from excel_manager import ExcelManager
from gui import TrackerGUI
from death_tracker import DeathTracker
import psutil


def Main():
    monitor = GameMonitor()
    excel = ExcelManager("souls_death_workout.xlsx")

    while True:

        # Wait for game launch
        game, process = monitor.wait_for_launch()
        print(f"{game} detected") # for debugging

        # Check if pushups are needed
        pushups = excel.get_pushups()

        if pushups > 0:

            monitor.show_popup(
                f"You owe {pushups} pushups before getting more!"
            )

            monitor.close_game(game)
            continue

        # Create only when actually needed
        gui = TrackerGUI()

        tracker = DeathTracker(
            game=game,
            gui=gui
        )

        tracker.start()

        gui.run()

        tracker.stop()

        excel.record_session(
            game,
            death_num=tracker.death_count
        )


if __name__ == "__Main__":
    Main()