from game_monitor import GameMonitor
from excel_manager import ExcelManager
from gui import TrackerGUI
from death_tracker import DeathTracker
import psutil
import time
import threading


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

            monitor.close_game(process)
            continue

        # Create only when actually needed
        gui = TrackerGUI()
        gui.current_game = game

        tracker = DeathTracker(
            game=game,
            gui=gui
        )
        gui.tracker = tracker

        tracker.start()

        gui.run()

        tracker.stop()
        time.sleep(0.2)

        final_deaths = tracker.death_count

        excel.record_session(
            game,
            final_deaths
        )


if __name__ == "__main__":
    Main()