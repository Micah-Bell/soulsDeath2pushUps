# Threads for watcher to work in background
import threading
import time
from death_detector import DeathDetector
from excel_manager import ExcelManager


class DeathTracker:

    def __init__(self, game, gui):
        self.game = game
        self.gui = gui

        self.detector = DeathDetector()
        self.excel = ExcelManager("souls_death_workout.xlsx")

        self.death_count = 0
        self.running = False
        self.thread = None


    #-------------------------#
    #  Start Tracking Thread  #
    #-------------------------#
    def start(self):
        """Begin Monitoring."""

        if self.running:
            return

        self.running = True

        self.thread = threading.Thread(
            target=self.monitor_screen,
            daemon=True
        )

        self.thread.start()


    #-------------------------#
    #      Stop Tracking      #
    #-------------------------#
    def stop(self):
        """Stop Monitoring"""

        self.running = False


    #-------------------------#
    #  Main Loop (bkgnd thrd) #
    #-------------------------#
    def monitor_screen(self):
        """Loops for Death Checks"""

        print("tracker thread started") # DEBUG

        while self.running:
            if self.detector.detect_death():

                self.death_count += 1

                # Update GUI Safely
                self.gui.root.after(
                    0,
                    self.gui.update_death_count(self.death_count)
                )

            time.sleep(3)