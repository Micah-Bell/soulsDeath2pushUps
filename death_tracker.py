# Threads for watcher to work in background
import threading

class DeathTracker:
    self.game= game
    self.gui = gui

    self.death_count = 0
    self.running = False
    self.thread = None


    def start(self):
        """Begin Monitoring."""

        self.running = True

        self.thread = threading.Thread(
            target = self.monitor_screen,
            daemon = True
        )

        self.thread.start()


    def stop(self):
        """Stop Monitoring"""

        self.running = False


    # Look for deaths
    def monitor_screen(self):
        """Loops for Death Checks"""

        while self.running:
            
            # Detect death here
            
            pass


    def record_death(self):
        """Update Counters and Save Data"""