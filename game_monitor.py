import psutil
import time
from tkinter import messagebox
import tkinter as tk


class GameMonitor:

    game_list = {
        "eldenring.exe": "Elden Ring",
        "nightreign.exe": "NightReign",
        "DarkSoulsIII.exe": "Dark Souls III",
        "DarkSoulsII.exe": "Dark Souls II",
        "DarkSoulsRemastered.exe": "Dark Souls",
        "sekiro.exe": "Sekiro",
    }

    def wait_for_launch(self):
        
        while True:

            for proc in psutil.process_iter(['name']):

                try: 
                    name = proc.info["name"]
                    if name in self.game_list:
                        return (
                            self.game_list[name],
                            proc
                        )
                    
                except (
                    psutil.NoSuchProcess,
                    psutil.AccessDenied,
                    psutil.ZombieProcess
                ):
                    continue

            time.sleep(5)


    def close_game(self, proc):
        
        try:
            proc.terminate()
            proc.wait(timeout=5)

        except psutil.TimeoutExpired:
            proc.kill()


    def show_popup(self, message):
        
        root = tk.Tk()
        root.withdraw()

        messagebox.showerror(
            "X Pushups Required X",
            message
        )

        root.destroy()