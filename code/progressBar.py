import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import threading
import time

from webScraping import *

def updateProgressBar(progressBar, percentage):
    progressBar['value'] = percentage
    progressBar.update()

# Create a function to run the long_running_function in a separate thread
def runFunctionWithProgressBar(function, *args):

    def runFunction(function, *args):
        global result
        result = function(*args)

    # Create a pop-up window to show the progress
    popup = tk.Tk()
    popup.title("Running...")
    
    # Create a progress bar widget
    progressBar = ttk.Progressbar(popup, orient=tk.HORIZONTAL, length=200, mode='determinate')
    progressBar.pack(pady=10)

    # Create a new thread for running the long_running_function
    args = args + tuple([progressBar])
    thread = threading.Thread(target=runFunction, args=(function, *args))

    # Start the thread
    thread.start()

    def check_thread():
        if thread.is_alive():
            # If the thread is still running, update the progress bar and schedule the next check
            progressBar.update()
            popup.after(100, check_thread)
        else:
            # If the thread has finished, close the pop-up window
            popup.destroy()

    # Start checking the thread's status
    popup.after(100, check_thread)

    popup.mainloop()

    return result

if __name__ == '__main__':
    players = runFunctionWithProgressBar(ws.getPlayersInfo, *[])
    print(players)


