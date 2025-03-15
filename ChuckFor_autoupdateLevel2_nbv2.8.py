import tkinter as tk
import pygetwindow as gw
import pyautogui
import time
import threading
from PIL import ImageGrab, ImageEnhance, Image
import pytesseract
from pynput import keyboard
import mss
import json
import os

# Set the path to the Tesseract OCR executable.
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# ---------------------------------------------------------------------------
# Configuration file functions to save and load window positions.
CONFIG_FILE = "window_positions.json"

def load_window_positions():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r") as f:
            return json.load(f)
    return {}

def save_window_positions(positions):
    with open(CONFIG_FILE, "w") as f:
        json.dump(positions, f, indent=4)

def delete_saved_window_locations():
    global window_positions
    if os.path.exists(CONFIG_FILE):
        os.remove(CONFIG_FILE)
        window_positions = {}
        print("Saved window locations have been deleted.")

# Global dictionary that holds positions keyed by window title.
window_positions = load_window_positions()
# ---------------------------------------------------------------------------

class DraggableBox:
    def __init__(self, root, x, y, width, height, color, label, group="", monitor=False):
        """
        A draggable and resizable box with persistent geometry.
        If color is "clear", uses a transparent background.
        The 'group' parameter helps route text between inputs and outputs.
        """
        self.group = group
        self.top = tk.Toplevel(root)
        self.top.title(label)
        
        # Load saved geometry if available; otherwise use the defaults.
        saved_geometry = window_positions.get(label)
        if saved_geometry:
            self.top.geometry(saved_geometry)
        else:
            self.top.geometry(f"{width}x{height}+{x}+{y}")
        
        self.top.resizable(True, True)
        self.top.wm_attributes('-alpha', 1.0)  # fully opaque
        self.top.wm_attributes('-topmost', True)  # always on top
        self.top.wm_attributes('-toolwindow', True)  # removes minimize/maximize buttons

        if color.lower() == "clear":
            transparent_color = "hotpink"
            self.top.configure(bg=transparent_color)
            self.top.wm_attributes("-transparentcolor", transparent_color)
            self.canvas = tk.Canvas(self.top, width=width, height=height,
                                    bg=transparent_color, highlightthickness=0)
            self.canvas.pack(fill=tk.BOTH, expand=True)
            self.id = self.canvas.create_rectangle(0, 0, width, height, outline="black", fill="")
        else:
            self.top.configure(bg="white")
            self.canvas = tk.Canvas(self.top, width=width, height=height, bg="white")
            self.canvas.pack(fill=tk.BOTH, expand=True)
            self.id = self.canvas.create_rectangle(0, 0, width, height, fill=color)

        # Bind events for dragging.
        self.canvas.tag_bind(self.id, '<ButtonPress-1>', self.on_press)
        self.canvas.tag_bind(self.id, '<B1-Motion>', self.on_drag)
        self.top.bind('<Configure>', self.on_configure)

        # Variables used for OCR monitoring.
        self.last_text = ""
        self.monitoring = False
        self.monitor_thread = None

    def on_press(self, event):
        self.start_x = event.x
        self.start_y = event.y

    def on_drag(self, event):
        dx = event.x - self.start_x
        dy = event.y - self.start_y
        x = self.top.winfo_x() + dx
        y = self.top.winfo_y() + dy
        self.top.geometry(f"+{x}+{y}")

    def on_configure(self, event):
        # Update the canvas dimensions and rectangle coordinates.
        self.canvas.config(width=event.width, height=event.height)
        self.canvas.coords(self.id, 0, 0, event.width, event.height)
        
        # Save current geometry (format: "widthxheight+x+y").
        geometry = self.top.geometry()
        window_positions[self.top.title()] = geometry
        save_window_positions(window_positions)

    def monitor_text(self):
        """
        Capture the box’s region every second, perform OCR (with preprocessing),
        filter the result and if the text changes then send it to the corresponding output windows.
        """
        def monitor():
            print(f"Monitoring started for {self.top.title()} (Group: {self.group}).")
            self.monitoring = True
            with mss.mss() as sct:
                while self.monitoring:
                    time.sleep(1)
                    if not self.top.winfo_exists():
                        break
                    x1 = self.top.winfo_rootx()
                    y1 = self.top.winfo_rooty()
                    x2 = x1 + self.top.winfo_width()
                    y2 = y1 + self.top.winfo_height()
                    monitor_region = {"top": y1, "left": x1, "width": x2 - x1, "height": y2 - y1}
                    print(f"Capturing region: {monitor_region}")
                    img = sct.grab(monitor_region)
                    img = Image.frombytes("RGB", img.size, img.bgra, "raw", "BGRX")
                    img = img.convert('L')
                    enhancer = ImageEnhance.Contrast(img)
                    img = enhancer.enhance(2.0)
                    ocr_result = pytesseract.image_to_string(img).strip()
                    print(f"OCR captured ({self.top.title()}):", ocr_result)
                    filtered_text = ''.join(filter(str.isalpha, ocr_result)).lower()
                    print(f"Filtered OCR ({self.top.title()}):", filtered_text)
                    if filtered_text and filtered_text != self.last_text:
                        self.last_text = filtered_text
                        type_text(filtered_text, self.group)
        self.monitor_thread = threading.Thread(target=monitor, daemon=True)
        self.monitor_thread.start()

    def stop_monitoring(self):
        self.monitoring = False
        if self.monitor_thread:
            self.monitor_thread.join(timeout=1)

def type_text(text, group):
    """
    For the given letter group (A, B, etc.), send the text sequentially to the
    group's output windows (Output 1 to Output 5) with a 0.5 second delay in between.
    """
    outputs = [f"{group} Output {i}" for i in range(1, 6)]
    sent = False
    for out in outputs:
        windows = gw.getWindowsWithTitle(out)
        if windows:
            sent = True
            box = windows[0]
            box.activate()
            time.sleep(0.3)  # Allow the window to fully activate
            # Click at the center of the box.
            click_x = box.left + box.width // 2
            click_y = box.top + box.height // 2
            print(f"Clicking at: ({click_x}, {click_y}) for {out}")
            pyautogui.click(x=click_x, y=click_y)
            time.sleep(0.1)
            pyautogui.write(text, interval=0.05)
            pyautogui.press('enter')
            print(f"Sent text to {out}")
            time.sleep(0.5)
    if not sent:
        print(f"No {group} output windows open; text not sent.")

def on_press(key):
    """
    When a key is pressed, if the active window has a title beginning with one of our groups
    (A, B, C, D) then send the keystroke to that group’s "Output 1" window.
    """
    try:
        active_window = gw.getActiveWindow()
        if active_window:
            title_parts = active_window.title.split(" ")
            if len(title_parts) >= 2 and title_parts[0] in {"A", "B", "C", "D"}:
                group = title_parts[0]
                if hasattr(key, 'char') and key.char:
                    windows = gw.getWindowsWithTitle(f"{group} Output 1")
                    if windows:
                        windows[0].activate()
                        pyautogui.write(key.char)
    except Exception:
        pass

# === Functions for Opening Input and Output Windows for Each Group ===
# A Group
def open_a_input():
    global a_input
    a_input = DraggableBox(root, 50, 50, 100, 100, 'clear', 'A Input', group="A")

def open_a_output1():
    global a_output1
    a_output1 = DraggableBox(root, 250, 50, 100, 100, 'clear', 'A Output 1', group="A")

def open_a_output2():
    global a_output2
    a_output2 = DraggableBox(root, 450, 50, 100, 100, 'clear', 'A Output 2', group="A")

def open_a_output3():
    global a_output3
    a_output3 = DraggableBox(root, 650, 50, 100, 100, 'clear', 'A Output 3', group="A")

def open_a_output4():
    global a_output4
    a_output4 = DraggableBox(root, 850, 50, 100, 100, 'clear', 'A Output 4', group="A")

def open_a_output5():
    global a_output5
    a_output5 = DraggableBox(root, 1050, 50, 100, 100, 'clear', 'A Output 5', group="A")

# B Group
def open_b_input():
    global b_input
    b_input = DraggableBox(root, 50, 200, 100, 100, 'clear', 'B Input', group="B")

def open_b_output1():
    global b_output1
    b_output1 = DraggableBox(root, 250, 200, 100, 100, 'clear', 'B Output 1', group="B")

def open_b_output2():
    global b_output2
    b_output2 = DraggableBox(root, 450, 200, 100, 100, 'clear', 'B Output 2', group="B")

def open_b_output3():
    global b_output3
    b_output3 = DraggableBox(root, 650, 200, 100, 100, 'clear', 'B Output 3', group="B")

def open_b_output4():
    global b_output4
    b_output4 = DraggableBox(root, 850, 200, 100, 100, 'clear', 'B Output 4', group="B")

def open_b_output5():
    global b_output5
    b_output5 = DraggableBox(root, 1050, 200, 100, 100, 'clear', 'B Output 5', group="B")

# C Group
def open_c_input():
    global c_input
    c_input = DraggableBox(root, 50, 350, 100, 100, 'clear', 'C Input', group="C")

def open_c_output1():
    global c_output1
    c_output1 = DraggableBox(root, 250, 350, 100, 100, 'clear', 'C Output 1', group="C")

def open_c_output2():
    global c_output2
    c_output2 = DraggableBox(root, 450, 350, 100, 100, 'clear', 'C Output 2', group="C")

def open_c_output3():
    global c_output3
    c_output3 = DraggableBox(root, 650, 350, 100, 100, 'clear', 'C Output 3', group="C")

def open_c_output4():
    global c_output4
    c_output4 = DraggableBox(root, 850, 350, 100, 100, 'clear', 'C Output 4', group="C")

def open_c_output5():
    global c_output5
    c_output5 = DraggableBox(root, 1050, 350, 100, 100, 'clear', 'C Output 5', group="C")

# D Group
def open_d_input():
    global d_input
    d_input = DraggableBox(root, 50, 500, 100, 100, 'clear', 'D Input', group="D")

def open_d_output1():
    global d_output1
    d_output1 = DraggableBox(root, 250, 500, 100, 100, 'clear', 'D Output 1', group="D")

def open_d_output2():
    global d_output2
    d_output2 = DraggableBox(root, 450, 500, 100, 100, 'clear', 'D Output 2', group="D")

def open_d_output3():
    global d_output3
    d_output3 = DraggableBox(root, 650, 500, 100, 100, 'clear', 'D Output 3', group="D")

def open_d_output4():
    global d_output4
    d_output4 = DraggableBox(root, 850, 500, 100, 100, 'clear', 'D Output 4', group="D")

def open_d_output5():
    global d_output5
    d_output5 = DraggableBox(root, 1050, 500, 100, 100, 'clear', 'D Output 5', group="D")

# === Monitoring Controls ===
def start_all_monitoring():
    if 'a_input' in globals() and a_input:
        a_input.monitor_text()
    if 'b_input' in globals() and b_input:
        b_input.monitor_text()
    if 'c_input' in globals() and c_input:
        c_input.monitor_text()
    if 'd_input' in globals() and d_input:
        d_input.monitor_text()

def stop_all_monitoring():
    if 'a_input' in globals() and a_input:
        a_input.stop_monitoring()
    if 'b_input' in globals() and b_input:
        b_input.stop_monitoring()
    if 'c_input' in globals() and c_input:
        c_input.stop_monitoring()
    if 'd_input' in globals() and d_input:
        d_input.stop_monitoring()

def start_a_monitoring():
    if 'a_input' in globals() and a_input:
        a_input.monitor_text()

def stop_a_monitoring():
    if 'a_input' in globals() and a_input:
        a_input.stop_monitoring()

def start_b_monitoring():
    if 'b_input' in globals() and b_input:
        b_input.monitor_text()

def stop_b_monitoring():
    if 'b_input' in globals() and b_input:
        b_input.stop_monitoring()

def start_c_monitoring():
    if 'c_input' in globals() and c_input:
        c_input.monitor_text()

def stop_c_monitoring():
    if 'c_input' in globals() and c_input:
        c_input.stop_monitoring()

def start_d_monitoring():
    if 'd_input' in globals() and d_input:
        d_input.monitor_text()

def stop_d_monitoring():
    if 'd_input' in globals() and d_input:
        d_input.stop_monitoring()

# === Restore Functionality ===
restore_mapping = {
    "A Input": open_a_input,
    "A Output 1": open_a_output1,
    "A Output 2": open_a_output2,
    "A Output 3": open_a_output3,
    "A Output 4": open_a_output4,
    "A Output 5": open_a_output5,
    "B Input": open_b_input,
    "B Output 1": open_b_output1,
    "B Output 2": open_b_output2,
    "B Output 3": open_b_output3,
    "B Output 4": open_b_output4,
    "B Output 5": open_b_output5,
    "C Input": open_c_input,
    "C Output 1": open_c_output1,
    "C Output 2": open_c_output2,
    "C Output 3": open_c_output3,
    "C Output 4": open_c_output4,
    "C Output 5": open_c_output5,
    "D Input": open_d_input,
    "D Output 1": open_d_output1,
    "D Output 2": open_d_output2,
    "D Output 3": open_d_output3,
    "D Output 4": open_d_output4,
    "D Output 5": open_d_output5
}

def restore_saved_windows():
    """
    Iterate through the saved window positions and call the appropriate function to re-create them,
    if they are not already open.
    """
    for title in window_positions.keys():
        if title in restore_mapping:
            if not gw.getWindowsWithTitle(title):
                restore_mapping[title]()
                print(f"Restored {title} with geometry: {window_positions[title]}")
            else:
                print(f"Window {title} is already open.")
        else:
            print(f"No restore function found for {title}.")

# === Main Window and Menu Setup ===
root = tk.Tk()
root.title("Main Window")

info_label = tk.Label(root, text="Some trading platforms will not work correctly unless you open this as administrator. To open something as an administrator in Windows, right-click the application or file, select 'Run as administrator', and then confirm the action when prompted.", wraplength=400, justify="left")
info_label.pack(pady=10)

menubar = tk.Menu(root)

# "Boxes" drop-down menu with sub-menus for each group.
boxes_menu = tk.Menu(menubar, tearoff=0)
# A Group submenu.
a_menu = tk.Menu(boxes_menu, tearoff=0)
a_menu.add_command(label="A Input", command=open_a_input)
a_menu.add_command(label="A Output 1", command=open_a_output1)
a_menu.add_command(label="A Output 2", command=open_a_output2)
a_menu.add_command(label="A Output 3", command=open_a_output3)
a_menu.add_command(label="A Output 4", command=open_a_output4)
a_menu.add_command(label="A Output 5", command=open_a_output5)
boxes_menu.add_cascade(label="A Boxes", menu=a_menu)
# B Group submenu.
b_menu = tk.Menu(boxes_menu, tearoff=0)
b_menu.add_command(label="B Input", command=open_b_input)
b_menu.add_command(label="B Output 1", command=open_b_output1)
b_menu.add_command(label="B Output 2", command=open_b_output2)
b_menu.add_command(label="B Output 3", command=open_b_output3)
b_menu.add_command(label="B Output 4", command=open_b_output4)
b_menu.add_command(label="B Output 5", command=open_b_output5)
boxes_menu.add_cascade(label="B Boxes", menu=b_menu)
# C Group submenu.
c_menu = tk.Menu(boxes_menu, tearoff=0)
c_menu.add_command(label="C Input", command=open_c_input)
c_menu.add_command(label="C Output 1", command=open_c_output1)
c_menu.add_command(label="C Output 2", command=open_c_output2)
c_menu.add_command(label="C Output 3", command=open_c_output3)
c_menu.add_command(label="C Output 4", command=open_c_output4)
c_menu.add_command(label="C Output 5", command=open_c_output5)
boxes_menu.add_cascade(label="C Boxes", menu=c_menu)
# D Group submenu.
d_menu = tk.Menu(boxes_menu, tearoff=0)
d_menu.add_command(label="D Input", command=open_d_input)
d_menu.add_command(label="D Output 1", command=open_d_output1)
d_menu.add_command(label="D Output 2", command=open_d_output2)
d_menu.add_command(label="D Output 3", command=open_d_output3)
d_menu.add_command(label="D Output 4", command=open_d_output4)
d_menu.add_command(label="D Output 5", command=open_d_output5)
boxes_menu.add_cascade(label="D Boxes", menu=d_menu)

menubar.add_cascade(label="Boxes", menu=boxes_menu)

# "Monitoring" drop-down menu.
monitor_menu = tk.Menu(menubar, tearoff=0)
monitor_menu.add_command(label="Start All Monitoring", command=start_all_monitoring)
monitor_menu.add_command(label="Stop All Monitoring", command=stop_all_monitoring)
# Group-specific monitoring submenu.
group_monitor_menu = tk.Menu(monitor_menu, tearoff=0)
group_monitor_menu.add_command(label="Start A Monitoring", command=start_a_monitoring)
group_monitor_menu.add_command(label="Stop A Monitoring", command=stop_a_monitoring)
group_monitor_menu.add_command(label="Start B Monitoring", command=start_b_monitoring)
group_monitor_menu.add_command(label="Stop B Monitoring", command=stop_b_monitoring)
group_monitor_menu.add_command(label="Start C Monitoring", command=start_c_monitoring)
group_monitor_menu.add_command(label="Stop C Monitoring", command=stop_c_monitoring)
group_monitor_menu.add_command(label="Start D Monitoring", command=start_d_monitoring)
group_monitor_menu.add_command(label="Stop D Monitoring", command=stop_d_monitoring)
monitor_menu.add_cascade(label="Group Monitoring", menu=group_monitor_menu)

menubar.add_cascade(label="Monitoring", menu=monitor_menu)

# "Options" drop-down menu with Save and Restore commands.
def save_all_window_locations():
    save_window_positions(window_positions)
    print("All window locations have been saved.")

options_menu = tk.Menu(menubar, tearoff=0)
options_menu.add_command(label="Save Window Locations", command=save_all_window_locations)
options_menu.add_command(label="Restore Saved Windows", command=restore_saved_windows)
options_menu.add_command(label="Delete Saved Window Locations", command=delete_saved_window_locations)
menubar.add_cascade(label="Options", menu=options_menu)

root.config(menu=menubar)

# Start the keyboard listener.
listener = keyboard.Listener(on_press=on_press)
listener.start()

root.mainloop()
