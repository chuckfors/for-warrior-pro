import sys
import os
import tkinter as tk
import pygetwindow as gw
import pyautogui
import time
import threading
from PIL import ImageGrab, ImageEnhance, Image, ImageTk
import pytesseract
from pynput import keyboard
import mss
import json

# --- Resource Path Helper ---
def resource_path(relative_path):
    """
    Get absolute path to resource, works for development and for PyInstaller.
    """
    try:
        # PyInstaller places files in a temporary folder referenced by sys._MEIPASS
        base_path = sys._MEIPASS  
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# Set the path to the Tesseract OCR executable.
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# ---------------------------------------------------------------------------
# Functions to save and load window positions.
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

# Global dictionary for window positions.
window_positions = load_window_positions()
# ---------------------------------------------------------------------------

class DraggableBox:
    def __init__(self, root, x, y, width, height, color, label, group="", monitor=False):
        """
        A draggable and resizable box with persistent geometry.
        """
        self.group = group
        self.base_title = label  # Save the original title.
        self.top = tk.Toplevel(root)
        self.top.title(self.base_title)
        
        # Set geometry (using saved geometry if available)
        saved_geometry = window_positions.get(label)
        if saved_geometry:
            self.top.geometry(saved_geometry)
        else:
            self.top.geometry(f"{width}x{height}+{x}+{y}")
        
        self.top.resizable(True, True)
        self.top.wm_attributes('-alpha', 1.0)
        self.top.wm_attributes('-topmost', True)
        self.top.wm_attributes('-toolwindow', True)

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

        # Bind dragging events.
        self.canvas.tag_bind(self.id, '<ButtonPress-1>', self.on_press)
        self.canvas.tag_bind(self.id, '<B1-Motion>', self.on_drag)
        self.top.bind('<Configure>', self.on_configure)

        # Variables for OCR monitoring.
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
        self.canvas.config(width=event.width, height=event.height)
        self.canvas.coords(self.id, 0, 0, event.width, event.height)
        geometry = self.top.geometry()
        # Save the geometry using the base title.
        window_positions[self.base_title] = geometry
        save_window_positions(window_positions)

    def set_monitoring_title(self, is_monitoring):
        """
        Update the window title by appending a compact monitoring indicator.
        When monitoring is active, "(◉)" is added; otherwise, the title reverts to its base.
        """
        new_title = self.base_title + (" (◉)" if is_monitoring else "")
        self.top.after(0, lambda: self.top.title(new_title))

    def monitor_text(self):
        def monitor():
            self.set_monitoring_title(True)  # Append (◉) when monitoring starts.
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
            self.set_monitoring_title(False)  # Remove the indicator when monitoring stops.

        self.monitor_thread = threading.Thread(target=monitor, daemon=True)
        self.monitor_thread.start()

    def stop_monitoring(self):
        self.monitoring = False
        self.set_monitoring_title(False)
        if self.monitor_thread:
            self.monitor_thread.join(timeout=1)

# --- Functions to Open Boxes for Each Group ---
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

# --- Monitoring Controls ---
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

# --- Restore Functionality ---
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
    for title in window_positions.keys():
        if title in restore_mapping:
            if not gw.getWindowsWithTitle(title):
                restore_mapping[title]()
                print(f"Restored {title} with geometry: {window_positions[title]}")
            else:
                print(f"Window {title} is already open.")
        else:
            print(f"No restore function found for {title}.")

# --- Main Window and Scrollable Instructions Setup ---
root = tk.Tk()
root.title("ChuckFor Auto Update Level 2")

# Set custom icon using resource_path.
try:
    root.iconbitmap(resource_path("custom_icon.ico"))
except Exception as e:
    print("iconbitmap failed, using iconphoto instead. Error:", e)
    icon_image = Image.open(resource_path("custom_icon.ico"))
    icon_photo = ImageTk.PhotoImage(icon_image)
    root.iconphoto(True, icon_photo)

top_info = tk.Label(root,
    text="Send feedback or ask permission, send to chuckforaul2@gmail.com",
    wraplength=400, justify="center")
top_info.pack(pady=10)

instructions_container = tk.Frame(root)
instructions_container.pack(fill="both", expand=True, padx=10, pady=10)

canvas = tk.Canvas(instructions_container, width=500, height=400)
scrollbar = tk.Scrollbar(instructions_container, orient="vertical", command=canvas.yview)
canvas.configure(yscrollcommand=scrollbar.set)

scrollable_frame = tk.Frame(canvas)
scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
canvas.pack(side="left", fill="both", expand=True)
scrollbar.pack(side="right", fill="y")

# Bind mouse wheel events for scrolling.
def _on_mousewheel(event):
    canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
canvas.bind_all("<MouseWheel>", _on_mousewheel)
canvas.bind_all("<Button-4>", lambda event: canvas.yview_scroll(-1, "units"))
canvas.bind_all("<Button-5>", lambda event: canvas.yview_scroll(1, "units"))

admin_info_text = (
    "Some trading platforms will not work correctly unless you open this as administrator.\n\n"
    "To open as an administrator in Windows:\n"
    "1. Right-click the application or file.\n"
    "2. Select 'Run as administrator'.\n"
    "3. Confirm the action when prompted."
)
admin_info_label = tk.Label(scrollable_frame, text=admin_info_text, justify="left", wraplength=400)
admin_info_label.pack(pady=5, padx=10)

instructions_title = tk.Label(scrollable_frame, text="Instructions", font=('Arial', 14, 'bold'))
instructions_title.pack(pady=5)

def load_fullsize_image(file_path):
    image_obj = Image.open(resource_path(file_path))
    return ImageTk.PhotoImage(image_obj)

step1_text = ("1. Position the Input Box:\n"
              "Place the Input Box directly over the large stock symbol in the Stock Quote section "
              "on Day Trade Dash. Resize it so that it neatly covers just the letters and the arrow.")
step1_label = tk.Label(scrollable_frame, text=step1_text, justify="left", wraplength=400)
step1_label.pack(anchor="w", padx=10)
try:
    input_img = load_fullsize_image("input_box.png")
    input_img_label = tk.Label(scrollable_frame, image=input_img)
    input_img_label.image = input_img 
    input_img_label.pack(pady=5)
except Exception as e:
    print("Could not load input_box.png:", e)

step2_text = ("2. Set Up the Output Box:\n"
              "Move and resize the Output Box so that it overlays your Level 2 text area. "
              "Depending on your trading platform, you might have to adjust its size for optimal performance.")
step2_label = tk.Label(scrollable_frame, text=step2_text, justify="left", wraplength=400)
step2_label.pack(anchor="w", padx=10)
try:
    output_img = load_fullsize_image("output_box.png")
    output_img_label = tk.Label(scrollable_frame, image=output_img)
    output_img_label.image = output_img
    output_img_label.pack(pady=5)
except Exception as e:
    print("Could not load output_box.png:", e)

step3_text = ("3. Start Monitoring:\n"
              "Click 'Start All Monitoring' to begin capturing and forwarding text from the Input Box.")
step3_label = tk.Label(scrollable_frame, text=step3_text, justify="left", wraplength=400)
step3_label.pack(anchor="w", padx=10)
try:
    monitoring_img = load_fullsize_image("start_monitoring.png")
    monitoring_img_label = tk.Label(scrollable_frame, image=monitoring_img)
    monitoring_img_label.image = monitoring_img
    monitoring_img_label.pack(pady=5)
except Exception as e:
    print("Could not load start_monitoring.png:", e)

step4_text = ("4. Save Your Layout (Optional):\n"
              "If you would like to preserve your current window arrangement, "
              "go to the Options menu and select 'Save All Window Locations'.")
step4_label = tk.Label(scrollable_frame, text=step4_text, justify="left", wraplength=400)
step4_label.pack(anchor="w", padx=10, pady=(0, 10))
try:
    save_img = load_fullsize_image("save_layout.png")
    save_img_label = tk.Label(scrollable_frame, image=save_img)
    save_img_label.image = save_img
    save_img_label.pack(pady=5)
except Exception as e:
    print("Could not load save_layout.png:", e)

menubar = tk.Menu(root)

boxes_menu = tk.Menu(menubar, tearoff=0)
a_menu = tk.Menu(boxes_menu, tearoff=0)
a_menu.add_command(label="A Input", command=open_a_input)
a_menu.add_command(label="A Output 1", command=open_a_output1)
a_menu.add_command(label="A Output 2", command=open_a_output2)
a_menu.add_command(label="A Output 3", command=open_a_output3)
a_menu.add_command(label="A Output 4", command=open_a_output4)
a_menu.add_command(label="A Output 5", command=open_a_output5)
boxes_menu.add_cascade(label="A Boxes", menu=a_menu)

b_menu = tk.Menu(boxes_menu, tearoff=0)
b_menu.add_command(label="B Input", command=open_b_input)
b_menu.add_command(label="B Output 1", command=open_b_output1)
b_menu.add_command(label="B Output 2", command=open_b_output2)
b_menu.add_command(label="B Output 3", command=open_b_output3)
b_menu.add_command(label="B Output 4", command=open_b_output4)
b_menu.add_command(label="B Output 5", command=open_b_output5)
boxes_menu.add_cascade(label="B Boxes", menu=b_menu)

c_menu = tk.Menu(boxes_menu, tearoff=0)
c_menu.add_command(label="C Input", command=open_c_input)
c_menu.add_command(label="C Output 1", command=open_c_output1)
c_menu.add_command(label="C Output 2", command=open_c_output2)
c_menu.add_command(label="C Output 3", command=open_c_output3)
c_menu.add_command(label="C Output 4", command=open_c_output4)
c_menu.add_command(label="C Output 5", command=open_c_output5)
boxes_menu.add_cascade(label="C Boxes", menu=c_menu)

d_menu = tk.Menu(boxes_menu, tearoff=0)
d_menu.add_command(label="D Input", command=open_d_input)
d_menu.add_command(label="D Output 1", command=open_d_output1)
d_menu.add_command(label="D Output 2", command=open_d_output2)
d_menu.add_command(label="D Output 3", command=open_d_output3)
d_menu.add_command(label="D Output 4", command=open_d_output4)
d_menu.add_command(label="D Output 5", command=open_d_output5)
boxes_menu.add_cascade(label="D Boxes", menu=d_menu)

menubar.add_cascade(label="Boxes", menu=boxes_menu)

monitor_menu = tk.Menu(menubar, tearoff=0)
monitor_menu.add_command(label="Start All Monitoring", command=start_all_monitoring)
monitor_menu.add_command(label="Stop All Monitoring", command=stop_all_monitoring)
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

options_menu = tk.Menu(menubar, tearoff=0)
options_menu.add_command(label="Save Window Locations", command=lambda: (save_window_positions(window_positions), print("All window locations have been saved.")))
options_menu.add_command(label="Restore Saved Windows", command=restore_saved_windows)
options_menu.add_command(label="Delete Saved Window Locations", command=delete_saved_window_locations)
menubar.add_cascade(label="Options", menu=options_menu)

root.config(menu=menubar)

def type_text(text, group):
    outputs = [f"{group} Output {i}" for i in range(1, 6)]
    sent = False
    for out in outputs:
        windows = gw.getWindowsWithTitle(out)
        if windows:
            sent = True
            box = windows[0]
            box.activate()
            time.sleep(0.3)
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

listener = keyboard.Listener(on_press=on_press)
listener.start()

root.mainloop()
