import tkinter as tk
import pygetwindow as gw
import pyautogui
import time
import threading
from PIL import ImageGrab, ImageEnhance
import pytesseract
from pynput import keyboard
import mss
from PIL import Image

# Set the path to the Tesseract OCR executable
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

class DraggableBox:
    def __init__(self, root, x, y, width, height, color, label, group="", monitor=False):
        """
        A draggable and resizable box.
        If color is "clear", uses a transparent background.
        The 'group' parameter helps route text between inputs and outputs.
        """
        self.group = group
        self.top = tk.Toplevel(root)
        self.top.geometry(f"{width}x{height}+{x}+{y}")
        self.top.title(label)
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

        self.canvas.tag_bind(self.id, '<ButtonPress-1>', self.on_press)
        self.canvas.tag_bind(self.id, '<B1-Motion>', self.on_drag)
        self.top.bind('<Configure>', self.on_resize)

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

    def on_resize(self, event):
        self.canvas.config(width=event.width, height=event.height)
        self.canvas.coords(self.id, 0, 0, event.width, event.height)

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
                    monitor = {"top": y1, "left": x1, "width": x2 - x1, "height": y2 - y1}
                    print(f"Capturing region: {monitor}")  # Debugging information
                    img = sct.grab(monitor)
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
    For the given letter group (AA, BB, etc.), send the text sequentially to the
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
            click_x = box.left + 10
            click_y = box.top + box.height + 10
            pyautogui.click(x=click_x, y=click_y)
            pyautogui.write(text)
            pyautogui.press('enter')
            print(f"Sent text to {out}")
            time.sleep(0.5)
    if not sent:
        print(f"No {group} output windows open; text not sent.")

def on_press(key):
    """
    When a key is pressed, if the active window has a title beginning with one of our groups
    (AA, BB, CC, DD) then send the keystroke to that group’s "Output 1" window.
    """
    try:
        active_window = gw.getActiveWindow()
        if active_window:
            title_parts = active_window.title.split(" ")
            if len(title_parts) >= 2 and title_parts[0] in {"AA", "BB", "CC", "DD"}:
                group = title_parts[0]
                if hasattr(key, 'char') and key.char:
                    windows = gw.getWindowsWithTitle(f"{group} Output 1")
                    if windows:
                        windows[0].activate()
                        pyautogui.write(key.char)
    except Exception:
        pass

# === Functions for Opening Input and Output Windows for Each Group ===
# AA Group
def open_aa_input():
    global aa_input
    aa_input = DraggableBox(root, 50, 50, 100, 100, 'clear', 'AA Input', group="AA")

def open_aa_output1():
    global aa_output1
    aa_output1 = DraggableBox(root, 250, 50, 100, 100, 'clear', 'AA Output 1', group="AA")

def open_aa_output2():
    global aa_output2
    aa_output2 = DraggableBox(root, 450, 50, 100, 100, 'clear', 'AA Output 2', group="AA")

def open_aa_output3():
    global aa_output3
    aa_output3 = DraggableBox(root, 650, 50, 100, 100, 'clear', 'AA Output 3', group="AA")

def open_aa_output4():
    global aa_output4
    aa_output4 = DraggableBox(root, 850, 50, 100, 100, 'clear', 'AA Output 4', group="AA")

def open_aa_output5():
    global aa_output5
    aa_output5 = DraggableBox(root, 1050, 50, 100, 100, 'clear', 'AA Output 5', group="AA")

# BB Group
def open_bb_input():
    global bb_input
    bb_input = DraggableBox(root, 50, 200, 100, 100, 'clear', 'BB Input', group="BB")

def open_bb_output1():
    global bb_output1
    bb_output1 = DraggableBox(root, 250, 200, 100, 100, 'clear', 'BB Output 1', group="BB")

def open_bb_output2():
    global bb_output2
    bb_output2 = DraggableBox(root, 450, 200, 100, 100, 'clear', 'BB Output 2', group="BB")

def open_bb_output3():
    global bb_output3
    bb_output3 = DraggableBox(root, 650, 200, 100, 100, 'clear', 'BB Output 3', group="BB")

def open_bb_output4():
    global bb_output4
    bb_output4 = DraggableBox(root, 850, 200, 100, 100, 'clear', 'BB Output 4', group="BB")

def open_bb_output5():
    global bb_output5
    bb_output5 = DraggableBox(root, 1050, 200, 100, 100, 'clear', 'BB Output 5', group="BB")

# CC Group
def open_cc_input():
    global cc_input
    cc_input = DraggableBox(root, 50, 350, 100, 100, 'clear', 'CC Input', group="CC")

def open_cc_output1():
    global cc_output1
    cc_output1 = DraggableBox(root, 250, 350, 100, 100, 'clear', 'CC Output 1', group="CC")

def open_cc_output2():
    global cc_output2
    cc_output2 = DraggableBox(root, 450, 350, 100, 100, 'clear', 'CC Output 2', group="CC")

def open_cc_output3():
    global cc_output3
    cc_output3 = DraggableBox(root, 650, 350, 100, 100, 'clear', 'CC Output 3', group="CC")

def open_cc_output4():
    global cc_output4
    cc_output4 = DraggableBox(root, 850, 350, 100, 100, 'clear', 'CC Output 4', group="CC")

def open_cc_output5():
    global cc_output5
    cc_output5 = DraggableBox(root, 1050, 350, 100, 100, 'clear', 'CC Output 5', group="CC")

# DD Group
def open_dd_input():
    global dd_input
    dd_input = DraggableBox(root, 50, 500, 100, 100, 'clear', 'DD Input', group="DD")

def open_dd_output1():
    global dd_output1
    dd_output1 = DraggableBox(root, 250, 500, 100, 100, 'clear', 'DD Output 1', group="DD")

def open_dd_output2():
    global dd_output2
    dd_output2 = DraggableBox(root, 450, 500, 100, 100, 'clear', 'DD Output 2', group="DD")

def open_dd_output3():
    global dd_output3
    dd_output3 = DraggableBox(root, 650, 500, 100, 100, 'clear', 'DD Output 3', group="DD")

def open_dd_output4():
    global dd_output4
    dd_output4 = DraggableBox(root, 850, 500, 100, 100, 'clear', 'DD Output 4', group="DD")

def open_dd_output5():
    global dd_output5
    dd_output5 = DraggableBox(root, 1050, 500, 100, 100, 'clear', 'DD Output 5', group="DD")

# === Monitoring Controls ===
def start_all_monitoring():
    if 'aa_input' in globals() and aa_input:
        aa_input.monitor_text()
    if 'bb_input' in globals() and bb_input:
        bb_input.monitor_text()
    if 'cc_input' in globals() and cc_input:
        cc_input.monitor_text()
    if 'dd_input' in globals() and dd_input:
        dd_input.monitor_text()

def stop_all_monitoring():
    if 'aa_input' in globals() and aa_input:
        aa_input.stop_monitoring()
    if 'bb_input' in globals() and bb_input:
        bb_input.stop_monitoring()
    if 'cc_input' in globals() and cc_input:
        cc_input.stop_monitoring()
    if 'dd_input' in globals() and dd_input:
        dd_input.stop_monitoring()

# Group-specific monitoring functions.
def start_aa_monitoring():
    if 'aa_input' in globals() and aa_input:
        aa_input.monitor_text()
def stop_aa_monitoring():
    if 'aa_input' in globals() and aa_input:
        aa_input.stop_monitoring()

def start_bb_monitoring():
    if 'bb_input' in globals() and bb_input:
        bb_input.monitor_text()
def stop_bb_monitoring():
    if 'bb_input' in globals() and bb_input:
        bb_input.stop_monitoring()

def start_cc_monitoring():
    if 'cc_input' in globals() and cc_input:
        cc_input.monitor_text()
def stop_cc_monitoring():
    if 'cc_input' in globals() and cc_input:
        cc_input.stop_monitoring()

def start_dd_monitoring():
    if 'dd_input' in globals() and dd_input:
        dd_input.monitor_text()
def stop_dd_monitoring():
    if 'dd_input' in globals() and dd_input:
        dd_input.stop_monitoring()

# === Main Window with Menubar Drop-downs ===
root = tk.Tk()
root.title("Main Window")

menubar = tk.Menu(root)

# "Open Boxes" drop-down menu with cascaded sub-menus for each group.
open_boxes_menu = tk.Menu(menubar, tearoff=0)
# AA Group submenu.
aa_menu = tk.Menu(open_boxes_menu, tearoff=0)
aa_menu.add_command(label="Open AA Input", command=open_aa_input)
aa_menu.add_command(label="Open AA Output 1", command=open_aa_output1)
aa_menu.add_command(label="Open AA Output 2", command=open_aa_output2)
aa_menu.add_command(label="Open AA Output 3", command=open_aa_output3)
aa_menu.add_command(label="Open AA Output 4", command=open_aa_output4)
aa_menu.add_command(label="Open AA Output 5", command=open_aa_output5)
open_boxes_menu.add_cascade(label="AA Boxes", menu=aa_menu)
# BB Group submenu.
bb_menu = tk.Menu(open_boxes_menu, tearoff=0)
bb_menu.add_command(label="Open BB Input", command=open_bb_input)
bb_menu.add_command(label="Open BB Output 1", command=open_bb_output1)
bb_menu.add_command(label="Open BB Output 2", command=open_bb_output2)
bb_menu.add_command(label="Open BB Output 3", command=open_bb_output3)
bb_menu.add_command(label="Open BB Output 4", command=open_bb_output4)
bb_menu.add_command(label="Open BB Output 5", command=open_bb_output5)
open_boxes_menu.add_cascade(label="BB Boxes", menu=bb_menu)
# CC Group submenu.
cc_menu = tk.Menu(open_boxes_menu, tearoff=0)
cc_menu.add_command(label="Open CC Input", command=open_cc_input)
cc_menu.add_command(label="Open CC Output 1", command=open_cc_output1)
cc_menu.add_command(label="Open CC Output 2", command=open_cc_output2)
cc_menu.add_command(label="Open CC Output 3", command=open_cc_output3)
cc_menu.add_command(label="Open CC Output 4", command=open_cc_output4)
cc_menu.add_command(label="Open CC Output 5", command=open_cc_output5)
open_boxes_menu.add_cascade(label="CC Boxes", menu=cc_menu)
# DD Group submenu.
dd_menu = tk.Menu(open_boxes_menu, tearoff=0)
dd_menu.add_command(label="Open DD Input", command=open_dd_input)
dd_menu.add_command(label="Open DD Output 1", command=open_dd_output1)
dd_menu.add_command(label="Open DD Output 2", command=open_dd_output2)
dd_menu.add_command(label="Open DD Output 3", command=open_dd_output3)
dd_menu.add_command(label="Open DD Output 4", command=open_dd_output4)
dd_menu.add_command(label="Open DD Output 5", command=open_dd_output5)
open_boxes_menu.add_cascade(label="DD Boxes", menu=dd_menu)

menubar.add_cascade(label="Open Boxes", menu=open_boxes_menu)

# "Monitoring" drop-down menu.
monitor_menu = tk.Menu(menubar, tearoff=0)
monitor_menu.add_command(label="Start All Monitoring", command=start_all_monitoring)
monitor_menu.add_command(label="Stop All Monitoring", command=stop_all_monitoring)

# Group-specific monitoring submenu.
group_monitor_menu = tk.Menu(monitor_menu, tearoff=0)
group_monitor_menu.add_command(label="Start AA Monitoring", command=start_aa_monitoring)
group_monitor_menu.add_command(label="Stop AA Monitoring", command=stop_aa_monitoring)
group_monitor_menu.add_command(label="Start BB Monitoring", command=start_bb_monitoring)
group_monitor_menu.add_command(label="Stop BB Monitoring", command=stop_bb_monitoring)
group_monitor_menu.add_command(label="Start CC Monitoring", command=start_cc_monitoring)
group_monitor_menu.add_command(label="Stop CC Monitoring", command=stop_cc_monitoring)
group_monitor_menu.add_command(label="Start DD Monitoring", command=start_dd_monitoring)
group_monitor_menu.add_command(label="Stop DD Monitoring", command=stop_dd_monitoring)
monitor_menu.add_cascade(label="Group Monitoring", menu=group_monitor_menu)

menubar.add_cascade(label="Monitoring", menu=monitor_menu)

root.config(menu=menubar)

# Start the keyboard listener.
listener = keyboard.Listener(on_press=on_press)
listener.start()

root.mainloop()
