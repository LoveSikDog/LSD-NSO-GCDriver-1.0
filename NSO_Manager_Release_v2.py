import sys
import os
import threading
import queue
import asyncio
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import webbrowser
import traceback
import ctypes.wintypes
from PIL import Image, ImageDraw

# ==========================================
# CONFIGURATION
# ==========================================
APP_TITLE = "NSO GameCube Driver"
CREATOR = "Created by LoveSikDogs"
VERSION = "v1.0"
NINTENDO_OUI = "3CA9AB"
DOLPHIN_INI_NAME = "lsd_nso.ini" 

# TUNING
STICK_MULTIPLIER = 32.0 
DEADZONE         = 0.15
TRIGGER_OFFSET   = 40
TRIGGER_SCALE    = 1.6 

stop_event = threading.Event()
gui_queue = queue.Queue()

# GLOBAL LAZY LOADERS
vg = None
BleakScanner = None
BleakClient = None
Icon = None
item = None
Menu = None

# ==========================================
# WINDOWS UTILS
# ==========================================
def get_real_documents_path():
    """
    Asks Windows for the TRUE Documents folder path.
    This works even if the user moved their Documents to a different drive (D:, E:, etc).
    """
    CSIDL_PERSONAL = 5       # My Documents
    SHGFP_TYPE_CURRENT = 0   # Get current, not default path
    buf = ctypes.create_unicode_buffer(ctypes.wintypes.MAX_PATH)
    ctypes.windll.shell32.SHGetFolderPathW(None, CSIDL_PERSONAL, None, SHGFP_TYPE_CURRENT, buf)
    return buf.value

def is_vigem_installed():
    """Checks for the physical driver file."""
    driver_path = os.path.join(os.environ['WINDIR'], "System32", "drivers", "ViGEmBus.sys")
    return os.path.exists(driver_path)

# ==========================================
# DOLPHIN PROFILE LOGIC
# ==========================================
PROFILE_CONTENT = """[Profile]
Device = XInput/0/Gamepad
Buttons/A = `Button A`
Buttons/B = `Button Y`
Buttons/X = `Button X`
Buttons/Y = `Button B`
Buttons/Z = `Shoulder R`
Buttons/Start = Start
Main Stick/Up = `Left Y+`
Main Stick/Down = `Left Y-`
Main Stick/Left = `Left X-`
Main Stick/Right = `Left X+`
Main Stick/Modifier = `Shift`
Main Stick/Calibration = 100.00 141.42 100.00 141.42 100.00 141.42 100.00 141.42
C-Stick/Up = `Right Y+`
C-Stick/Down = `Right Y-`
C-Stick/Left = `Right X-`
C-Stick/Right = `Right X+`
C-Stick/Modifier = `Ctrl`
C-Stick/Calibration = 100.00 141.42 100.00 141.42 100.00 141.42 100.00 141.42
Triggers/L = `Trigger L`
Triggers/R = `Trigger R`
Triggers/L-Analog = `Trigger L`
Triggers/R-Analog = `Trigger R`
D-Pad/Up = `Pad N`
D-Pad/Down = `Pad E`
D-Pad/Left = `Pad S`
D-Pad/Right = `Pad W`
"""

def auto_install_profile():
    """Attempts to automatically drop the file in the standard location."""
    try:
        docs = get_real_documents_path()
        dolphin_path = os.path.join(docs, "Dolphin Emulator", "Config", "Profiles", "GCPad")
        
        # We only auto-install if we see the folder structure exists
        if os.path.exists(os.path.join(docs, "Dolphin Emulator")):
            os.makedirs(dolphin_path, exist_ok=True)
            target_file = os.path.join(dolphin_path, DOLPHIN_INI_NAME)
            
            with open(target_file, "w") as f:
                f.write(PROFILE_CONTENT)
            return True
    except:
        pass
    return False

def manual_export_profile():
    """Allows the user to save the .ini file manually via a popup."""
    # We must use the root window for the dialog, but it might be hidden
    # So we create a temporary hidden window if needed
    root = tk.Tk()
    root.withdraw() 
    
    file_path = filedialog.asksaveasfilename(
        title="Save Dolphin Profile",
        initialfile=DOLPHIN_INI_NAME,
        defaultextension=".ini",
        filetypes=[("Configuration Settings", "*.ini")]
    )
    
    if file_path:
        try:
            with open(file_path, "w") as f:
                f.write(PROFILE_CONTENT)
            messagebox.showinfo("Success", f"Profile saved to:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not save file:\n{e}")
    root.destroy()

# ==========================================
# INPUT MAPPING ENGINE
# ==========================================
def map_input(report, xbox):
    try:
        if report[2] & 0x02: xbox.press_button(button=vg.XUSB_BUTTON.XUSB_GAMEPAD_A)
        else:                xbox.release_button(button=vg.XUSB_BUTTON.XUSB_GAMEPAD_A)
        if report[2] & 0x01: xbox.press_button(button=vg.XUSB_BUTTON.XUSB_GAMEPAD_B)
        else:                xbox.release_button(button=vg.XUSB_BUTTON.XUSB_GAMEPAD_B)
        if report[2] & 0x08: xbox.press_button(button=vg.XUSB_BUTTON.XUSB_GAMEPAD_X)
        else:                xbox.release_button(button=vg.XUSB_BUTTON.XUSB_GAMEPAD_X)
        if report[2] & 0x04: xbox.press_button(button=vg.XUSB_BUTTON.XUSB_GAMEPAD_Y)
        else:                xbox.release_button(button=vg.XUSB_BUTTON.XUSB_GAMEPAD_Y)
        if report[2] & 0x40: xbox.press_button(button=vg.XUSB_BUTTON.XUSB_GAMEPAD_START)
        else:                xbox.release_button(button=vg.XUSB_BUTTON.XUSB_GAMEPAD_START)
        if report[2] & 0x20: xbox.press_button(button=vg.XUSB_BUTTON.XUSB_GAMEPAD_RIGHT_SHOULDER)
        else:                xbox.release_button(button=vg.XUSB_BUTTON.XUSB_GAMEPAD_RIGHT_SHOULDER)

        if report[3] & 0x08: xbox.press_button(button=vg.XUSB_BUTTON.XUSB_GAMEPAD_DPAD_UP)
        else:                xbox.release_button(button=vg.XUSB_BUTTON.XUSB_GAMEPAD_DPAD_UP)
        if report[3] & 0x01: xbox.press_button(button=vg.XUSB_BUTTON.XUSB_GAMEPAD_DPAD_DOWN)
        else:                xbox.release_button(button=vg.XUSB_BUTTON.XUSB_GAMEPAD_DPAD_DOWN)
        if report[3] & 0x04: xbox.press_button(button=vg.XUSB_BUTTON.XUSB_GAMEPAD_DPAD_LEFT)
        else:                xbox.release_button(button=vg.XUSB_BUTTON.XUSB_GAMEPAD_DPAD_LEFT)
        if report[3] & 0x02: xbox.press_button(button=vg.XUSB_BUTTON.XUSB_GAMEPAD_DPAD_RIGHT)
        else:                xbox.release_button(button=vg.XUSB_BUTTON.XUSB_GAMEPAD_DPAD_RIGHT)
        
        if report[4] & 0x10: xbox.press_button(button=vg.XUSB_BUTTON.XUSB_GAMEPAD_GUIDE)
        else:                xbox.release_button(button=vg.XUSB_BUTTON.XUSB_GAMEPAD_GUIDE)

        raw_l = report[12]
        raw_r = report[13]
        is_l_click = (report[3] & 0x10) != 0
        is_r_click = (report[2] & 0x10) != 0

        if is_l_click: final_l = 255
        else:
            final_l = int(max(0, raw_l - TRIGGER_OFFSET) * TRIGGER_SCALE)
            if final_l > 255: final_l = 255

        if is_r_click: final_r = 255
        else:
            final_r = int(max(0, raw_r - TRIGGER_OFFSET) * TRIGGER_SCALE)
            if final_r > 255: final_r = 255

        xbox.left_trigger(value=final_l)
        xbox.right_trigger(value=final_r)

        raw_lx = report[5] | ((report[6] & 0x0F) << 8)
        raw_ly = (report[6] >> 4) | (report[7] << 4)
        raw_cx = report[8] | ((report[9] & 0x0F) << 8)
        raw_cy = (report[9] >> 4) | (report[10] << 4)

        def scale(val):
            centered = val - 2048
            if abs(centered) < (2048 * DEADZONE): return 0
            x = int(centered * STICK_MULTIPLIER)
            return max(-32768, min(32767, x))

        xbox.left_joystick(x_value=scale(raw_lx), y_value=scale(raw_ly))
        xbox.right_joystick(x_value=scale(raw_cx), y_value=scale(raw_cy))
        
        xbox.update()
    except:
        pass

# ==========================================
# DRIVER LOOP
# ==========================================
async def driver_loop():
    global vg, BleakScanner, BleakClient
    
    gui_queue.put(("Loading", "gray", "Loading Drivers..."))
    
    if not is_vigem_installed():
        gui_queue.put(("MISSING_DRIVER", "red", "ViGEmBus Not Found"))
        return

    try:
        import vgamepad
        vg = vgamepad
        from bleak import BleakScanner as BS, BleakClient as BC
        BleakScanner = BS
        BleakClient = BC
    except ImportError as e:
        gui_queue.put(("Error", "red", f"Library Error: {e}"))
        return

    try:
        xbox = vg.VX360Gamepad()
    except Exception as e:
        gui_queue.put(("Error", "red", f"Driver Error: {e}"))
        return

    while not stop_event.is_set():
        gui_queue.put(("Scanning", "orange", "Scanning for Nintendo Controller..."))
        found_device = None

        try:
            devices = await BleakScanner.discover(timeout=4.0)
            for d in devices:
                if d.address.replace(":", "").upper().startswith(NINTENDO_OUI):
                    found_device = d
                    break
        except Exception:
            gui_queue.put(("Error", "red", "BLUETOOTH ERROR:\nIs Bluetooth on?"))
            await asyncio.sleep(2.0)
            continue

        if found_device:
            gui_queue.put(("Found", "blue", f"Found: {found_device.address}\nConnecting..."))
            
            try:
                async with BleakClient(found_device.address, timeout=10.0, services=None) as client:
                    gui_queue.put(("Connected", "green", f"✅ CONNECTED\n{found_device.address}"))
                    try: await client.get_services()
                    except: pass

                    async def input_handler(sender, data):
                        if len(data) > 13: map_input(list(data), xbox)

                    for service in client.services:
                        for char in service.characteristics:
                            if "notify" in char.properties:
                                await client.start_notify(char.uuid, input_handler)
                    
                    while client.is_connected and not stop_event.is_set():
                        await asyncio.sleep(1)
                    
                    gui_queue.put(("Lost", "orange", "Signal Lost. Rescanning..."))
            except Exception:
                gui_queue.put(("Error", "orange", "Connection Failed. Retrying..."))
                await asyncio.sleep(1.0)
        
        await asyncio.sleep(0.5)

# ==========================================
# GUI APPLICATION
# ==========================================
class DriverApp:
    def __init__(self, root):
        self.root = root
        self.root.title(f"{APP_TITLE} {VERSION}")
        self.root.geometry("400x260")
        self.root.resizable(False, False)

        style = ttk.Style()
        style.configure("Header.TLabel", font=("Segoe UI", 12, "bold"))
        style.configure("Status.TLabel", font=("Segoe UI", 10))
        style.configure("Footer.TLabel", font=("Segoe UI", 8))

        ttk.Label(root, text=APP_TITLE, style="Header.TLabel").pack(pady=(15, 5))
        ttk.Label(root, text=VERSION, style="Footer.TLabel").pack()
        
        self.status_lbl = ttk.Label(root, text="Initializing...", style="Status.TLabel", justify="center")
        self.status_lbl.pack(pady=30)
        
        ttk.Label(root, text=CREATOR, style="Footer.TLabel", foreground="#555").pack(side="bottom", pady=5)
        ttk.Label(root, text="(Close window to minimize to Tray)", style="Footer.TLabel").pack(side="bottom")

        self.running = True
        threading.Thread(target=self.start_asyncio, daemon=True).start()
        self.check_queue()
        
        # Auto Install Profile (Silent)
        auto_install_profile()

        self.root.protocol('WM_DELETE_WINDOW', self.minimize_to_tray)
        self.tray_running = False

    def start_asyncio(self):
        asyncio.run(driver_loop())

    def check_queue(self):
        if not self.running: return
        try:
            while True:
                status_type, color, text = gui_queue.get_nowait()
                if status_type == "MISSING_DRIVER":
                    ans = messagebox.askyesno("Missing Driver", 
                        "The ViGEmBus Driver is missing!\n\n"
                        "This program requires it to create the Virtual Controller.\n"
                        "Would you like to download it now?")
                    if ans:
                        webbrowser.open("https://github.com/nefarius/ViGEmBus/releases/latest")
                    self.root.destroy()
                    sys.exit()
                else:
                    self.status_lbl.config(text=text, foreground=color)
        except queue.Empty:
            pass
        finally:
            if self.running: self.root.after(100, self.check_queue)

    def minimize_to_tray(self):
        global Icon, item, Menu
        
        if Icon is None:
            try:
                from pystray import Icon as I, MenuItem as MI, Menu as M
                Icon, item, Menu = I, MI, M
            except ImportError:
                return 

        self.root.withdraw()
        if not self.tray_running:
            self.tray_running = True
            threading.Thread(target=self.run_tray, daemon=True).start()

    def run_tray(self):
        def show(icon, item): 
            self.root.after(0, self.root.deiconify)
        
        def save_profile(icon, item):
            # Run the export in the GUI thread to allow popups
            self.root.after(0, manual_export_profile)

        def exit_app(icon, item):
            self.running = False
            stop_event.set()
            icon.stop()
            self.root.after(0, self.root.destroy)
            sys.exit()

        image = Image.new('RGB', (64, 64), (100, 65, 165))
        dc = ImageDraw.Draw(image)
        dc.rectangle((16, 16, 48, 48), fill=(255, 255, 255))
        
        # THE NEW TRAY MENU
        menu = Menu(
            item('Show Status', show),
            item('Save Dolphin Profile...', save_profile), # <-- THE NEW FEATURE
            item('Exit', exit_app)
        )
        
        icon = Icon("NSO_GC", image, "NSO Driver", menu)
        icon.run()

if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = DriverApp(root)
        root.mainloop()
    except Exception:
        sys.exit()