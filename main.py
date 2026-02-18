import subprocess
import time
import customtkinter as ctk
from PIL import Image
import os, sys
from threading import Thread
import webbrowser

working = False
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath("")

    return os.path.join(base_path, relative_path)

def check_activation_powershell():
    try:
        cmd = 'powershell "(Get-CimInstance -Query \'SELECT * FROM SoftwareLicensingProduct WHERE LicenseStatus=1\').count -gt 0"'
        result = subprocess.run(cmd, capture_output=True, text=True, shell=True)
        return "true" in result.stdout.lower()
    except Exception as e:
        print(f"PowerShell Error: {e}")
        return False

license_dict = {
    "Education": "slmgr /ipk NW6C2-QMPVW-D7KKK-3GKT6-VCFB2",
    "Enterprise":  "slmgr /ipk NPPR9-FWDCX-D2C8J-H872K-2YT43",
    "Home": "slmgr /ipk TX9XD-98N7V-6WMQ6-BX7FG-H8Q99",
    "Pro": "slmgr /ipk W269N-WFGWX-YVC9B-4J6C9-T83GX",
}
cmd1 = "slmgr /skms kms8.msguides.com"
cmd2 = "slmgr /ato"

uninstall = "slmgr.vbs -upk"


def get_windows_type():
    try:
        info = subprocess.check_output(
            'systeminfo', 
            stderr=subprocess.STDOUT,
            universal_newlines=True,
            shell=True
        )
        for line in info.split('\n'):
            if "OS Name" in line:
                windows =  line.split(":")[1].strip()
                return windows.split(" ")[-1]
    except subprocess.CalledProcessError:
        return "Unable to get system info"

def silent_cmd(cmd):
    try:
        subprocess.run(cmd, shell=True, creationflags=subprocess.CREATE_NO_WINDOW)
    except subprocess.CalledProcessError as e:
        print(f"Error installing product key: {e.stderr.decode('utf-8')}")

def active():
    global working

    working = True
    silent_cmd(license_dict.get(get_windows_type()))
    time.sleep(5)
    silent_cmd(cmd1)
    time.sleep(5)
    silent_cmd(cmd2)
    time.sleep(5)
    working = False

def deactive():
    global working
    
    working = True
    silent_cmd(uninstall)
    time.sleep(5)
    working = False

class GUI(ctk.CTk):

    def __init__(self, fg_color = None, **kwargs):
        super().__init__(fg_color, **kwargs)
        ctk.set_appearance_mode("dark")
        self.title("Free windows activation :)")
        self.iconbitmap(resource_path("favicon.ico"))
        self.resizable(0, 0)
        banner = Image.open(resource_path("banner.png"))
        ctk.CTkLabel(self, text="", image=ctk.CTkImage(dark_image=banner, light_image=banner, size=(400, 60))).pack()
        
        dev = ctk.CTkButton(self, text="Discord", height=10, command=lambda: webbrowser.open_new_tab("https://beacons.ai/youssefalkhodary"))
        dev.pack(fill='x')

        self.checking_label = ctk.CTkLabel(self, text="Checking activation...", font=("roboto", 12, "bold"))
        self.checking_label.pack(pady=5)

        self.loading_bar = ctk.CTkProgressBar(self, width=360, height=6, corner_radius=15, mode="indeterminate")
        self.loading_bar.pack(pady=(0, 5))
        self.loading_bar.start()

        self.main_button = ctk.CTkButton(self, text="Activate now", corner_radius=15, state="disabled")
        self.main_button.pack(pady=10)

        self.update()
        x = int((self.winfo_screenwidth() / 2) - (self.winfo_width() / 2))
        y = int((self.winfo_screenheight() / 2) - (self.winfo_height() / 2))
        self.geometry(f"{x}+{y - 50}")
        self.after(10, Thread(target=self.check_availability, daemon=True).start)

    def check_availability(self):
        if license_dict.get(get_windows_type()) == None:
            self.checking_label.configure(text="Current windows version is not suppported yet!", text_color="brown")
            self.stop_loading()
        else:
            Thread(target=self.check_activation, daemon=True).start()
    
    def start_loading(self):
        self.loading_bar.configure(mode="indeterminate")
        self.loading_bar.start()
    
    def stop_loading(self):
        self.loading_bar.stop()
        self.loading_bar.configure(mode="determinate")
        self.loading_bar.set(value=1.0)
    
    def check_activation(self):
        if check_activation_powershell():
            self.checking_label.configure(text="Windows is active.", text_color="cyan")
            self.main_button.configure(text="Remove activation", fg_color="brown", command=lambda: Thread(target=self.remove_activation, daemon=True).start())
            self.stop_loading()
        else:
            self.checking_label.configure(text="Windows not activated!", text_color="brown")
            self.main_button.configure(text="Activate now", command=lambda: Thread(target=self.active, daemon=True).start())
            self.stop_loading()
        self.main_button.configure(state="normal")

    def active(self):
        global working
        self.checking_label.configure(text="Activating...", text_color="cyan")
        self.main_button.configure(state="disabled")
        self.start_loading()
        Thread(target=active, daemon=True).start()
        while 1:
            if not working:
                self.checking_label.configure(text="Windows activated successfully!", text_color="green")
                break
            time.sleep(1)
        self.stop_loading()
        self.main_button.configure(state="normal")
        self.main_button.configure(text="Remove activation", fg_color="brown", command=lambda: Thread(target=self.remove_activation, daemon=True).start())

    def remove_activation(self):
        global working
        self.checking_label.configure(text="Removing activation...", text_color="brown")
        self.main_button.configure(state="disabled")
        self.start_loading()
        Thread(target=deactive, daemon=True).start()
        while 1:
            if not working:
                self.checking_label.configure(text="Windows activation removed!", text_color="red")
                break
            time.sleep(1)
        self.stop_loading()
        self.main_button.configure(state="normal")
        self.main_button.configure(text="Activate now", command=lambda: Thread(target=self.active, daemon=True).start())

GUI().mainloop()