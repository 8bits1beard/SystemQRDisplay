"""
Script Name: SystemInfoQRApp.py
Description: This script generates a GUI application that displays system information,
             generates QR codes for URLs, and manages system-related tasks such as fetching
             Windows registry values and update history. It ensures all dependencies are
             installed before executing GUI related operations.

Dependencies: tkinter, PIL (Pillow), qrcode, wmi, pywin32, datetime, socket, winreg
Usage: Run this script directly with Python 3. Ensure you have administrative rights if required
       by any of the system information fetching operations.
Author: Joshua Walderbach
"""

import subprocess
import sys

def install_and_import(package, import_name=None):
    """
    Install and import a package using pip. Specify an import_name if the import
    syntax differs from the package name.
    
    Args:
        package (str): The package name to install.
        import_name (str): The name used to import the package in scripts.
    """
    if import_name is None:
        import_name = package
    try:
        __import__(import_name)
    except ImportError:
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", package])
        except subprocess.CalledProcessError:
            print(f"Failed to install {package}")
            return
        try:
            __import__(import_name)
        except ImportError:
            print(f"Failed to import {import_name} after installation. Please check the installation.")

# List of third-party packages to check and potentially install
packages = [
    ("Pillow", "PIL"),
    ("qrcode", None),
    ("wmi", None),
    ("pywin32", "win32com.client")
]

for package, import_name in packages:
    install_and_import(package, import_name)

# Safe imports after ensuring dependencies are present
import socket
import tkinter as tk
from tkinter import ttk
from tkinter.font import Font
from PIL import Image, ImageTk
import qrcode
from datetime import datetime
import wmi
import win32com.client
import winreg

# Constants for QR Code generation
QR_CONFIG = {
    'version': 7,
    'error_correction': qrcode.constants.ERROR_CORRECT_L,
    'box_size': 10,
    'border': 2,
}

def get_registry_value(path, name):
    """
    Retrieves a value from the Windows registry.
    
    Args:
        path (str): Registry path.
        name (str): Name of the registry value.
    
    Returns:
        str: The value from the registry or 'UNKNOWN' if not found.
    """
    try:
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path)
        value, regtype = winreg.QueryValueEx(key, name)
        winreg.CloseKey(key)
        return value
    except FileNotFoundError:
        default_value = "UNKNOWN"
        print(f"Registry key {path}\\{name} not found on this system. Using default value: {default_value}")
        return default_value
    except Exception as e:
        print(f"Unexpected error accessing registry {path}\\{name}: {str(e)}")
        return None

def get_latest_update():
    """
    Retrieves the title of the latest Cumulative Update from the Windows Update history.
    
    Returns:
        str: The title of the latest update or 'No Updates Found'.
    """
    try:
        wua = win32com.client.Dispatch("Microsoft.Update.Session")
        searcher = wua.CreateUpdateSearcher()
        history_count = searcher.GetTotalHistoryCount()
        history = searcher.QueryHistory(0, history_count)
        for update in history:
            if "Cumulative Update" in update.Title:
                return update.Title
        return "No Updates Found"
    except Exception as e:
        print(f"Error fetching update history: {str(e)}")
        return "Error Fetching Updates"

def fetch_system_info():
    """
    Retrieves various system information such as hostname, IP address, and OS details.
    
    Returns:
        dict: A dictionary containing key system information.
    """
    c = wmi.WMI()
    system = c.Win32_ComputerSystem()[0]
    os_info = c.Win32_OperatingSystem()[0]
    bios = c.Win32_BIOS()[0]
    disk = c.Win32_DiskDrive()[0]
    hostname = socket.gethostname()
    ip_address = socket.gethostbyname(hostname)
    last_reboot = datetime.strptime(os_info.LastBootUpTime.split('.')[0], "%Y%m%d%H%M%S")
    installed_on = datetime.strptime(os_info.InstallDate.split('.')[0], "%Y%m%d%H%M%S")
    role = get_registry_value(r"SOFTWARE\Workstation\Build", "Role")
    latest_update = get_latest_update()

    disk_size_gb = f"{round(int(disk.Size) / (1024**3))} GB"
    memory_gb = f"{round(int(system.TotalPhysicalMemory) / (1024**3))} GB"

    return {
        'Build Date': installed_on.strftime("%m/%d/%Y %H:%M"),
        'Role': role,
        'Computer Name': hostname,
        'Manufacturer': system.Manufacturer,
        'Model': system.Model,
        'Serial': bios.SerialNumber,
        'Disk Size': disk_size_gb,
        'Memory': memory_gb,
        'IP Address': ip_address,
        'OS Build': os_info.Caption,
        'Latest Update': latest_update,
        'Last Reboot': last_reboot.strftime("%m/%d/%Y %H:%M")
    }

def generate_qr_code(url):
    """
    Creates a QR code image for a given URL.
    
    Args:
        url (str): URL to generate a QR code for.
    
    Returns:
        Image: A PIL image object of the QR code.
    """
    qr = qrcode.QRCode(**QR_CONFIG)
    qr.add_data(url)
    qr.make(fit=True)
    return qr.make_image(fill='black', back_color='white').resize((200, 200), Image.Resampling.LANCZOS)

def create_gui(root, android_qr_img, ios_qr_img, system_info):
    """
    Generates a GUI window displaying Android and iOS QR codes along with system information using image badges.
    
    Args:
        root (tk.Tk): The root window.
        android_qr_img (ImageTk.PhotoImage): Image for Android QR code.
        ios_qr_img (ImageTk.PhotoImage): Image for iOS QR code.
        system_info (dict): Dictionary containing system information.
    """
    root.title("Download the Me@Walmart App for Android or iOS")

    header_font = Font(root, ("Helvetica", 12, "bold"))
    body_font = Font(root, ("Helvetica", 11, "normal"))

    android_badge_image = Image.open("google_play_badge.png")
    ios_badge_image = Image.open("app_store_badge.png")
    android_badge_photo = ImageTk.PhotoImage(android_badge_image)
    ios_badge_photo = ImageTk.PhotoImage(ios_badge_image)

    main_frame = ttk.Frame(root, padding="10 10 10 10")
    main_frame.grid(row=0, column=0, sticky=(tk.N, tk.W, tk.E, tk.S))
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)

    android_frame = ttk.Frame(main_frame)
    android_frame.grid(row=0, column=0, sticky='ew', padx=150)
    android_badge_label = ttk.Label(android_frame, image=android_badge_photo)
    android_badge_label.pack(side='top', fill='x')
    android_label = ttk.Label(android_frame, image=android_qr_img)
    android_label.pack(side='top', fill='x')

    ios_frame = ttk.Frame(main_frame)
    ios_frame.grid(row=0, column=1, sticky='ew', padx=(20, 150))
    ios_badge_label = ttk.Label(ios_frame, image=ios_badge_photo)
    ios_badge_label.pack(side='top', fill='x')
    ios_label = ttk.Label(ios_frame, image=ios_qr_img)
    ios_label.pack(side='top', fill='x')

    info_frame = ttk.Frame(main_frame)
    info_frame.grid(row=1, column=0, columnspan=2, sticky='ew')
    row = 0
    for key, value in system_info.items():
        ttk.Label(info_frame, text=f"{key}:", font=header_font).grid(row=row, column=0, sticky='e', padx=10, pady=2)
        ttk.Label(info_frame, text=value, font=body_font).grid(row=row, column=1, sticky='w', padx=5, pady=2)
        row += 1

    style = ttk.Style()
    style.configure('Red.TButton', font=('Helvetica', 10, 'bold'), foreground='red', background='white')

    close_button = ttk.Button(info_frame, text="CLOSE THIS WINDOW", command=root.destroy, style='Red.TButton')
    close_button.grid(row=row, column=0, columnspan=2, pady=10, sticky='nsew', padx=10)

    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')

    root.mainloop()

if __name__ == "__main__":
    root = tk.Tk()
    root.iconbitmap('icon.ico')  # Set the icon for the window
    android_url = "https://play.google.com/store/apps/details?id=com.example"
    ios_url = "https://apps.apple.com/us/app/example"
    android_qr = generate_qr_code(android_url)
    ios_qr = generate_qr_code(ios_url)
    system_info = fetch_system_info()

    android_qr_img = ImageTk.PhotoImage(android_qr)
    ios_qr_img = ImageTk.PhotoImage(ios_qr)

    create_gui(root, android_qr_img, ios_qr_img, system_info)
