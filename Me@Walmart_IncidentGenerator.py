"""
Script Name: System Info and QR Code Generator
Description: Retrieves system information, generates QR codes for specified URLs, and displays all these in a GUI. It is designed to help users download specific apps via QR codes and view their system information.
Author: Joshua Walderbach, Windows Engineering OS
Date: 03 MAY 2024
"""

import subprocess
import sys

# Ensure all required modules are imported first
try:
    import wmi
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "WMI"])
    import wmi  # Re-import after installation

try:
    from PIL import Image, ImageTk
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "Pillow"])
    from PIL import Image, ImageTk  # Re-import after installation

try:
    import qrcode
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "qrcode[pil]"])
    import qrcode  # Re-import after installation

# Constants for QR Code generation, now that qrcode is definitely imported
QR_VERSION = 7
QR_ERROR_CORRECT_LEVEL = qrcode.constants.ERROR_CORRECT_L
QR_BOX_SIZE = 10
QR_BORDER = 2

import tkinter as tk
from tkinter import ttk
import socket
from datetime import datetime

def fetch_fqdn_details():
    """Fetch the fully qualified domain name and extract useful parts."""
    try:
        fqdn = socket.getfqdn()
        parts = fqdn.split('.')
        if len(parts) >= 4:
            return parts[1], '.'.join(parts[2:-1])
        return "Unknown", "Unknown"
    except Exception as e:
        print(f"Error retrieving FQDN details: {str(e)}")
        return "Error", "Error"

def get_wmi_info():
    """Fetch system hardware information using WMI."""
    try:
        c = wmi.WMI()
        return c.Win32_ComputerSystem()[0], c.Win32_OperatingSystem()[0], c.Win32_BIOS()[0], c.Win32_DiskDrive()[0]
    except wmi.x_wmi as e:
        print(f"Error accessing WMI data: {str(e)}")
        return None

def get_system_info():
    """Compile system information into a dictionary."""
    system, os_info, bios, disk = get_wmi_info()
    if system and os_info and bios and disk:
        store_number, domain = fetch_fqdn_details()
        last_reboot = datetime.strptime(os_info.LastBootUpTime.split('.')[0], "%Y%m%d%H%M%S")
        formatted_last_reboot = last_reboot.strftime("%m/%d/%Y %H:%M")
        
        return {
            "Machine Name": socket.gethostname(),
            "Store Number": store_number,
            "Domain": domain,
            "IP Address": socket.gethostbyname(socket.gethostname()),
            "Manufacturer": system.Manufacturer,
            "Model": system.Model,
            "Serial Number": bios.SerialNumber,
            "Disk Size": f"{int(disk.Size) / (1024**3):.2f} GB",
            "RAM Total": f"{int(system.TotalPhysicalMemory) / (1024**3):.2f} GB",
            "Windows OS Version": os_info.Caption,
            "Last Reboot": formatted_last_reboot,
        }
    else:
        return {}

def generate_qr_code(url):
    """Generate and return a QR code image from a URL."""
    try:
        qr = qrcode.QRCode(version=QR_VERSION, error_correction=QR_ERROR_CORRECT_LEVEL, box_size=QR_BOX_SIZE, border=QR_BORDER)
        qr.add_data(url)
        qr.make(fit=True)
        return qr.make_image(fill="black", back_color="white").resize((200, 200), Image.Resampling.LANCZOS)
    except Exception as e:
        print(f"Error generating QR code: {str(e)}")
        return None

def load_image(path):
    """Load an image from a specified path."""
    try:
        img = Image.open(path)
        return ImageTk.PhotoImage(img)
    except FileNotFoundError:
        print(f"Error: The file {path} does not exist.")
        return None
    except Exception as e:
        print(f"Error loading image from {path}: {str(e)}")
        return None

# Main application window setup
def show_window(android_qr, ios_qr, system_info):
    root = tk.Tk()
    root.title("Download the Me@Walmart App for Android or iOS")

    try:
        root.iconbitmap("icon.ico")
    except Exception as e:
        print(f"Error setting icon: {str(e)}")

    android_photo = load_image("google_play_badge.png")
    if android_photo:
        android_label = ttk.Label(root, image=android_photo)
        android_label.image = android_photo
        android_label.grid(row=1, column=0, padx=10, pady=10)

    ios_photo = load_image("app_store_badge.png")
    if ios_photo:
        ios_label = ttk.Label(root, image=ios_photo)
        ios_label.image = ios_photo
        ios_label.grid(row=1, column=1, padx=10, pady=10)

    android_qr_photo = ImageTk.PhotoImage(android_qr)
    android_qr_label = ttk.Label(root, image=android_qr_photo)
    android_qr_label.image = android_qr_photo
    android_qr_label.grid(row=2, column=0, padx=10, pady=10)

    ios_qr_photo = ImageTk.PhotoImage(ios_qr)
    ios_qr_label = ttk.Label(root, image=ios_qr_photo)
    ios_qr_label.image = ios_qr_photo
    ios_qr_label.grid(row=2, column=1, padx=10, pady=10)

    row = 3
    for key, value in system_info.items():
        ttk.Label(root, text=f"{key}:", font=("Helvetica", 10, "bold")).grid(row=row, column=0, sticky=tk.E)
        ttk.Label(root, text=value, font=("Helvetica", 10)).grid(row=row, column=1, sticky=tk.W)
        row += 1

    close_button = ttk.Button(root, text="Close", command=root.destroy)
    close_button.grid(row=row, column=0, columnspan=2, pady=20)

    root.mainloop()

if __name__ == "__main__":
    android_url = "https://play.google.com/store/apps/details?id=com.walmart.squiggly&utm_campaign=WinEng_2025TicketReduction&pcampaignid=pcampaignidMKT-Other-global-all-co-prtnr-py-PartBadge-Mar2515-1"
    ios_url = "https://apps.apple.com/us/app/me-walmart/id1459898418"
    android_qr = generate_qr_code(android_url)
    ios_qr = generate_qr_code(ios_url)
    system_info = get_system_info()
    show_window(android_qr, ios_qr, system_info)
