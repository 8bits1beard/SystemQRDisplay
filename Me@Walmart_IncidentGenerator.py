import tkinter as tk
from tkinter import ttk
from PIL import Image, ImageTk
import qrcode
import socket
import wmi
from datetime import datetime
import os

def fetch_fqdn_details():
    try:
        fqdn = socket.getfqdn()
        parts = fqdn.split('.')
        if len(parts) >= 4:
            return parts[1], '.'.join(parts[2:-1])
        return "Unknown", "Unknown"
    except Exception as e:
        print(f"Error retrieving FQDN details: {e}")
        return "Error", "Error"

def get_wmi_info():
    c = wmi.WMI()
    system = c.Win32_ComputerSystem()[0]
    os_info = c.Win32_OperatingSystem()[0]
    bios = c.Win32_BIOS()[0]
    disk = c.Win32_DiskDrive()[0]
    return system, os_info, bios, disk

def get_system_info():
    try:
        system, os_info, bios, disk = get_wmi_info()
        store_number, domain = fetch_fqdn_details()
        last_reboot = datetime.strptime(os_info.LastBootUpTime.split('.')[0], "%Y%m%d%H%M%S")
        formatted_last_reboot = last_reboot.strftime("%m/%d/%Y %H:%M")
        
        info = {
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
            "Last Reboot": formatted_last_reboot
        }
        return info
    except Exception as e:
        print(f"Error retrieving system information: {e}")
        return {}

def generate_qr_code(url):
    try:
        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_L,
            box_size=10,
            border=4,
        )
        qr.add_data(url)
        qr.make(fit=True)
        return qr.make_image(fill='black', back_color='white').resize((200, 200), Image.Resampling.LANCZOS)
    except Exception as e:
        print(f"Error generating QR code: {e}")
        return None

def load_image(path, size=(200, 200)):
    try:
        img = Image.open(path)
        return ImageTk.PhotoImage(img.resize(size, Image.Resampling.LANCZOS))
    except Exception as e:
        print(f"Error loading image from {path}: {e}")
        return None

def show_window(android_qr, ios_qr, system_info):
    root = tk.Tk()
    root.title("Report Issue with the Me@Walmart App for Android / iOS")

    android_photo = ImageTk.PhotoImage(Image.open("google_play_badge.png"))
    android_label = ttk.Label(root, image=android_photo)
    android_label.image = android_photo
    android_label.grid(row=1, column=0, padx=10, pady=10)

    android_qr_photo = ImageTk.PhotoImage(android_qr)
    android_qr_label = ttk.Label(root, image=android_qr_photo)
    android_qr_label.image = android_qr_photo
    android_qr_label.grid(row=2, column=0, padx=10, pady=10)

    ios_photo = ImageTk.PhotoImage(Image.open("app_store_badge.png"))
    ios_label = ttk.Label(root, image=ios_photo)
    ios_label.image = ios_photo
    ios_label.grid(row=1, column=1, padx=10, pady=10)

    ios_qr_photo = ImageTk.PhotoImage(ios_qr)
    ios_qr_label = ttk.Label(root, image=ios_qr_photo)
    ios_qr_label.image = ios_qr_photo
    ios_qr_label.grid(row=2, column=1, padx=10, pady=10)

    # Display system information
    row = 3
    for key, value in system_info.items():
        ttk.Label(root, text=f"{key}:", font=("Helvetica", 10, "bold")).grid(row=row, column=0, sticky=tk.E)
        ttk.Label(root, text=value, font=("Helvetica", 10)).grid(row=row, column=1, sticky=tk.W)
        row += 1

    # Close button
    close_button = ttk.Button(root, text="Close", command=root.destroy)
    close_button.grid(row=row, column=0, columnspan=2, pady=20)

    # Centering the window
    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - (root.winfo_width() // 2)
    y = (root.winfo_screenheight() // 2) - (root.winfo_height() // 2)
    root.geometry(f"+{x}+{y}")

    root.mainloop()

if __name__ == "__main__":
    android_url = "https://play.google.com/store/apps/details?id=com.walmart.squiggly&utm_campaign=WinEng_2025TicketReduction&pcampaignid=pcampaignidMKT-Other-global-all-co-prtnr-py-PartBadge-Mar2515-1"
    ios_url = "https://apps.apple.com/us/app/me-walmart/id1459898418"
    android_qr = generate_qr_code(android_url)
    ios_qr = generate_qr_code(ios_url)
    system_info = get_system_info()
    show_window(android_qr, ios_qr, system_info)


