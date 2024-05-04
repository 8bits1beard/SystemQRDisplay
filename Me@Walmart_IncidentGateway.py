import tkinter as tk
from tkinter import ttk
from PIL import Image, ImageTk
import qrcode
import socket
import wmi
from datetime import datetime

def get_system_info():
    c = wmi.WMI()
    system = c.Win32_ComputerSystem()[0]
    os_info = c.Win32_OperatingSystem()[0]
    bios = c.Win32_BIOS()[0]
    disk = c.Win32_DiskDrive()[0]
    last_reboot = datetime.strptime(os_info.LastBootUpTime.split('.')[0], "%Y%m%d%H%M%S")
    formatted_last_reboot = last_reboot.strftime("%m/%d/%Y %H:%M")
    info = (
        f"Machine Name: {socket.gethostname()}\n"
        f"IP Address: {socket.gethostbyname(socket.gethostname())}\n"
        f"Manufacturer: {system.Manufacturer}\n"
        f"Model: {system.Model}\n"
        f"Serial Number: {bios.SerialNumber}\n"
        f"Disk Size: {int(disk.Size) / (1024**3):.2f} GB\n"
        f"RAM Total: {int(system.TotalPhysicalMemory) / (1024**3):.2f} GB\n"
        f"Windows OS Version: {os_info.Caption}\n"
        f"Last Reboot: {formatted_last_reboot}"
    )
    return info

def generate_qr_code(url):
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=10,
        border=4,
    )
    qr.add_data(url)
    qr.make(fit=True)
    img = qr.make_image(fill='black', back_color='white')
    img = img.resize((256, 256), Image.Resampling.LANCZOS)
    return img

def show_window(android_qr, ios_qr, system_info):
    root = tk.Tk()
    root.title("Report Issue with the Me@Walmart App for Android / iOS")

    android_photo = ImageTk.PhotoImage(Image.open("C:\\temp\\google_play_badge.png"))
    android_label = ttk.Label(root, image=android_photo)
    android_label.image = android_photo
    android_label.grid(row=1, column=0, padx=10, pady=10)

    android_qr_photo = ImageTk.PhotoImage(android_qr)
    android_qr_label = ttk.Label(root, image=android_qr_photo)
    android_qr_label.image = android_qr_photo
    android_qr_label.grid(row=2, column=0, padx=10, pady=10)

    ios_photo = ImageTk.PhotoImage(Image.open("C:\\temp\\app_store_badge.png"))
    ios_label = ttk.Label(root, image=ios_photo)
    ios_label.image = ios_photo
    ios_label.grid(row=1, column=1, padx=10, pady=10)

    ios_qr_photo = ImageTk.PhotoImage(ios_qr)
    ios_qr_label = ttk.Label(root, image=ios_qr_photo)
    ios_qr_label.image = ios_qr_photo
    ios_qr_label.grid(row=2, column=1, padx=10, pady=10)

    info_text = ttk.Label(root, text=system_info, font=("Helvetica", 10, "bold"), justify=tk.LEFT)
    info_text.grid(row=3, column=0, columnspan=2, padx=10, pady=40)

    # Close button using grid with appropriate spacing
    close_button = ttk.Button(root, text="Close", command=root.destroy)
    close_button.grid(row=4, column=0, columnspan=2, pady=40)

    # Centering the window and ensuring adequate size
    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - (root.winfo_width() // 2)
    y = (root.winfo_screenheight() // 2) - (root.winfo_height() // 2)
    root.geometry(f"+{x}+{y}")
    root.minsize(root.winfo_width(), root.winfo_height())

    root.mainloop()

if __name__ == "__main__":
    android_url = "https://play.google.com/store/apps/details?id=com.walmart.squiggly&utm_campaign=WinEng_2025TicketReduction&pcampaignid=pcampaignidMKT-Other-global-all-co-prtnr-py-PartBadge-Mar2515-1"
    ios_url = "https://apps.apple.com/us/app/me-walmart/id1459898418"
    android_qr = generate_qr_code(android_url)
    ios_qr = generate_qr_code(ios_url)
    system_info = get_system_info()
    show_window(android_qr, ios_qr, system_info)
