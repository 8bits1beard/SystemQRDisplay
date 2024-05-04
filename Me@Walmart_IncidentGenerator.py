import socket
import tkinter as tk
from tkinter import ttk, font
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
    """Retrieves a value from the Windows registry."""
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
    """Retrieves the title of the latest Cumulative Update from the Windows Update history."""
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

    # Calculate disk size and memory in GB and append ' GB' suffix
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
    """Creates a QR code image for a given URL."""
    qr = qrcode.QRCode(**QR_CONFIG)
    qr.add_data(url)
    qr.make(fit=True)
    return qr.make_image(fill='black', back_color='white').resize((200, 200), Image.Resampling.LANCZOS)

def create_gui(root, android_qr_img, ios_qr_img, system_info):
    """Generates a GUI window displaying Android and iOS QR codes along with system information using image badges."""
    root.title("Download the Me@Walmart App for Android or iOS")

    # Define fonts
    header_font = Font(root, ("Helvetica", 12, "bold"))
    body_font = Font(root, ("Helvetica", 11, "normal"))

    # Load badge images
    android_badge_image = Image.open("google_play_badge.png")
    ios_badge_image = Image.open("app_store_badge.png")
    android_badge_photo = ImageTk.PhotoImage(android_badge_image)
    ios_badge_photo = ImageTk.PhotoImage(ios_badge_image)

    # Use a main frame to group elements
    main_frame = ttk.Frame(root, padding="10 10 10 10")
    main_frame.grid(row=0, column=0, sticky=(tk.N, tk.W, tk.E, tk.S))
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)

    # Frame for Android QR and badge
    android_frame = ttk.Frame(main_frame)
    android_frame.grid(row=0, column=0, sticky='ew', padx=20)
    android_badge_label = ttk.Label(android_frame, image=android_badge_photo)
    android_badge_label.pack(side='top', fill='x')
    android_label = ttk.Label(android_frame, image=android_qr_img)
    android_label.pack(side='top', fill='x')

    # Frame for iOS QR and badge
    ios_frame = ttk.Frame(main_frame)
    ios_frame.grid(row=0, column=1, sticky='ew', padx=20)
    ios_badge_label = ttk.Label(ios_frame, image=ios_badge_photo)
    ios_badge_label.pack(side='top', fill='x')
    ios_label = ttk.Label(ios_frame, image=ios_qr_img)
    ios_label.pack(side='top', fill='x')

    # Ensure that both columns in the main frame distribute space evenly
    main_frame.columnconfigure(0, weight=1)
    main_frame.columnconfigure(1, weight=1)

    # Keep a reference to the images to avoid garbage collection
    android_label.image = android_qr_img
    ios_label.image = ios_qr_img
    android_badge_label.image = android_badge_photo
    ios_badge_label.image = ios_badge_photo

    # Display system info
    info_frame = ttk.Frame(main_frame)
    info_frame.grid(row=1, column=0, columnspan=2, sticky='ew')
    row = 0
    for key, value in system_info.items():
        ttk.Label(info_frame, text=f"{key}:", font=header_font).grid(row=row, column=0, sticky='e', padx=5, pady=2)
        ttk.Label(info_frame, text=value, font=body_font).grid(row=row, column=1, sticky='w', padx=5, pady=2)
        row += 1

    # Customize the close button
    close_button = ttk.Button(info_frame, text="Close", command=root.destroy)
    close_button.grid(row=row, column=0, columnspan=2, pady=10, sticky='nsew')

    # Center the window on the screen
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
