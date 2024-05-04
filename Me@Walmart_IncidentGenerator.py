# The `import sys` statement in Python is used to import the `sys` module. The `sys` module provides
# access to some variables used or maintained by the Python interpreter and to functions that interact
# strongly with the interpreter. It allows you to access command-line arguments, interact with the
# Python runtime environment, and perform system-specific operations.
import sys
# The `import subprocess` statement in Python is used to import the `subprocess` module. The
# `subprocess` module allows you to spawn new processes, connect to their input/output/error pipes,
# and obtain their return codes. It provides a way to execute system commands, run other programs, and
# communicate with them from within your Python script.
import subprocess
# The `import socket` statement in Python is used to import the `socket` module. The `socket` module
# provides access to the BSD socket interface, which allows your Python programs to establish network
# connections, send and receive data over the network, and perform various network-related operations.
# It enables you to create network sockets, work with IP addresses and ports, and implement networking
# protocols in your Python code.
import socket
# The statement `import tkinter as tk` in Python is importing the `tkinter` module and aliasing it as
# `tk`. This allows you to refer to the `tkinter` module using the shorter alias `tk` throughout your
# code. This is a common practice to make the code more concise and readable, especially when working
# with modules that have long names.
# The statement `import tkinter as tk` in Python is importing the `tkinter` module and aliasing it as
# `tk`. This allows you to refer to the `tkinter` module using the shorter alias `tk` throughout your
# code. This is a common practice to make the code more concise and readable, especially when working
# with modules that have long names.
import tkinter as tk
# The statement `from tkinter import ttk` in Python is specifically importing the `ttk` module from
# the `tkinter` package. This allows you to directly access the classes and functions provided by the
# `ttk` module without having to prefix them with `ttk.` every time you use them in your code. The
# `ttk` module in `tkinter` provides themed widget classes that offer a more modern and visually
# appealing look compared to the standard `tkinter` widgets.
from tkinter import ttk
# The statement `from PIL import Image, ImageTk` in the Python code is importing the `Image` and
# `ImageTk` classes specifically from the `PIL` package. This allows direct access to these classes
# without having to prefix them with `PIL.` every time they are used in the code. The `Image` class is
# used for handling images in the Python Imaging Library (PIL), which is now known as Pillow. The
# `ImageTk` class is used for converting images from PIL format to a format that can be displayed in a
# Tkinter GUI window.
from PIL import Image, ImageTk
# The `import qrcode` statement in the Python code is importing the `qrcode` module. This module
# allows you to generate QR codes in Python. By using this module, you can create QR codes containing
# various types of data such as URLs, text, contact information, etc. Additionally, the `qrcode`
# module provides customization options for the generated QR codes, such as setting error correction
# levels, box sizes, borders, and more.
import qrcode
# The statement `from datetime import datetime` in the Python code is specifically importing the
# `datetime` class from the `datetime` module. This allows direct access to the `datetime` class
# without having to prefix it with `datetime.` every time it is used in the code. The `datetime` class
# in Python's `datetime` module is used for manipulating dates and times in various formats and
# performing operations related to date and time calculations.
from datetime import datetime
# The `import wmi` statement in the Python code is importing the `wmi` module. The `wmi` module is
# used for interacting with the Windows Management Instrumentation (WMI) infrastructure on Windows
# systems. It provides a way to access and manage system information, configuration settings, and
# resources on a Windows machine programmatically. By using the `wmi` module, you can query various
# aspects of the system, such as hardware information, operating system details, network
# configuration, and more, making it a powerful tool for system administration and monitoring tasks in
# Python scripts.
import wmi
# The `import win32com.client` statement in the Python code is importing the `win32com.client` module.
# This module is part of the `pywin32` library, which provides Python bindings for the Microsoft
# Windows API.
import win32com.client
# The `import winreg` statement in the Python code is importing the `winreg` module. The `winreg`
# module provides access to the Windows Registry API on Windows systems. It allows Python scripts to
# interact with the Windows Registry, read and write registry keys and values, query registry
# information, and perform operations related to system configuration and settings stored in the
# Windows Registry.
import winreg

# Helper function to ensure required packages are installed
def ensure_package(package_name, import_name=None):
    """
    The function `ensure_package` ensures that a specified package is imported in Python, installing it
    if necessary.
    
    :param package_name: The `package_name` parameter is a string that represents the name of the
    package that you want to ensure is installed. This function will attempt to import the package, and
    if it is not found, it will use `pip` to install the package before attempting to import it again
    :param import_name: The `import_name` parameter in the `ensure_package` function is used to specify
    the name under which the package should be imported. If `import_name` is not provided, it defaults
    to the same value as `package_name`. This allows you to import a package using a different name than
    the
    :return: The `ensure_package` function returns the imported module corresponding to the
    `import_name` parameter. If the module cannot be imported due to an `ImportError`, the function
    attempts to install the package using `pip` and then imports the module again before returning it.
    """
    if import_name is None:
        import_name = package_name
    try:
        return __import__(import_name)
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", package_name])
        return __import__(import_name)

# Import required packages using the helper function
wmi = ensure_package('WMI')
Image, ImageTk = ensure_package('Pillow', 'PIL.Image').Image, ensure_package('PIL.ImageTk')
qrcode = ensure_package('qrcode[pil]')

# The `# Constants for QR Code generation` section in the Python code is defining a dictionary named
# `QR_CONFIG` that contains configuration settings for generating QR codes using the `qrcode` module.
# These settings include:
QR_CONFIG = {
    'version': 7,
    'error_correction': qrcode.constants.ERROR_CORRECT_L,
    'box_size': 10,
    'border': 2,
}

def get_registry_value(path, name):
    """
    The function `get_registry_value` retrieves a value from the Windows registry given a specific path
    and name, handling exceptions for file not found and unexpected errors.
    
    :param path: The `path` parameter in the `get_registry_value` function is a string that represents
    the registry key path where the value is located. It typically follows a format like
    "SOFTWARE\Microsoft\Windows\CurrentVersion" in the Windows registry
    :param name: The `name` parameter in the `get_registry_value` function refers to the name of the
    registry value you want to retrieve from the specified registry path. It is used to identify the
    specific value within the registry key that you are interested in accessing
    :return: The `get_registry_value` function returns the value of the specified registry key `name`
    located at the given `path`. If the key is not found, it prints a message indicating that the key
    was not found and returns `None`. If an unexpected error occurs during the registry access, it
    prints an error message and also returns `None`.
    """
    try:
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path)
        value, regtype = winreg.QueryValueEx(key, name)
        winreg.CloseKey(key)
        return value
    except FileNotFoundError:
        print(f"Registry key {path}\\{name} not found on this system.")
        return None  # Optionally, return a default value if needed
    except Exception as e:
        print(f"Unexpected error accessing registry {path}\\{name}: {str(e)}")
        return None



def get_latest_update():
    """
    The function `get_latest_update` retrieves the title of the latest Cumulative Update from the
    Windows Update history using the Microsoft Update API.
    :return: The function `get_latest_update` will return the title of the latest Cumulative Update if
    found in the update history. If no Cumulative Update is found, it will return "No Updates Found". If
    there is an error fetching the update history, it will return "Error Fetching Updates".
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
    The `fetch_system_info` function retrieves various system information such as hostname, IP address,
    manufacturer, model, serial number, disk size, RAM size, OS details, installation date, last reboot
    time, role, and latest update.
    :return: The function `fetch_system_info` returns a dictionary containing various system information
    such as hostname, IP address, manufacturer, model, serial number, disk size, RAM size, operating
    system details, installation date, last reboot time, role, and latest update information.
    """
    c = wmi.WMI()
    system, os_info, bios, disk = c.Win32_ComputerSystem()[0], c.Win32_OperatingSystem()[0], c.Win32_BIOS()[0], c.Win32_DiskDrive()[0]
    hostname = socket.gethostname()
    ip_address = socket.gethostbyname(hostname)
    last_reboot = datetime.strptime(os_info.LastBootUpTime.split('.')[0], "%Y%m%d%H%M%S")
    installed_on = datetime.strptime(os_info.InstallDate.split('.')[0], "%Y%m%d%H%M%S")
    role = get_registry_value(r"SOFTWARE\Workstation\Build", "Role")
    if role is None:
        role = "UNKNOWN"  # Provide a default value or handle accordingly
    latest_update = get_latest_update()
    return {
        'hostname': hostname,
        'ip_address': ip_address,
        'manufacturer': system.Manufacturer,
        'model': system.Model,
        'serial_number': bios.SerialNumber,
        'disk_size_gb': int(disk.Size) / (1024**3),
        'ram_gb': int(system.TotalPhysicalMemory) / (1024**3),
        'os': os_info.Caption,
        'installed_on': installed_on.strftime("%m/%d/%Y %H:%M"),
        'last_reboot': last_reboot.strftime("%m/%d/%Y %H:%M"),
        'role': role,
        'latest_update': latest_update
    }


def generate_qr_code(url):
    """
    The function `generate_qr_code` creates a QR code image for a given URL with specified configuration
    settings.
    
    :param url: The `generate_qr_code` function takes a URL as input and generates a QR code image for
    that URL. The function uses the `qrcode` library to create the QR code
    :return: The function `generate_qr_code` returns a QR code image generated from the provided URL.
    The image is resized to 200x200 pixels using Lanczos resampling and has a black fill color on a
    white background.
    """
    qr = qrcode.QRCode(**QR_CONFIG)
    qr.add_data(url)
    qr.make(fit=True)
    return qr.make_image(fill='black', back_color='white').resize((200, 200), Image.Resampling.LANCZOS)

def create_gui(android_qr_img, ios_qr_img, system_info):
    """
    The function `create_gui` generates a GUI window displaying Android and iOS QR codes along with
    system information.
    
    :param android_qr_img: It seems like your message got cut off. Could you please provide more
    information about the `android_qr_img` parameter so that I can assist you further with the
    `create_gui` function?
    :param ios_qr_img: It seems like your message got cut off. Could you please provide more information
    about the ios_qr_img parameter so that I can assist you further?
    :param system_info: System_info is a dictionary containing information about the system. It could
    include details such as the operating system, version, device model, and any other relevant system
    information that you want to display in the GUI
    """
    root = tk.Tk()
    root.title("System Info and QR Code Viewer")
    ttk.Label(root, text="Android App QR:").grid(row=0, column=0)
    ttk.Label(root, text="iOS App QR:").grid(row=0, column=1)
    android_label = ttk.Label(root, image=android_qr_img)
    ios_label = ttk.Label(root, image=ios_qr_img)
    android_label.grid(row=1, column=0)
    ios_label.grid(row=1, column=1)
    row = 2
    for key, value in system_info.items():
        ttk.Label(root, text=f"{key}: {value}").grid(row=row, column=0, columnspan=2)
        row += 1
    ttk.Button(root, text="Close", command=root.destroy).grid(row=row, column=0, columnspan=2)
    root.mainloop()

# The above code is a common Python idiom that checks if the script is being run as the main program
# or if it is being imported as a module into another script. If the script is being run as the main
# program, the code block following this check will be executed. This is often used to define behavior
# that should only occur when the script is run directly, and not when it is imported as a module.
if __name__ == "__main__":
    android_url = "https://play.google.com/store/apps/details?id=com.example"
    ios_url = "https://apps.apple.com/us/app/example"
    android_qr = generate_qr_code(android_url)
    ios_qr = generate_qr_code(ios_url)
    system_info = fetch_system_info()
    android_qr_img = ImageTk.PhotoImage(android_qr)
    ios_qr_img = ImageTk.PhotoImage(ios_qr)
    create_gui(android_qr_img, ios_qr_img, system_info)
