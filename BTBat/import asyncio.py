import subprocess
import re
import pystray
from pystray import MenuItem as item
from pystray import Menu
from PIL import Image, ImageDraw
import threading
import time
from winotify import Notification as winNot, audio 
import os

icoPath = os.path.abspath(__file__).replace("import asyncio.py","TrayBTB.png")
devices = []
exitTrap = False
icon = None
choosenDevice = ""
choosenDeviceID = ""
state = 1 # 0 - updating, 1 - no device choosed, 2 - device has chosen


def get_hex_color(val):
    """
    Возвращает hex-код цвета (#RRGGBB)
    """
    normalized = val / 100
    normalized = max(0, min(1, normalized))
    
    red = int(255 * (1 - normalized))
    green = int(255 * normalized)
    
    return f"#{red:02x}{green:02x}00"



def get_updated_menu():
    """Получаем обновленное меню"""
    return (
        item('Обновить список девайсов', update_devices),
        item('Девайсы', makeMenuDevices()),
        item('Выход', exit_app)
    )
def get_connected_menu():
    """Получаем обновленное меню"""
    return (
        item('Обновить список девайсов', update_devices),
        item('Девайсы', makeMenuDevices()),
        item('Отключиться от устройства', disconnect),
        item('Выход', exit_app)
    )

def create_image(color='blue'):
    """Создаем иконку с разными цветами"""
    image = Image.new('RGB', (64, 64), 'white')
    dc = ImageDraw.Draw(image)
    dc.rectangle((16, 16, 48, 48), fill=color)
    return image

def update_icon(new_color):
    """Обновляем иконку"""
    new_image = create_image(new_color)
    icon.icon = new_image
    icon.update_menu()  # Обновляем меню

def update_tooltip(new_tooltip):
    """Обновляем всплывающую подсказку"""
    icon.title = new_tooltip
    icon.update_menu()

def change_name(new_name):
    """Меняем название приложения"""
    icon.name = new_name
    icon.update_menu()

def winNotification(msg):
    notification = winNot(app_id="TrayBTB", title="TrayBTB", msg=msg, icon=icoPath)
    notification.set_audio(audio.Default,False)
    notification.show()
# Функции для меню
def chooseDevice(name,id):
    def handler(icon,item):
        global choosenDevice, choosenDeviceID, state
        choosenDevice = name
        choosenDeviceID = id
        state = 2
        winNotification(f"Connected to {choosenDevice}!")
        icon.menu = get_connected_menu()
        icon.update_menu()
    return handler

def makeMenuDevices():
    """Создание меню устройств"""
    if devices and len(devices) > 0:
        menu_items = [
            item(device.get("name"), chooseDevice(device.get("name"),device.get("id")))
            for device in devices
        ]
    else:
        menu_items = [item("Нет устройств", lambda icon, item: None)]
    return Menu(*menu_items)

def disconnect():
    global choosenDeviceID, choosenDevice,state
    winNotification(f"Disconnected from {choosenDevice}. \nPlease choose device!")
    choosenDevice = ""
    choosenDeviceID = ""
    icon.menu = (
    item('Обновить список девайсов',update_devices),
    item('Девайсы', makeMenuDevices()),
    item('Выход', exit_app)
    )
    icon.update_menu()
    state = 1

def exit_app():
    global exitTrap
    exitTrap = True
    icon.stop()

def get_battery_via_powershell(device_id):

    try:
        # Этот метод работает для ограниченного числа устройств
        result = subprocess.run([
            "powershell", "-Command",
            f"Get-PnpDeviceProperty -InstanceId '{device_id}' -KeyName 'DEVPKEY_Device_BatteryLevel'| select data"
        ], capture_output=True, text=True, encoding='cp866')
        
        if result.returncode == 0:
            match = re.search(r'[0-9]+', result.stdout)
            if match:
                return int(match.group(0))
        
        return None
        
    except Exception as e:
        print(f"Ошибка PowerShell: {e}")
        return None

def get_bluetooth_devices_windows():
    try:
        # Используем PowerShell для получения Bluetooth устройств
        result = subprocess.run(
            ["powershell", "-Command", 
            "Get-PnpDevice -Class Bluetooth | Where-Object Status -eq 'OK' | ForEach-Object {$batteryLevel = (Get-PnpDeviceProperty -InstanceId $_.InstanceId -KeyName 'DEVPKEY_Device_BatteryLevel' -ErrorAction SilentlyContinue).Data "
            "\nif ($null -ne $batteryLevel) {$_ | Select-Object FriendlyName, InstanceId, @{Name='BatteryLevel';Expression={$batteryLevel}}}}| Format-List"
            ],
            capture_output=True, text=True, check=True, encoding='cp866'
        )
        
        devices = []
        lines = result.stdout.split('\n')
        current_device = {}
        
        for line in lines:
            line = line.strip()
            if line.startswith('FriendlyName'):
                if current_device:
                    devices.append(current_device)
                current_device = {'name': line.split(':', 1)[1].strip()}
            elif line.startswith('InstanceId'):
                current_device['id'] = line.split(':', 1)[1].strip()
        if current_device:
            devices.append(current_device)
        return devices
        
    except Exception as e:
        print(f"Ошибка: {e}")
        print(f'\nПодробности: {e.stderr}')
        return []

def get_bluetooth_battery_windows(device_id):
    battery_level = get_battery_via_powershell(device_id)
    if battery_level is not None:
        return battery_level

def update_devices():
    global state
    global choosenDevice
    global devices
    state = 0;    
    devices = get_bluetooth_devices_windows()
    winNotification("Choose device!")
    if choosenDevice != "":
        state = 2
    else:
        state = 1
    icon.menu = get_updated_menu()
    icon.update_menu()
# Создаем меню
menu = (
    item('Обновить список девайсов',update_devices),
    item('Девайсы', makeMenuDevices()),
    item('Выход', exit_app)

)

# Создаем иконку в трее
icon = pystray.Icon("TrayBTB", create_image(), "TrayBTB --Updating devices--", menu)

# Запускаем иконку в отдельном потоке
def run_tray():
    icon.run()

tray_thread = threading.Thread(target=run_tray, daemon=True)
tray_thread.start()

state = 1
winNotification("App started. \nPlease, update your device list and choose device!")
while exitTrap!=True:
    if state == 0:
        update_icon("blue")
        update_tooltip("TrayBTB --Updating devices--")
    if state == 1:
        update_tooltip("TrayBTB --Choose device--")
        update_icon("black")
    if state == 2:
        bat = get_bluetooth_battery_windows(choosenDeviceID)
        update_tooltip(f"TrayBTB --{bat}--")
        update_icon(get_hex_color(bat))
    time.sleep(1)
