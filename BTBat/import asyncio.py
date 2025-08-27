import subprocess
import re
import sys
import pystray
from pystray import MenuItem as item
from PIL import Image, ImageDraw
import threading
import time

exitTrap = False
icon = None

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

# Функции для меню

def show_settings():
    pass

def exit_app():
    global exitTrap
    exitTrap = True
    icon.stop()

    

# Создаем меню
menu = (
    item('Девайсы', show_settings),
    item('Выход', exit_app)
)

# Создаем иконку в трее
icon = pystray.Icon("TrayBTB", create_image(), "TrayBTB --Updating devices--", menu)

# Запускаем иконку в отдельном потоке
def run_tray():
    icon.run()

tray_thread = threading.Thread(target=run_tray, daemon=True)
tray_thread.start()

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
    
devices = get_bluetooth_devices_windows()
print("Bluetooth устройства в Windows:")
for i, device in enumerate(devices, 0):
    print(f"{i}. {device.get('name', 'Unknown')}")
while True:
    try:
        ChoosedNum = int(input("Выберите девайс "))
        if ChoosedNum> len(devices)-1 or ChoosedNum<0:
            raise IndexError
        
    except ValueError:
        print("Нужно написать номер, без доп. символов")
    except IndexError:
        print("Впиши номер из списка")
    else:
        device_name = devices[ChoosedNum].get("name")
        device_id = devices[ChoosedNum].get("id")
        battery_level = get_bluetooth_battery_windows(device_id)
        print (battery_level)
        break

update_tooltip("TrayBTB --Choose device--")
update_icon("black")
while exitTrap!=True:
    pass
