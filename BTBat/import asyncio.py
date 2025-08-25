import subprocess
import re

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
             "Get-PnpDevice -Class Bluetooth | Where-Object {$_.Status -eq 'OK'} | "
             "Select-Object FriendlyName, InstanceId | Format-List"],
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
        return []

def get_bluetooth_battery_windows( device_id):
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
        exit()