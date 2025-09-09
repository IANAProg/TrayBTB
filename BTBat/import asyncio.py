import asyncio
import logging
import subprocess
import re
import pystray
import os
import threading
from pystray import MenuItem as item
from PIL import Image, ImageDraw
from winotify import Notification as WinNotification, audio
from enum import Enum
from typing import List, Dict, Optional
from datetime import datetime
from dataclasses import dataclass

fulltime = str(datetime.now())

@dataclass
class BatStatus:
    level: Optional[int]
    last_update: float
    error_count: int = 0

class Logs:
    def __init__(self):
        self.log = logging.getLogger("TrayBTB")
        if not os.path.exists("logs"):
            os.makedirs("logs")
        logFileStart = "logs\\TrayBTB_"
        logFileExt = ".log"
        self.logFileName = logFileStart+fulltime[:fulltime.rfind('.')].replace(" ","_").replace(":","-")+logFileExt
        self.logLevel = logging.INFO
        self.log.setLevel(self.logLevel)
        self.handler = logging.FileHandler(filename=self.logFileName, mode = "a", encoding='utf-8')
        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        )
        self.handler.setFormatter(formatter)
        self.log.addHandler(self.handler)
        self.log.info("Logging initialized")

    def changeLogLevel(self, level: int):
        self.log.setLevel(level)
    
class DeviceState(Enum):
    UPDATING = 0
    NO_DEVICE = 1
    DEVICE_CHOSEN = 2

class BatteryMonitor:
    def __init__(self):
        self.device_id: str = ""
        
    def get_battery_level(self) -> Optional[int]:
        try:
            result = subprocess.run([
                "powershell", "-Command",
                f"Get-PnpDeviceProperty -InstanceId '{self.device_id}' -KeyName 'DEVPKEY_Device_BatteryLevel'| select data"
            ], capture_output=True, text=True, encoding='cp866')
            
            if result.returncode == 0:
                match = re.search(r'[0-9]+', result.stdout)
                if match:
                    return int(match.group(0))
            return None
        except Exception as e:
            print(f"PowerShell error: {e}")
            return None

class DeviceManager:
    def get_devices(self) -> List[Dict[str, str]]:
        try:
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
            print(f"Error: {e}")
            return []

class IconManager:
    def __init__(self, size: tuple = (64, 64)):
        self.size = size
        
    def create_image(self, color: str = 'blue') -> Image:
        image = Image.new('RGB', self.size, 'white')
        dc = ImageDraw.Draw(image)
        dc.rectangle((16, 16, 48, 48), fill=color)
        return image    
    @staticmethod
    def get_hex_color(val: float) -> str:
        normalized = max(0, min(1, val / 100))
        red = int(255 * (1 - normalized))
        green = int(255 * normalized)
        return f"#{red:02x}{green:02x}00"

class NotificationManager:
    def __init__(self, app_name: str, icon_path: str):
        self.app_name = app_name
        self.icon_path = icon_path
        
    def show_notification(self, message: str):
        notification = WinNotification(
            app_id=self.app_name,
            title=self.app_name,
            msg=message,
            icon=self.icon_path
        )
        notification.set_audio(audio.Reminder, False)
        notification.show()

class TrayApplication:
    def __init__(self):
        self.state = DeviceState.NO_DEVICE
        self.exit_flag = False
        self.chosen_device = ""
        self.chosen_device_id = ""
        self.devices = []
        
        self.icon_manager = IconManager()
        self.battery_monitor = BatteryMonitor()
        self.device_manager = DeviceManager()
        self.notification_manager = NotificationManager(
            "TrayBTB",
            os.path.join(os.path.dirname(__file__), "TrayBTB.png")
        )
        self.update_interval = 1.0
        self.error_threshold = 10
        self.log_handler = Logs()
        self.battery_status = BatStatus(level=None,last_update=0)

        self.icon = None
        self.setup_tray()

    def setup_tray(self):
        menu = (
            item('Обновить список девайсов', self.update_devices),
            item('Девайсы', self.make_menu_devices()),
            item('Выход', self.exit_app)
        )
        self.icon = pystray.Icon(
            "TrayBTB",
            self.icon_manager.create_image(),
            "TrayBTB --Updating devices--",
            menu
        )

    def run(self):
        tray_thread = threading.Thread(target=self.icon.run, daemon=True)
        tray_thread.start()

        self.log_handler.log.info("App started")
        self.notification_manager.show_notification(
            "App started. \nPlease, update your device list and choose device!"
        )
        
        asyncio.run(self.main_loop())

    async def main_loop(self):
        try:
            while not self.exit_flag:
                await self.handle_state()
                await asyncio.sleep(self.update_interval)
                
        except Exception as e:
            self.log_handler.log.error(f"Main loop error: {e}")
            self.notification_manager.show_notification(
                "Error in main loop. Application will restart."
            )
            self.exit_app()

    async def handle_state(self):
        """Handle different application states and update UI accordingly."""
        try:
            match self.state:
                case DeviceState.UPDATING:
                    await self.handle_updating_state()
                case DeviceState.NO_DEVICE:
                    await self.handle_no_device_state()
                case DeviceState.DEVICE_CHOSEN:
                    await self.handle_device_chosen_state()
                case _:
                    self.log_handler.log.warning(f"Unknown state: {self.state}")
                    
        except Exception as e:
            self.log_handler.log.error(f"State handling error: {e}")
            self.battery_status.error_count += 1
            
            if self.battery_status.error_count >= self.error_threshold:
                self.notification_manager.show_notification(
                    "Multiple errors occurred. Please check device connection."
                )
                await self.auto_disconnect()

    async def handle_updating_state(self):
        """Handle the updating state UI."""
        self.update_icon("blue")
        self.update_tooltip("TrayBTB --Updating devices--")

    async def handle_no_device_state(self):
        """Handle the no device state UI."""
        self.update_tooltip("TrayBTB --Choose device--")
        self.update_icon("black")

    async def handle_device_chosen_state(self):
        """Handle the device chosen state, including battery monitoring."""
        bat_level = self.battery_monitor.get_battery_level()
        
        if bat_level is not None:
            self.battery_status.level = bat_level
            self.battery_status.error_count = 0
            self.update_tooltip(f"TrayBTB --{bat_level}%--")
            self.update_icon(self.icon_manager.get_hex_color(bat_level))
            
            # Alert on low battery
            if bat_level <= 20:
                self.notification_manager.show_notification(
                    f"Низкий заряд батареи: {bat_level}%"
                )
        else:
            self.battery_status.error_count += 1
            self.log_handler.log.warning("Failed to get battery level")

    async def auto_disconnect(self):
        """Automatically disconnect device after too many errors."""
        self.log_handler.log.info("Auto-disconnecting due to errors")
        self.disconnect_device(None)
        self.battery_status.error_count = 0

    def update_icon(self, new_color: str):
        self.icon.icon = self.icon_manager.create_image(new_color)
        self.icon.update_menu()

    def update_tooltip(self, new_tooltip: str):
        self.icon.title = new_tooltip
        self.icon.update_menu()

    def update_devices(self):
        self.log_handler.log.info("State: Updating devices")
        self.state = DeviceState.UPDATING
        self.devices = self.device_manager.get_devices()
        self.log_handler.log.info("Ended seeking for devices")
        self.notification_manager.show_notification("Choose device!")
        self.state = (DeviceState.DEVICE_CHOSEN 
                     if self.chosen_device 
                     else DeviceState.NO_DEVICE)
        self.icon_manager.create_image("Black" if self.state == DeviceState.NO_DEVICE else self.icon_manager.get_hex_color({self.battery_monitor.get_battery_level}))
        self.icon.menu = self.get_updated_menu()
        self.icon.update_menu()

    def make_menu_devices(self) -> pystray.Menu:
        """
        Creates a menu of available Bluetooth devices.
        
        Returns:
            pystray.Menu: Menu containing device options
        """
        if self.devices and len(self.devices) > 0:
            menu_items = [
                item(
                    device.get("name"), 
                    self.choose_device(device.get("name"), device.get("id"))
                )
                for device in self.devices
            ]
        else:
            menu_items = [item("No devices", lambda icon, item: None)]
            
        return pystray.Menu(*menu_items)

    def choose_device(self, name: str, device_id: str):
        """
        Creates a handler function for device selection.
        
        Args:
            name: Device friendly name
            device_id: Device instance ID
        
        Returns:
            Callable: Handler function for menu item
        """
        def handler(icon: pystray.Icon, item: item):
            self.chosen_device = name
            self.chosen_device_id = device_id
            self.battery_monitor.device_id = device_id
            self.state = DeviceState.DEVICE_CHOSEN
            
            self.log_handler.log.info(f"State: Device chosen. It's {self.chosen_device}")
            self.notification_manager.show_notification(
                f"Connected to {self.chosen_device}!"
            )
            
            # Update menu to include disconnect option
            self.icon.menu = self.get_connected_menu()
            self.icon.update_menu()
            
        return handler

    def get_connected_menu(self) -> pystray.Menu:
        """
        Creates menu for when device is connected.
        
        Returns:
            pystray.Menu: Updated menu with disconnect option
        """
        return pystray.Menu(
            item('Update Devices', self.update_devices),
            item('Devices', self.make_menu_devices()),
            item('Disconnect Device', self.disconnect_device),
            item('Exit', self.exit_app)
        )

    def disconnect_device(self, item: item):
        """Handles disconnecting from current device"""
        self.log_handler.log.info(f"disconnect from {self.chosen_device}")
        self.notification_manager.show_notification(
            f"Disconnected from {self.chosen_device}. \nPlease choose device!"
        )
        self.chosen_device = ""
        self.chosen_device_id = ""
        self.battery_monitor.device_id = ""
        self.state = DeviceState.NO_DEVICE
        self.icon_manager.create_image("Black")
        # Reset menu to original state
        self.icon.menu = self.get_updated_menu()
        self.icon.update_menu()
    
    def exit_app(self):
        self.exit_flag = True
        self.log_handler.log.info("Exiting")
        self.icon.stop()

    def get_updated_menu(self):
        return (
            item('Обновить список девайсов', self.update_devices),
            item('Девайсы', self.make_menu_devices()),
            item('Выход', self.exit_app)
        )
    def get_connected_menu(self):
        return (
            item('Обновить список девайсов', self.update_devices),
            item('Девайсы', self.make_menu_devices()),
            item('Отключиться от устройства', self.disconnect_device),
            item('Выход', self.exit_app)
        )
        

def main():
    app = TrayApplication()
    app.run()

if __name__ == "__main__":
    main()
