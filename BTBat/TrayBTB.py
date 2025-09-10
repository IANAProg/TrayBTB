import asyncio
import logging
import subprocess
import re
import pystray
import os
import threading
import sys
import shutil
import win32com.client
import concurrent.futures
import time
import pythoncom
from pystray import MenuItem as item
from PIL import Image, ImageDraw
from winotify import Notification as WinNotification, audio
from enum import Enum
from typing import List, Dict, Optional
from datetime import datetime
from dataclasses import dataclass

NO_WINDOW = 0
if sys.platform == "win32" and hasattr(subprocess, "CREATE_NO_WINDOW"):
    NO_WINDOW = subprocess.CREATE_NO_WINDOW

fulltime = str(datetime.now())
if sys.platform == "win32":
    try:
        asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
    except Exception:
        pass

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
        self.device_type: str = ""  # 'ble' или 'pnp'
        
    def _read_pnp_battery(self, instance_id: str) -> Optional[int]:
        try:
            result = subprocess.run(
                [
                    "powershell", "-NoProfile", "-NonInteractive", "-Command",
                    f"(Get-PnpDeviceProperty -InstanceId '{instance_id}' -KeyName 'DEVPKEY_Device_BatteryLevel' -ErrorAction SilentlyContinue).Data"
                ],
                capture_output=True, text=True, encoding='cp866', creationflags=NO_WINDOW
            )
            if result.returncode == 0 and result.stdout:
                match = re.search(r'[0-9]+', result.stdout)
                if match:
                    return int(match.group(0))
            return None
        except Exception as e:
            self._log_error(e)
            return None


    def get_battery_level(self) -> Optional[int]:
        if not self.device_id or not self.device_type:
            return None
        if self.device_type == "ble":
            log_handler.log.error(f"somewhere got ble device, check it {self.device_id} {self.device_type}")
        else:
            return self._read_pnp_battery(self.device_id)

    def _log_error(self, e: Exception):
        # короткая локальная функция логирования (использует существующий логгер, если есть)
        try:
            log_handler.log.error(f"Error in BT module: {e}")
        except Exception:
            print(f"BatteryMonitor error: {e}")

class DeviceManager:
    def get_devices(self) -> List[Dict[str, str]]:
        """
        Быстрый WMI-скан для кандидатов (фильтр по имени) + параллельный
        вызов Get-PnpDeviceProperty для каждого InstanceId.
        Возвращает список dict: {name, id, id_type='pnp', battery}
        """
        devices: List[Dict[str, str]] = []
        initialized_com = False

        # Инициализируем COM в текущем потоке (безопасно — логируем любые ошибки)
        try:
            try:
                pythoncom.CoInitialize()
                initialized_com = True
                try:
                    log_handler.log.debug("pythoncom.CoInitialize() succeeded")
                except Exception:
                    pass
            except Exception as e:
                try:
                    log_handler.log.warning(f"pythoncom.CoInitialize() failed or already initialized: {e}")
                except Exception:
                    pass

            # --- основной код метода (WMI + параллельный PowerShell) ---
            try:
                # подключаемся к WMI и получаем PNPDeviceID + Name только для кандидатов
                locator = win32com.client.Dispatch("WbemScripting.SWbemLocator")
                svc = locator.ConnectServer(".", "root\\cimv2")
                # фильтр по имени уменьшает количество проверяемых устройств
                q = (
                    "SELECT PNPDeviceID, Name FROM Win32_PnPEntity "
                    "WHERE Status='OK' AND ("
                    "Name LIKE '%Headphone%' OR Name LIKE '%Headphones%' OR "
                    "Name LIKE '%Audio%' OR Name LIKE '%Hands-Free%' OR "
                    "Name LIKE '%AirPods%' OR Name LIKE '%WH%' OR Name LIKE '%BT%'"
                    ")"
                )
                items = svc.ExecQuery(q)
                candidates = []
                for it in items:
                    inst = getattr(it, "PNPDeviceID", None)
                    name = getattr(it, "Name", None)
                    if inst:
                        candidates.append((inst, name or inst))
                try:
                    log_handler.log.info(f"WMI candidates count: {len(candidates)}")
                except Exception:
                    pass
            except Exception as e:
                try:
                    log_handler.log.error(f"WMI query failed: {e}")
                except Exception:
                    print(f"WMI query failed: {e}")
                candidates = []

            if not candidates:
                return []

            # runner для PowerShell: пробуем pwsh, иначе powershell
            runner = shutil.which("pwsh") or "powershell"

            def probe(instance_and_name):
                inst, name = instance_and_name
                try:
                    # вызываем только для одного InstanceId; возвращаем int battery или None
                    ps_cmd = f"(Get-PnpDeviceProperty -InstanceId '{inst}' -KeyName 'DEVPKEY_Device_BatteryLevel' -ErrorAction SilentlyContinue).Data"
                    res = subprocess.run(
                        [runner, "-NoProfile", "-NonInteractive", "-Command", ps_cmd],
                        capture_output=True, text=True, encoding='cp866', timeout=8, creationflags=NO_WINDOW
                    )
                    out = (res.stdout or "").strip()
                    if not out:
                        return None
                    m = re.search(r'\d+', out)
                    if not m:
                        return None
                    return {"name": name, "id": inst, "id_type": "pnp", "battery": int(m.group(0))}
                except subprocess.TimeoutExpired:
                    try:
                        log_handler.log.debug(f"PS timeout for {inst}")
                    except Exception:
                        pass
                    return None
                except Exception as e:
                    try:
                        log_handler.log.debug(f"PS error for {inst}: {e}")
                    except Exception:
                        pass
                    return None

            # параллельно опрашиваем кандидатов (ограничиваем worker-ы)
            results = []
            with concurrent.futures.ThreadPoolExecutor(max_workers=12) as ex:
                futs = [ex.submit(probe, c) for c in candidates]
                for f in concurrent.futures.as_completed(futs):
                    try:
                        r = f.result()
                        if r:
                            results.append(r)
                    except Exception:
                        pass

            # убираем дубли по instance id
            seen = set()
            for r in results:
                key = r.get("id")
                if key and key not in seen:
                    seen.add(key)
                    devices.append(r)

            try:
                log_handler.log.info(f"DeviceManager (pywin32) found: {len(devices)}")
            except Exception:
                pass

            return devices

        finally:
            # Обязательно разинициализируем COM в текущем потоке (если инициализировали)
            if initialized_com:
                try:
                    pythoncom.CoUninitialize()
                    try:
                        log_handler.log.debug("pythoncom.CoUninitialize() called")
                    except Exception:
                        pass
                except Exception as e:
                    try:
                        log_handler.log.warning(f"pythoncom.CoUninitialize() failed: {e}")
                    except Exception:
                        pass

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
        self.battery_status = BatStatus(level=None,last_update=0)

        self.icon = None

        # lock для сериализации изменений меню/иконки между потоками (pystray GUI и background)
        self._menu_lock = threading.Lock()
        # простая защита: принять изменения меню только если прошло >= minimal_menu_update_s сек
        self._last_menu_update_ts = 0.0
        self._minimal_menu_update_s = 0.1

        self.setup_tray()

    def safe_set_menu(self, menu: pystray.Menu):
        """Set menu and call update_menu under a lock to avoid concurrent UI races."""
        try:
            with self._menu_lock:
                now = time.time()
                if now - self._last_menu_update_ts < self._minimal_menu_update_s:
                    # пропускаем слишком частые обновления меню
                    return
                self._last_menu_update_ts = now
                self.icon.menu = menu
                # update_menu должен вызываться под lock
                self.icon.update_menu()
        except Exception as e:
            try:
                log_handler.log.debug(f"safe_set_menu failed: {e}")
            except Exception:
                pass

    def setup_tray(self):
        # используем get_updated_menu(), оно возвращает pystray.Menu
        menu = self.get_updated_menu()
        self.icon = pystray.Icon(
            "TrayBTB",
            self.icon_manager.create_image(),
            "TrayBTB --Updating devices--",
            menu
        )

    def run(self):
        tray_thread = threading.Thread(target=self.icon.run, daemon=True)
        tray_thread.start()

        log_handler.log.info("App started")
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
            log_handler.log.error(f"Main loop error: {e}")
            self.notification_manager.show_notification(
                "Error in main loop. Application will be closed."
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
                    log_handler.log.warning(f"Unknown state: {self.state}")
                    
        except Exception as e:
            log_handler.log.error(f"State handling error: {e}")
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
            log_handler.log.warning("Failed to get battery level")

    async def auto_disconnect(self):
        """Automatically disconnect device after too many errors."""
        self.log_handler.log.info("Auto-disconnecting due to errors")
        self.disconnect_device(None)
        self.battery_status.error_count = 0

    def update_icon(self, new_color: str):
        try:
            # изменение иконки — НЕ вызывать update_menu (это вызывает пересборку нативного меню)
            with self._menu_lock:
                self.icon.icon = self.icon_manager.create_image(new_color)
                # НЕ: self.icon.update_menu()
        except Exception:
            pass

    def update_tooltip(self, new_tooltip: str):
        try:
            # изменение tooltip — не трогаем меню
            with self._menu_lock:
                self.icon.title = new_tooltip
                # НЕ: self.icon.update_menu()
        except Exception:
            pass

    def update_devices(self, icon=None, item=None):
         log_handler.log.info("Scheduling background update for devices")
         if getattr(self, "_updating_thread", None) and self._updating_thread.is_alive():
             return
         self._updating_thread = threading.Thread(target=self._bg_update_devices, daemon=True)
         self._updating_thread.start()

    def _bg_update_devices(self):
        try:
            log_handler.log.info("Background: updating devices")
            self.state = DeviceState.UPDATING
            devices = self.device_manager.get_devices()
            self.devices = devices
            log_handler.log.info("Background: ended seeking for devices")
            self.notification_manager.show_notification("Choose device!")
            self.state = DeviceState.DEVICE_CHOSEN if self.chosen_device else DeviceState.NO_DEVICE
            # обновляем иконку и меню через безопасный wrapper
            if self.state == DeviceState.NO_DEVICE:
                with self._menu_lock:
                    self.icon.icon = self.icon_manager.create_image("black")
                    #self.icon.update_menu()
            else:
                level = self.battery_monitor.get_battery_level() or 0
                with self._menu_lock:
                    self.icon.icon = self.icon_manager.create_image(self.icon_manager.get_hex_color(level))
                    #self.icon.update_menu()
            # обновляем меню (через safe_set_menu чтобы избежать гонок)
            self.safe_set_menu(self.get_updated_menu())
        except Exception as e:
            try:
                log_handler.log.error(f"Background update failed: {e}")
            except Exception:
                pass
            # восстановим меню
            self.safe_set_menu(self.get_updated_menu())

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
                    self.choose_device(device.get("name"), device.get("id"), device.get("id_type"))
                )
                for device in self.devices
            ]
        else:
            menu_items = [item("No devices", lambda icon, item: None)]
            
        return pystray.Menu(*menu_items)

    def choose_device(self, name: str, device_id: str, device_type: str):
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
            self.battery_monitor.device_type = device_type
            self.state = DeviceState.DEVICE_CHOSEN
            
            log_handler.log.info(f"State: Device chosen. It's {self.chosen_device}")
            self.notification_manager.show_notification(
                f"Connected to {self.chosen_device}!"
            )
            
            # Update menu to include disconnect option (через сериализованный safe_set_menu)
            try:
                self.safe_set_menu(self.get_connected_menu())
            except Exception:
                # backup: попробовать обновить прямо, если safe_set_menu упал
                try:
                    with self._menu_lock:
                        self.icon.menu = self.get_connected_menu()
                        self.icon.update_menu()
                except Exception:
                    pass
            
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

    def disconnect_device(self, icon=None, item=None):
        """Handles disconnecting from current device"""
        log_handler.log.info(f"disconnect from {self.chosen_device}")
        self.notification_manager.show_notification(
            f"Disconnected from {self.chosen_device}. \nPlease choose device!"
        )
        self.chosen_device = ""
        self.chosen_device_id = ""
        self.battery_monitor.device_id = ""
        self.battery_monitor.device_type = ""
        self.state = DeviceState.NO_DEVICE
        self.icon_manager.create_image("Black")
        # Reset menu to original state
        try:
            self.safe_set_menu(self.get_updated_menu())
        except Exception:
            # backup: попробовать обновить прямо, если safe_set_menu упал
            try:
                with self._menu_lock:
                    self.icon.menu = self.get_updated_menu()
                    self.icon.update_menu()
            except Exception:
                pass
    
    def exit_app(self, icon=None, item=None):
        self.exit_flag = True
        try:
            if self.icon:
                self.icon.stop()
        except Exception:
            pass
        log_handler.log.info("App exiting")
        # завершение процесса
        sys.exit(0)
 
    def get_updated_menu(self) -> pystray.Menu:
        return pystray.Menu(
            item('Обновить список девайсов', self.update_devices),
            item('Девайсы', self.make_menu_devices()),
            item('Выход', self.exit_app)
        )
    
    def get_connected_menu(self) -> pystray.Menu:
        return pystray.Menu(
            item('Обновить список девайсов', self.update_devices),
            item('Девайсы', self.make_menu_devices()),
            item('Отключиться от устройства', self.disconnect_device),
            item('Выход', self.exit_app)
        )

log_handler = Logs()
def main():
    app = TrayApplication()
    app.run()

if __name__ == "__main__":
    main()