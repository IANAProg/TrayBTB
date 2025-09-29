import tkinter as tk
from tkinter import ttk
import sys
import customtkinter as ctk

class SettingsWindow:
    def __init__(self):
        self._mainWindow = ctk.CTk()
        self._logs_on = ctk.BooleanVar()
        self._log_level = ctk.StringVar()
        self._log_name = ctk.StringVar()
        self._sys_startup = ctk.BooleanVar()
        self._log_levels = {
            "ALL": 0,
            "DEBUG": 10,
            "INFO": 20,
            "WARNING": 30,
            "ERROR": 40,
            "CRITICAL": 50
        }

    def log_block(self, startRow: int):

        logs_frame = ctk.CTkFrame(self._mainWindow, border_width=1)
        logs_frame.grid(row=startRow, column=0, columnspan=4, padx=10, pady=10, sticky="ew")

        # Настройка 1 (текстовое поле)
        LogNamelabel1 = ctk.CTkLabel(logs_frame, text="Название лога")
        LogNamelabel1.grid(row=0, column=0, padx=10, pady=5, sticky="w")
        entry1 = ctk.CTkEntry(logs_frame, textvariable=self._log_name)
        entry1.grid(row=0, column=1, padx=10, pady=5, sticky="ew",columnspan=3)

        # Настройка 2 (выпадающий список)
        LogLevelLabel2 = ctk.CTkLabel(logs_frame, text="Уровень логирования")
        LogLevelLabel2.grid(row=1, column=0, padx=10, pady=5, sticky="w")
        combo2 = ctk.CTkComboBox(logs_frame, values=list(self._log_levels.keys()), 
                            textvariable=self._log_level,state="readonly")
        combo2.grid(row=1, column=1, padx=10, pady=5, sticky="ew",columnspan=3)

        # Настройка 3 (галочка)
        LogEnableLabel3 = ctk.CTkLabel(logs_frame, text="Включить логирование")
        LogEnableLabel3.grid(row=2, column=0, padx=10, pady=5, sticky="w")
        checkbox3 = ctk.CTkCheckBox(logs_frame, variable=self._logs_on)
        checkbox3.grid(row=2, column=1, padx=10, pady=5, sticky="w")

    def sys_block(self):
        LogEnableLabel3 = ctk.CTkLabel(self._mainWindow, text="Запускать при старте системы")
        LogEnableLabel3.grid(row=3, column=0, padx=10, pady=5, sticky="w")
        checkbox3 = ctk.CTkCheckBox(self._mainWindow, variable=self._sys_startup)
        checkbox3.grid(row=3, column=1, padx=10, pady=5, sticky="w")

    def lower_button_block(self, btn_row):
        btn_accept = ctk.CTkButton(self._mainWindow, text="Принять", command=self._on_accept)
        btn_accept.grid(row=btn_row, column=1, padx=10, pady=10, sticky="e")

        btn_cancel = ctk.CTkButton(self._mainWindow, text="Отменить", command=self._on_cancel)
        btn_cancel.grid(row=btn_row, column=2, padx=10, pady=10, sticky="e")

        btn_exit = ctk.CTkButton(self._mainWindow, text="Выход", command=self._on_exit)
        btn_exit.grid(row=btn_row, column=3, padx=10, pady=10, sticky="e")

    def open_settings_window(self):
        self._mainWindow.title("Settings")
        self._mainWindow.geometry("400x300")

        self.log_block(0)

        self.sys_block()

        btn_row_main=4
        self.lower_button_block(btn_row_main)

        self._mainWindow.grid_columnconfigure(1, weight=1)
        self._mainWindow.grid_rowconfigure(0, weight=1)
        self._mainWindow.grid_rowconfigure(1, weight=1)
        self._mainWindow.grid_rowconfigure(2, weight=1)

        self._mainWindow.mainloop()

    def _on_accept(self):
        print(f"log_name: {self._log_name.get()}, log_level: {self._log_level.get()}, logs_on: {self._logs_on.get()}, sys_startup: {self._sys_startup.get()}")

    def _on_cancel(self):
        print("Отменить нажато")

    def _on_exit(self):
        self._mainWindow.destroy()
        sys.exit(0)

    def get_log_level(self):
        return self._log_levels.get(self._log_level.get(), 0)

    def get_logs_on(self):
        return self._logs_on.get()

    def get_sys_startup(self):
        return self._sys_startup.get()

    def get_log_name(self):
        return self._log_name.get()

if __name__ == "__main__":
    settings = SettingsWindow()
    settings.open_settings_window()