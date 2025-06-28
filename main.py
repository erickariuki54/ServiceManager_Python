import customtkinter as ctk
import subprocess
import json
import os
import threading
import time
import pystray
from pystray import MenuItem as item
from PIL import Image, ImageDraw

SERVICES_FILE = os.path.join(os.getenv("LOCALAPPDATA"), "ServiceManager", "services.json")


class ServiceManagerApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Service Manager")
        self.geometry("540x460")
        ctk.set_appearance_mode("system")

        self.services = []
        self.service_widgets = {}

        self.create_widgets()
        self.load_services()
        self.refresh_services()
        self.auto_refresh()
        self.create_tray_icon()

        # Hide instead of close
        self.protocol("WM_DELETE_WINDOW", self.hide_window)

    def create_widgets(self):
        self.entry = ctk.CTkEntry(self, placeholder_text="Enter service name")
        self.entry.pack(pady=10, padx=10, fill="x")

        self.add_btn = ctk.CTkButton(self, text="Add Service", command=self.add_service)
        self.add_btn.pack(pady=5)

        self.services_frame = ctk.CTkScrollableFrame(self)
        self.services_frame.pack(expand=True, fill="both", pady=10, padx=10)

    def load_services(self):
        if os.path.exists(SERVICES_FILE):
            with open(SERVICES_FILE, "r") as f:
                self.services = sorted(json.load(f))
        self.update_service_list()

    def save_services(self):
        os.makedirs(os.path.dirname(SERVICES_FILE), exist_ok=True)
        with open(SERVICES_FILE, "w") as f:
            json.dump(self.services, f)

    def add_service(self):
        name = self.entry.get().strip()
        if name and name not in self.services:
            self.services.append(name)
            self.services.sort()
            self.save_services()
            self.entry.delete(0, 'end')
            self.update_service_list()

    def get_service_status(self, name):
        try:
            result = subprocess.run(["sc", "query", name], capture_output=True, text=True)
            output = result.stdout
            if "RUNNING" in output:
                return "Running"
            elif "STOPPED" in output:
                return "Stopped"
            else:
                return "Unknown"
        except:
            return "Error"

    def control_service(self, name, action):
        try:
            subprocess.run(["sc", action, name], check=True, capture_output=True)
        except subprocess.CalledProcessError:
            # Fix: wrap the full command inside a single string
            cmd = f'sc {action} "{name}"'
            subprocess.run([
                "powershell", "-Command",
                f'Start-Process -FilePath "sc.exe" -ArgumentList \'{action} "{name}"\' -Verb RunAs'
            ])

    def update_service_list(self):
        for widget in self.services_frame.winfo_children():
            widget.destroy()

        self.service_widgets.clear()

        for service in self.services:
            status_text = self.get_service_status(service)
            row = ctk.CTkFrame(self.services_frame)
            row.pack(fill="x", pady=2, padx=5)

            name = ctk.CTkLabel(row, text=service, width=100, anchor="w")
            name.pack(side="left")

            dot_color = "green" if status_text == "Running" else ("red" if status_text == "Stopped" else "gray")
            dot = ctk.CTkLabel(row, text="●", text_color=dot_color, width=10)
            dot.pack(side="left", padx=(0, 4))

            status = ctk.CTkLabel(row, text=status_text, width=80)
            status.pack(side="left", padx=5)

            start_btn = ctk.CTkButton(row, text="Start", width=60,
                                      command=lambda s=service: self.handle_action(s, "start"))
            start_btn.pack(side="left", padx=2)
            start_btn.configure(state="disabled" if status_text == "Running" else "normal")

            stop_btn = ctk.CTkButton(row, text="Stop", width=60,
                                     command=lambda s=service: self.handle_action(s, "stop"))
            stop_btn.pack(side="left", padx=2)

            restart_btn = ctk.CTkButton(row, text="Restart", width=60,
                                        command=lambda s=service: self.handle_action(s, "restart"))
            restart_btn.pack(side="left", padx=2)

            del_btn = ctk.CTkButton(row, text="❌", width=30, fg_color="red",
                                    command=lambda s=service: self.remove_service(s))
            del_btn.pack(side="left")

            self.service_widgets[service] = {
                "status": status,
                "dot": dot,
                "start_btn": start_btn
            }

    def handle_action(self, service, action):
        def run():
            if action == "restart":
                self.control_service(service, "stop")
                time.sleep(1)
                self.control_service(service, "start")
            else:
                self.control_service(service, action)
            time.sleep(1)
            self.refresh_services()
        threading.Thread(target=run, daemon=True).start()

    def remove_service(self, name):
        self.services.remove(name)
        self.save_services()
        self.update_service_list()

    def refresh_services(self):
        def safe_update():
            for service in self.services:
                status = self.get_service_status(service)
                widgets = self.service_widgets.get(service, {})
                if widgets:
                    try:
                        widgets["status"].configure(text=status)
                        color = "green" if status == "Running" else ("red" if status == "Stopped" else "gray")
                        widgets["dot"].configure(text_color=color)
                        widgets["start_btn"].configure(state="disabled" if status == "Running" else "normal")
                    except:
                        pass
        self.after(0, safe_update)

    def auto_refresh(self):
        def loop():
            while True:
                self.refresh_services()
                time.sleep(5)
        threading.Thread(target=loop, daemon=True).start()

    def create_tray_icon(self):
        image = Image.new("RGB", (64, 64), (255, 255, 255))
        draw = ImageDraw.Draw(image)
        draw.ellipse((16, 16, 48, 48), fill="black")

        menu = (
            item("Show", self.show_window),
            item("Exit", self.exit_app)
        )

        self.tray_icon = pystray.Icon("ServiceManager", image, "Service Manager", menu)
        threading.Thread(target=self.tray_icon.run, daemon=True).start()

    def show_window(self, icon=None, item=None):
        self.after(0, self.deiconify)

    def hide_window(self):
        self.withdraw()

    def exit_app(self, icon=None, item=None):
        if hasattr(self, "tray_icon"):
            self.tray_icon.stop()
        self.destroy()


if __name__ == "__main__":
    app = ServiceManagerApp()
    app.mainloop()
