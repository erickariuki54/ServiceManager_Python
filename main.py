import customtkinter as ctk
import subprocess
import json
import os
import threading
import time

SERVICES_FILE = os.path.join(os.getenv("LOCALAPPDATA"), "ServiceManager", "services.json")

class ServiceManagerApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Service Manager")
        self.geometry("500x400")
        ctk.set_appearance_mode("system")

        self.services = []
        self.service_widgets = {}

        self.create_widgets()
        self.load_services()
        self.refresh_services()
        self.auto_refresh()

    def create_widgets(self):
        self.entry = ctk.CTkEntry(self, placeholder_text="Enter service name")
        self.entry.pack(pady=10)

        self.add_btn = ctk.CTkButton(self, text="Add Service", command=self.add_service)
        self.add_btn.pack()

        self.services_frame = ctk.CTkScrollableFrame(self)
        self.services_frame.pack(expand=True, fill="both", pady=10)

    def load_services(self):
        if os.path.exists(SERVICES_FILE):
            with open(SERVICES_FILE, "r") as f:
                self.services = json.load(f)
        self.update_service_list()

    def save_services(self):
        os.makedirs(os.path.dirname(SERVICES_FILE), exist_ok=True)
        with open(SERVICES_FILE, "w") as f:
            json.dump(self.services, f)

    def add_service(self):
        name = self.entry.get().strip()
        if name and name not in self.services:
            self.services.append(name)
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
            # Properly quote the whole argument string passed to sc
            cmd = f'Start-Process -FilePath sc.exe -ArgumentList \'{action} "{name}"\' -Verb RunAs'
            subprocess.run(["powershell", "-Command", cmd])


    def update_service_list(self):
        for widget in self.services_frame.winfo_children():
            widget.destroy()

        for service in self.services:
            row = ctk.CTkFrame(self.services_frame)
            row.pack(fill="x", pady=2, padx=5)

            name = ctk.CTkLabel(row, text=service, width=100, anchor="w")
            name.pack(side="left")

            status = ctk.CTkLabel(row, text=self.get_service_status(service), width=80)
            status.pack(side="left", padx=5)
            self.service_widgets[service] = status

            for action in ["start", "stop", "restart"]:
                btn = ctk.CTkButton(row, text=action.capitalize(), width=60,
                                     command=lambda s=service, a=action: self.handle_action(s, a))
                btn.pack(side="left", padx=2)

            del_btn = ctk.CTkButton(row, text="❌", width=30, fg_color="red",
                                    command=lambda s=service: self.remove_service(s))
            del_btn.pack(side="left")

    def handle_action(self, service, action):
        if action == "restart":
            self.control_service(service, "stop")
            time.sleep(1)
            self.control_service(service, "start")
        else:
            self.control_service(service, action)
        self.refresh_services()

    def remove_service(self, name):
        self.services.remove(name)
        self.save_services()
        self.update_service_list()

    def refresh_services(self):
        for service in self.services:
            status = self.get_service_status(service)
            if service in self.service_widgets:
                self.service_widgets[service].configure(text=status)

    def auto_refresh(self):
        def loop():
            while True:
                self.refresh_services()
                time.sleep(5)

        t = threading.Thread(target=loop, daemon=True)
        t.start()

if __name__ == "__main__":
    app = ServiceManagerApp()
    app.mainloop()
