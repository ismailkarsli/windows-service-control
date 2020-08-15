import win32serviceutil
import time
import tkinter as tk
import json


def services(action, serviceName):
    if action == 'start':
        win32serviceutil.StartService(serviceName)
    elif action == 'stop':
        win32serviceutil.StopService(serviceName)
    elif action == 'restart':
        win32serviceutil.RestartService(serviceName)
    elif action == 'status':
        try:
            if win32serviceutil.QueryServiceStatus(serviceName)[1] == 4:
                return "Running"
            else:
                return "Not running"
        except:
            return "Not exists"

    time.sleep(1)
    servicesFrame.grid_forget()
    printServices()


class Settings():
    def getSettings(self):
        try:
            with open("settings.json", "r", encoding='utf-8') as file:
                settings = json.loads(file.read())
        except:
            settings = {
                "services": []
            }

            with open("settings.json", "w", encoding='utf-8') as file:
                data = json.dumps(settings)
                file.write(data)
        finally:
            return settings

    def setSetting(self, setting, value):
        settings = self.getSettings()

        if settings[setting] != value:
            settings[setting] = value

            with open("settings.json", "w", encoding='utf-8') as file:
                data = json.dumps(settings)
                file.write(data)

                return True
        return False


settings = Settings()


root = tk.Tk()
root.title("Windows Service Control")

servicesFrame = tk.Frame(root)


def printServices():
    servicesList = settings.getSettings()['services']
    servicesFrame.grid_forget()

    for i in range(len(servicesList)):

        currentService = servicesList[i]
        serviceStatus = services('status', currentService['serviceName'])

        tk.Label(servicesFrame, text=currentService['name'], width=20).grid(
            row=i, column=0)

        tk.Label(servicesFrame,
                 text=serviceStatus, width=10).grid(row=i, column=1)

        tk.Button(servicesFrame,
                  text='Start' if serviceStatus != 'Running' else 'Stop',
                  width=5,
                  command=lambda: services('start' if serviceStatus != 'Running' else 'stop', currentService['serviceName'])).grid(row=i, column=2)

        tk.Button(servicesFrame, text='Restart', width=5,
                  command=lambda: services('restart', currentService['serviceName'])).grid(row=i, column=3)

        tk.Button(servicesFrame, text='Remove from list',
                  width=15).grid(row=i, column=4)

    servicesFrame.grid(row=0, column=0)


printServices()

tk.Button(root, text="Add new service").grid(sticky='we')
root.grid_columnconfigure(0, weight=1)

root.mainloop()
