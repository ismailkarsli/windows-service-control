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
    printServices()


class Settings():
    def getSettings(self):
        try:
            with open("settings.json", "r", encoding='utf-8') as file:
                settings = json.loads(file.read())
        except:
            settings = {
                "services": [{
                    "name": "Display service name",
                    "serviceName": "Service name"

                }]
            }

            with open("settings.json", "w", encoding='utf-8') as file:
                data = json.dumps(settings, indent=True)
                file.write(data)
        finally:
            return settings

    def setSetting(self, setting, value):
        settings = self.getSettings()

        if settings[setting] != value:
            settings[setting] = value

            with open("settings.json", "w", encoding='utf-8') as file:
                data = json.dumps(settings, indent=True)
                file.write(data)

                return True
        return False


settings = Settings()


root = tk.Tk()
root.title("Windows Service Control")

servicesFrame = tk.Frame(root)


def editServices():
    def saveServices():
        settings.setSetting('services', json.loads(
            servicesText.get(1.0, tk.END).rstrip()))
        printServices()
        editWindow.destroy()

    editWindow = tk.Toplevel(root)
    editWindowFrame = tk.Frame(editWindow)

    servicesList = settings.getSettings()['services']

    servicesText = tk.Text(editWindowFrame, height=40,
                           width=80)
    servicesText.insert(1.0, json.dumps(servicesList, indent=True))
    servicesText.grid(row=0, column=0, columnspan=2)

    tk.Button(editWindowFrame, text='Save',
              command=saveServices).grid(row=1, column=0)
    tk.Button(editWindowFrame, text='Close',
              command=lambda: editWindow.destroy()).grid(row=1, column=1)
    editWindowFrame.pack()


def printServices():
    servicesList = settings.getSettings()['services']

    for widget in servicesFrame.winfo_children():
        widget.destroy()

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
                  command=lambda action='start' if serviceStatus != 'Running' else 'stop', targetService=currentService['serviceName']: services(action, targetService)).grid(row=i, column=2)

        tk.Button(servicesFrame, text='Restart', width=5,
                  command=lambda targetService=currentService['serviceName']: services('restart', targetService)).grid(row=i, column=3)

    servicesFrame.grid(row=0, column=0)


printServices()

tk.Button(root, text="Edit services",
          command=editServices).grid(sticky='we')
root.grid_columnconfigure(0, weight=1)

root.mainloop()
