import pythoncom
from threading import Thread

from win32gui import GetWindowText, GetForegroundWindow
import win32process
import wmi
import subprocess


def get_all_process():
    all_process = []
    cmd = 'powershell "gps | where {$_.MainWindowTitle } | select id,name,Description'
    proc = subprocess.Popen(cmd, shell=True, stdout=subprocess.PIPE)
    for counter, line in enumerate(proc.stdout):
        if counter > 2 and line.rstrip():
            properties = line.decode().rstrip().lstrip()
            id_app, name, *rest = properties.split(" ")
            app_name_filter = list(filter(lambda item: item != '', rest))
            all_process.append(
                {"id": id_app, "name": name, "description": ' '.join([str(item) for item in app_name_filter])})
    return all_process


def get_current_app_process():
    exe = ''
    title = GetWindowText(GetForegroundWindow())
    _, pid = win32process.GetWindowThreadProcessId(GetForegroundWindow())

    connector_one = wmi.WMI()

    for p in connector_one.query('SELECT Name FROM Win32_Process WHERE ProcessId = %s' % str(pid)):
        exe = p.Name
        break

    if exe == '':
        return

    app_name, app_extension = exe.split(".")

    return {"id": pid, "app": app_name, "title": title}


def listener():
    pythoncom.CoInitialize()
    connector_main = wmi.WMI()
    process_watcher = connector_main.Win32_Process.watch_for("creation")
    while True:
        new_process = process_watcher()
        processes = get_all_process()
        current_app_process = get_current_app_process()
        print(processes)
        print(current_app_process)


if __name__ == '__main__':
    t = Thread(target=listener)
    t.daemon = True
    t.start()

    while True:
        pass
