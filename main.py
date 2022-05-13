import pythoncom
from threading import Thread

from win32gui import GetWindowText, GetForegroundWindow
import win32process
import wmi
import subprocess

apps_using = []
hide_process = ['ApplicationFrameHost', 'Music.UI', 'TextInputHost', 'SystemSettings']


def get_all_process():
    all_process = []
    cmd = 'powershell "gps | where {$_.MainWindowTitle } | select id,name,Description'
    proc = subprocess.Popen(cmd, shell=True, stdout=subprocess.PIPE)
    for counter, line in enumerate(proc.stdout):
        if counter > 2 and line.rstrip():
            properties = line.decode().rstrip().lstrip()
            id_app, name, *rest = properties.split(" ")
            app_name_filter = list(filter(lambda item: item != '', rest))

            if name in hide_process:
                continue

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


def manage_apps_running(action_type, process_name):
    global apps_using
    if action_type == 'Creation':
        app_exists_found = list(filter(lambda item: item["name"] == process_name, apps_using))

        if len(app_exists_found) > 0:
            return

        apps_using.append({"name": process_name})
        return

    apps_using_filter = list(filter(lambda item: item["name"] != process_name, apps_using))
    apps_using = apps_using_filter


def on_modify_process(connector_main, action_type="Creation"):
    process_watcher = connector_main.watch_for(notification_type=action_type,
                                               wmi_class="Win32_Process",
                                               delay_secs=1)

    while True:
        process = process_watcher()
        process_id = process.ProcessId
        process_name = process.Name
        format_process_name = process_name.split(".")[0]

        processes = get_all_process()
        current_app_process = get_current_app_process()
        process_found = list(filter(lambda item: item['name'] == format_process_name, processes))

        if len(process_found) > 0:
            manage_apps_running(action_type, process_name)
            print(apps_using)


def listener_creation():
    pythoncom.CoInitialize()
    connector_main = wmi.WMI()
    on_modify_process(connector_main)


def listener_deletion():
    pythoncom.CoInitialize()
    connector_main = wmi.WMI()
    on_modify_process(connector_main, "Deletion")


if __name__ == '__main__':
    t = Thread(target=listener_creation)
    t.daemon = True
    t.start()

    t2 = Thread(target=listener_deletion)
    t2.daemon = True
    t2.start()

    while True:
        pass
