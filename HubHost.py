# 5
import os
import psutil
import socket
import winreg
import updaterHost


def get_installed_software():

    def filter_name(name):

        name = name.replace(" (x64)", "").replace(

            " (x86)", "").replace(" x64", "")

        name = name.replace(" x86", "").replace(" X64", "").replace(" X86", "")

        name = name.replace(" (64-bit)", "").split(" - ", 1)[0].strip()

        return name

    def find_duplicates(software_list):

        unique_software = []

        duplicates = []

        for s in software_list:

            if s not in unique_software:

                unique_software.append(s)

            else:

                duplicates.append(s)

        return unique_software, duplicates

    software_list = []

    keys = [winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER]

    flags = [winreg.KEY_WOW64_32KEY, winreg.KEY_WOW64_64KEY]

    for key in keys:

        for flag in flags:

            try:

                aReg = winreg.ConnectRegistry(None, key)

                aKey = winreg.OpenKey(

                    aReg, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", 0, winreg.KEY_READ | flag)

                count_subkey = winreg.QueryInfoKey(aKey)[0]

                for i in range(count_subkey):

                    software = {}

                    try:

                        asubkey_name = winreg.EnumKey(aKey, i)

                        asubkey = winreg.OpenKey(aKey, asubkey_name)

                        software["name"] = winreg.QueryValueEx(

                            asubkey, "DisplayName")[0]

                        software["version"] = winreg.QueryValueEx(

                            asubkey, "DisplayVersion")[0]

                        software["publisher"] = winreg.QueryValueEx(

                            asubkey, "Publisher")[0]

                        software["name"] = filter_name(software["name"])

                        software_list.append(software)

                    except EnvironmentError:

                        pass

            except:

                pass

    unique_software, duplicates = find_duplicates(software_list)

    # If there are duplicates, print them and remove them from the list

    if duplicates:

        software_list = unique_software

    return software_list


def main():

    current_version = 5

    updaterHost.update_program(current_version)

    # Check ESET

    eset_active = any("egui.exe" in proc.name().lower(

    ) or "ekrn.exe" in proc.name().lower() for proc in psutil.process_iter())

    # Check WG

    hostSensor_active = any("host_sensor.exe" in proc.name().lower()

                            for proc in psutil.process_iter())

    # Check BitLocker

    bitlocker_enabled = False

    if os.name == "nt":

        for line in os.popen("manage-bde -status C:"):

            if "Protection Status" in line and "On" in line:

                bitlocker_enabled = True

    # Get IP

    ip_address = socket.gethostbyname(socket.gethostname())

    # Get version of Office

    office_key = winreg.OpenKey(

        winreg.HKEY_LOCAL_MACHINE, "SOFTWARE\\Microsoft\\Office")

    office_versions = []

    try:

        i = 0

        while True:

            subkey = winreg.EnumKey(office_key, i)

            try:

                version = float(subkey)

                office_versions.append(version)

            except ValueError:

                pass

            i += 1

    except WindowsError:

        pass

    if len(office_versions) > 0:

        office_version = str(max(office_versions))

    else:

        office_version = "Not found"

    # Check admin

    admin_rights = os.access(os.sep.join(

        [os.environ.get("SystemRoot", "C:\\windows"), "temp"]), os.W_OK)

    # Get computer name and Windows version

    computer_name = socket.gethostname()

    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion") as key:

        windows_version = winreg.QueryValueEx(key, "EditionID")[0]

    # Get list of installed software

    installed_software = get_installed_software()

    # Write results to database

    updaterHost.replace_the_database(computer_name)

    # Write results to file

    user_data = [

        f"ESET,{eset_active}",

        f"WG,{hostSensor_active}",

        f"BITLOCKER,{bitlocker_enabled}",

        f"IP,{ip_address}",

        f"OFFICE,{office_version}",

        f"ADMIN,{admin_rights}",

        f"PRO,{windows_version}",

        f"NAME,{computer_name}",

    ] + [f"{s['name']},{s['version']},{s['publisher']}"

         for s in installed_software]

    updaterHost.write_to_database(computer_name, user_data)


if __name__ == '__main__':
    main()
