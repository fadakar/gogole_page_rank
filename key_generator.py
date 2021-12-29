import hashlib
from termcolor import colored
from pyfiglet import *


def get_serial_number_of_physical_disk(drive_letter='C:'):
    import wmi
    c = wmi.WMI()
    logical_disk = c.Win32_LogicalDisk(Caption=drive_letter)[0]
    partition = logical_disk.associators()[1]
    physical_disc = partition.associators()[0]
    return physical_disc.SerialNumber


def get_system_key() -> str:
    serial = get_serial_number_of_physical_disk(os.getenv('SystemDrive'))
    serial = "gholamreza_" + serial + "_fadakar"
    return hashlib.md5(serial.encode('utf-8')).hexdigest()


def check_key(key: str) -> bool:
    system_key = get_system_key()
    if key == system_key:
        return True
    else:
        return False


if __name__ == "__main__":
    f = Figlet()
    print(colored(f.renderText('Page Rank License'), 'green'))
    print(colored('\tCreated By Gholamreza Fadakar', 'green'))
    print(colored('\tLicense Key: ', 'red'), end='')
    print(colored(get_system_key(), 'green'))
    print(colored('\tput this key in key.license file in application folder', 'red'))

