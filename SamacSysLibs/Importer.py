import subprocess
import tkinter as tk
from tkinter import filedialog
import os

SYMACSYS_LIB_DIR = './KiCadLibraryProcessor/SamacSysLibs/'
SAMACSYS_SYMBOL_FILE = SYMACSYS_LIB_DIR + 'JVER_Symbols/JVER.kicad_sym'
SAMACSYS_MODELS_DIR = SYMACSYS_LIB_DIR + '/JVER_Models/'
SAMACSYS_PRETTY_DIR = SYMACSYS_LIB_DIR + 'JVER.pretty/'
SAMACSYS_LOG_FILE = SYMACSYS_LIB_DIR + 'JVER_Installed.yaml'
DOWNLOADS_DIR = '/home/justin/Downloads/'

#
# Function: main
# --------------
# Open a compressed component and install it into
# the JVER KiCad component library.
#

def main():
    file_location = getDownloadedLibrary()
    file_name = ".".join(file_location.split('/')[-1].split('.')[:-1])
    new_directory = "%s%s/" % (SYMACSYS_LIB_DIR, file_name)
    unzip_directory = file_name[4:]

    if (alreadyInstalled(unzip_directory)):
        print("Component already installed!\n")
        return

    make_dir_format = "mkdir -p '%s%s'/"
    make_dir_command = make_dir_format % (SYMACSYS_LIB_DIR, file_name)
    copy_format = "cp '%s' '%s'"
    copy_command = copy_format % (file_location, new_directory)
    unzip_format = "unzip '%s%s.zip' -d '%s'"
    unzip_command = unzip_format % (new_directory, file_name, new_directory)
    subprocess.call(make_dir_command, shell=True)
    subprocess.call(copy_command, shell=True)
    subprocess.call(unzip_command, shell=True)

    unzip_directory = file_name[4:]
    kicad_file_dir = "%s%s/KiCad/" % (new_directory, unzip_directory)
    for file in os.listdir(kicad_file_dir):
        extension = file.split(".")[-1]
        if extension == 'kicad_sym':
            symbol_file = "%s%s" % (kicad_file_dir, file)
            addSymbol(symbol_file)
        elif extension == 'kicad_mod':
            move_format = "mv '%s%s' '%s'"
            move_args = (kicad_file_dir, file, SAMACSYS_PRETTY_DIR)
            move_command = move_format % move_args
            subprocess.call(move_command, shell=True)

    model_dir = "%s%s/3D/" % (new_directory, unzip_directory)
    for file in os.listdir(model_dir):
        move_format = "mv '%s%s' '%s'"
        move_args = (model_dir, file, SAMACSYS_MODELS_DIR)
        move_command = move_format % move_args
        subprocess.call(move_command, shell=True)

    if len(new_directory) > 5:
        remove_format = "rm -rf '%s'"
        remove_command = remove_format % new_directory
        subprocess.call(remove_command, shell=True)
    else:
        print("rm -rf protection")

    with open(SAMACSYS_LOG_FILE, "a") as file:
        file.write("  %s: True\n" % unzip_directory)

#
# Function: getDownloadedLibrary
# ------------------------------
# Open a dialog window for the user to input the location of the
# downloaded component.
#


def getDownloadedLibrary():
    root = tk.Tk()
    root.iconify()
    file_location = filedialog.askopenfile(initialdir=DOWNLOADS_DIR).name
    return file_location

#
# Function: alreadyInstalled
# --------------------------
# Check if the component is already installed prior.
#
# return: Boolean state depending if component is installed.
#


def alreadyInstalled(component):
    with open(SAMACSYS_LOG_FILE, "r") as file:
        data = file.read()
        if component in data:
            return 1
        return 0
    return 1

#
# Function: addSymbol
# -------------------
# Add symbol information to the symbol libray.
#


def addSymbol(symbol_file):
    library = open(SAMACSYS_SYMBOL_FILE, "r")
    symbol = open(symbol_file, "r")

    lines = symbol.readlines()
    symbol.close()
    while "(symbol" not in lines[0]:
        lines = lines[1:]
    total = "\n".join(lines)
    total = total[::-1]
    index = total.index(')')
    total = total[:index] + total[index+1:]
    total = total[::-1]

    library_data = library.read()
    library.close()
    library_data = library_data[::-1]
    index = library_data.index(')')
    library_data = library_data[:index] + library_data[index+1:]
    library_data = library_data[::-1]
    library_data += total
    library_data.strip()
    library_data += "\n)"

    for _ in range(20):
        library_data = library_data.replace("\n\n", "\n")

    library = open(SAMACSYS_SYMBOL_FILE, "w")
    library.write(library_data)
    library.close()


if __name__ == "__main__":
    main()
