import subprocess
import tkinter as tk
from tkinter import filedialog
import os

SYMACSYS_LIB_DIR = '/media/justin/StorageSSD/Programming/KiCadLibraryProcessor/SamacSysLibs/'
SAMACSYS_SYMBOL_FILE = '/media/justin/StorageSSD/Programming/KiCadLibraryProcessor/SamacSysLibs/JVER_Symbols/JVER.kicad_sym'
SAMACSYS_MODELS_DIR = '/media/justin/StorageSSD/Programming/KiCadLibraryProcessor/SamacSysLibs/JVER_Models/'
SAMACSYS_PRETTY_DIR = '/media/justin/StorageSSD/Programming/KiCadLibraryProcessor/SamacSysLibs/JVER.pretty/'
SAMACSYS_LOG_FILE = '/media/justin/StorageSSD/Programming/KiCadLibraryProcessor/SamacSysLibs/JVER_Installed.yaml'
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

    make_dir_command = "mkdir -p '%s%s'/" % (SYMACSYS_LIB_DIR, file_name)
    copy_command = "cp '%s' '%s'" % (file_location, new_directory)
    unzip_command = "unzip '%s%s.zip' -d '%s'" % (new_directory, file_name, new_directory)
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
            move_command = "mv '%s%s' '%s'" % (kicad_file_dir, file, SAMACSYS_PRETTY_DIR)
            subprocess.call(move_command, shell=True)

    model_dir = "%s%s/3D/" % (new_directory, unzip_directory)
    for file in os.listdir(model_dir):
        move_command = "mv '%s%s' '%s'" % (model_dir, file, SAMACSYS_MODELS_DIR)
        subprocess.call(move_command, shell=True)

    if len(new_directory) > 5:
        remove_command = "rm -rf '%s'" % (new_directory)
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
    if unzip_dir in component:
        return 1
    return 0

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