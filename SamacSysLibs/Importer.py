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

def main():
    root = tk.Tk()
    root.iconify()
    file_location = filedialog.askopenfile(initialdir=DOWNLOADS_DIR).name
    file_name = ".".join(file_location.split('/')[-1].split('.')[:-1])
    new_directory = "%s%s/" % (SYMACSYS_LIB_DIR, file_name)
    unzip_dir = file_name[4:]

    with open(SAMACSYS_LOG_FILE, "r") as file:
        data = file.read()
        if unzip_dir in data:
            print("Component already installed!\n")
            return

    subprocess.call("mkdir -p '%s%s'/" % (SYMACSYS_LIB_DIR, file_name), shell=True)
    subprocess.call("cp '%s' '%s'" % (file_location, new_directory, ), shell=True)
    subprocess.call("unzip '%s%s.zip' -d '%s'" % (new_directory, file_name, new_directory), shell=True)

    unzip_dir = file_name[4:]
    kicad_file_dir = "%s%s/KiCad/" % (new_directory, unzip_dir)
    for file in os.listdir(kicad_file_dir):
        extension = file.split(".")[-1]
        if extension == 'kicad_sym':
            symbol_file = "%s%s" % (kicad_file_dir, file)
            addSymbol(symbol_file)
        elif extension == 'kicad_mod':
            subprocess.call("mv '%s%s' '%s'" % (kicad_file_dir, file, SAMACSYS_PRETTY_DIR), shell=True)

    model_dir = "%s%s/3D/" % (new_directory, unzip_dir)
    for file in os.listdir(model_dir):
        subprocess.call("mv '%s%s' '%s'" % (model_dir, file, SAMACSYS_MODELS_DIR), shell=True)

    if len(new_directory) > 5:
        subprocess.call("rm -rf '%s'" % (new_directory), shell=True)
    else:
        print("rm -rf protection")

    with open(SAMACSYS_LOG_FILE, "a") as file:
        file.write("  %s: True\n" % unzip_dir)

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