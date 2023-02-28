import os, subprocess, time
from subprocess import CREATE_NO_WINDOW, HIGH_PRIORITY_CLASS


def convert(
                queue,
                files,
                pathOutputFolder,
                pathConverterFolder,
                varProjectType
            ):

    conv_IfcToCsv = os.path.join(pathConverterFolder, 'IfcToCsv.exe')
    conv_IfcColladaExporter = os.path.join(pathConverterFolder, 'IfcColladaExporter.exe')
    conv_RvtToCsv = os.path.join(pathConverterFolder, 'RvtToCsv.exe')
    conv_RvtColladaExporter = os.path.join(pathConverterFolder, 'RvtColladaExporter.exe')

    csv_files = []

    if varProjectType == ".ifc":
        for f in files:
            file = f["name"]
            filepath = f["path"]
            fileOutputPath = os.path.join(pathOutputFolder, file[:-3])

            queue.put("Converting " + file + " ...")

            subprocess.run([conv_IfcToCsv, filepath, fileOutputPath + 'csv'], cwd=pathConverterFolder, creationflags=CREATE_NO_WINDOW | HIGH_PRIORITY_CLASS)
            queue.put("Conversion Done: " + file[:-3] + 'csv')
            time.sleep(1)
            subprocess.run([conv_IfcColladaExporter, filepath, fileOutputPath + 'dae'], cwd=pathConverterFolder,creationflags=CREATE_NO_WINDOW | HIGH_PRIORITY_CLASS)
            queue.put("Conversion Done: " + file[:-3] + 'dae')

            csv_files.append(file[:-3]+'csv')

    elif varProjectType == '.rvt':
        for f in files:
            file = f["name"]
            filepath = f["path"]

            fileOutputPath = os.path.join(pathOutputFolder, file[:-3])

            queue.put("Converting " + file + " ...")

            subprocess.run([conv_RvtColladaExporter, filepath, fileOutputPath + 'dae'], cwd=pathConverterFolder, creationflags = CREATE_NO_WINDOW | HIGH_PRIORITY_CLASS)
            queue.put("Conversion Done: " + file[:-3] + 'dae')

            time.sleep(1)

            subprocess.run([conv_RvtToCsv, filepath, fileOutputPath + 'dae', fileOutputPath + 'csv'], cwd=pathConverterFolder, creationflags = CREATE_NO_WINDOW | HIGH_PRIORITY_CLASS)
            queue.put("Conversion Done: " + file[:-3] + 'csv')

            csv_files.append(file[:-3]+'csv')
        time.sleep(3)
    return csv_files
