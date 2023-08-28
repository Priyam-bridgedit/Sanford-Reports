import sys
from cx_Freeze import setup, Executable

# Build options
build_options = {
    'packages': ['os', 'tkinter', 'pandas', 'pyodbc', 'configparser', 'apscheduler'],
    'excludes': [],
    'include_files': ['config.ini'],  # Include the config.ini file
}

# Executable options
executables = [
    Executable(
        script='sandfordReport.py',  # This is the target script for the first executable
        base=None, 
        targetName='SanfordLimitedReport.exe',  
    ),
    Executable(
        script='background_scheduler.py',  # This is the target script for the second executable
        base=None, 
        targetName='BackgroundScheduler.exe', 
    )
]

# Create the setup
setup(
    name='Sanford Limited Report',  # Name of the application
    version='1.0',
    description='Sanford Limited Report',
    options={'build_exe': build_options},
    executables=executables,
)
