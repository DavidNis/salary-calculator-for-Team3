import PyInstaller.__main__
import sys
import os

# Get the directory containing the script
current_dir = os.path.dirname(os.path.abspath(__file__))

# Define the paths to your Python files and resources
main_script = os.path.join(current_dir, 'gui.py')
salary_calc = os.path.join(current_dir, 'salary_calc.py')
icon_file = os.path.join(current_dir, 'Celery.ico')

# Define PyInstaller arguments
args = [
    main_script,  # Your main script
    '--onefile',  # Create a single executable file
    '--windowed',  # Don't show console window when running the executable
    '--icon', icon_file,  # Set the icon for the executable
    '--name', 'SalaryCalculator',  # Name of the output executable
    '--add-data', f'{icon_file};.',  # Include the icon file in the executable
    '--add-data', f'{salary_calc};.',  # Include the salary_calc.py file
    # Add required packages
    '--hidden-import', 'tkinter',
    '--hidden-import', 'pandas',
    '--hidden-import', 'openpyxl',
    '--hidden-import', 'ttkthemes',
    '--hidden-import', 'tkinter.font',
    '--hidden-import', 'babel.numbers',
    # Additional imports for Excel handling
    '--hidden-import', 'numpy',
    '--hidden-import', 'datetime',
    '--hidden-import', 'threading',
    # Clean build directories
    '--clean',
    # Exclude unnecessary packages to reduce size
    '--exclude-module', 'matplotlib',
    '--exclude-module', 'PyQt5',
    '--exclude-module', 'PyQt6',
    '--exclude-module', 'PySide2',
    '--exclude-module', 'PySide6',
    # Additional options for optimization
    '--noupx',  # Disable UPX compression (can cause issues on some systems)
    '--noconsole',  # Another way to specify no console window
    # Debug options (comment out for final build)
    #'--debug', 'all',
    #'--log-level', 'DEBUG',
]

# Run PyInstaller
PyInstaller.__main__.run(args)

# Print success message
print("\nBuild completed! Check the 'dist' folder for your executable.")
print("Note: If the executable doesn't work, try running it from the command line to see error messages.")