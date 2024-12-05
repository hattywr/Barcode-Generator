import keyboard
import subprocess

def run_python_script():
    # Replace 'your_script.py' with the path to your Python script
    subprocess.run(['python', "serialNumBarcodeGenerator"])
# Bind the function to the F5 key
keyboard.add_hotkey('F1', run_python_script)

print("Press F1 to run the Python script")

# Keep the script running
keyboard.wait('esc')  # Exit on pressing 'esc'
