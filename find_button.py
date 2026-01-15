from pywinauto import Application
import pyautogui
import time
import sys

APP_TITLE = "Area Access Manager"

print("Connecting to Area Access Manager...")
app_win32 = Application(backend="win32").connect(title_re=APP_TITLE)
main_win32 = app_win32.window(title_re=APP_TITLE)

print(f"✔ Connected: {main_win32.window_text()}")

rect = main_win32.rectangle()
print(f"\nWindow position: ({rect.left}, {rect.top})")
print(f"Window size: {rect.width()} x {rect.height()}")

print("\n" + "="*60)
print("MOUSE COORDINATE TRACKER")
print("="*60)
print("Hover your mouse over the 'Assign Access' button")
print("The coordinates shown are RELATIVE to the window's top-left corner")
print("Press Ctrl+C when you've found the button position\n")

try:
    while True:
        # Get absolute mouse position
        mouse_x, mouse_y = pyautogui.position()
        
        # Calculate relative position to window
        relative_x = mouse_x - rect.left
        relative_y = mouse_y - rect.top
        
        # Print on same line
        print(f"\rMouse position - Absolute: ({mouse_x:4d}, {mouse_y:4d})  |  Window-relative: ({relative_x:4d}, {relative_y:4d})  ", end='', flush=True)
        
        time.sleep(0.1)
        
except KeyboardInterrupt:
    print("\n\nDone! Update ASSIGN_ACCESS_OFFSET in openSesame.py with the window-relative coordinates.")
