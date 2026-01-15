from pywinauto import Application
import time
import keyboard

# CONFIGURATION
APP_TITLE = "Area Access Manager"
NETIDS = ["adwitv", "bryan747"]
DELAY = 2.0

# Main window coordinates (relative to main window)
ASSIGN_ACCESS_OFFSET = (701, 87)

# Wizard window coordinates (RELATIVE to wizard window's top-left corner)
TAB_2_CLICK_REL = (507, 244)
NETID_FIELD_CLICK_REL = (1119, 456)
NEXT_BUTTON_REL = (1240, 1048)
SET_ACTIVATION_DATES_REL = (447, 179)
ACTIVATION_DATES_OK_REL = (1165, 950)
NEXT_BUTTON_2_REL = (1243, 1049)
FINAL_OK_REL = (1188, 639)

# Kill switch flag
KILL_SWITCH = False

def on_kill_switch():
    global KILL_SWITCH
    KILL_SWITCH = True
    print("\n\n KILL SWITCH ACTIVATED - Aborting program...")

# Register Ctrl+K as kill switch
keyboard.add_hotkey('ctrl+k', on_kill_switch)

print("Connecting to application...")
print("(Press Ctrl+K anytime to abort the program)\n")

try:
    app_win32 = Application(backend="win32").connect(title_re=APP_TITLE)
    main_win32 = app_win32.window(title_re=APP_TITLE)
    print(f" Connected: {main_win32.window_text()}")
    main_win32.wait("ready", timeout=5)
    print(" Window is ready\n")
except Exception as e:
    print(f" Failed to connect: {e}")
    exit(1)

for netid in NETIDS:
    # Check kill switch at the start of each iteration
    if KILL_SWITCH:
        print("Aborting...")
        break
    
    print(f"Processing {netid}")
    
    try:
        # Click Assign Access button
        print(f"   Clicking Assign Access button at {ASSIGN_ACCESS_OFFSET}")
        main_win32.set_focus()
        time.sleep(0.3)
        main_win32.click_input(coords=ASSIGN_ACCESS_OFFSET)
        time.sleep(DELAY)
        
        if KILL_SWITCH:
            print("Aborting...")
            break
        
        # Get the wizard window and its position
        print(f"   Waiting for wizard window...")
        wizard_window = None
        for attempt in range(15):  # Try for up to 15 seconds
            try:
                wizard = Application(backend="win32").connect(title_re=".*Assignment Wizard.*")
                wizard_window = wizard.top_window()
                rect = wizard_window.rectangle()
                wizard_x = rect.left
                wizard_y = rect.top
                print(f"   Wizard window found at ({wizard_x}, {wizard_y})")
                break
            except:
                if attempt < 14:
                    time.sleep(1)
                    continue
                else:
                    raise Exception("Wizard window did not appear after 15 seconds")
        
        # Click 2nd tab (People)
        print(f"   Clicking 2nd tab")
        wizard_window.click_input(coords=(TAB_2_CLICK_REL[0], TAB_2_CLICK_REL[1]))
        time.sleep(DELAY)
        
        if KILL_SWITCH:
            print("Aborting...")
            break
        
        # Click NetID field
        print(f"   Clicking NetID field and entering {netid}")
        wizard_window.click_input(coords=(NETID_FIELD_CLICK_REL[0], NETID_FIELD_CLICK_REL[1]))
        time.sleep(0.5)
        
        # Clear the field first (Ctrl+A to select all, then delete)
        keyboard.press_and_release('ctrl+a')
        time.sleep(0.1)
        keyboard.press_and_release('delete')
        time.sleep(0.1)
        
        # Type the NetID
        keyboard.write(netid)
        time.sleep(DELAY)
        
        if KILL_SWITCH:
            print("Aborting...")
            break
        
        # Click Next button
        print(f"   Clicking Next (Step 1/4  2/4)")
        wizard_window.click_input(coords=(NEXT_BUTTON_REL[0], NEXT_BUTTON_REL[1]))
        time.sleep(DELAY)
        
        # Click Next button again
        print(f"   Clicking Next (Step 2/4  3/4)")
        wizard_window.click_input(coords=(NEXT_BUTTON_REL[0], NEXT_BUTTON_REL[1]))
        time.sleep(DELAY)
        
        if KILL_SWITCH:
            print("Aborting...")
            break
        
        # Click Set Activation Dates button
        print(f"   Clicking Set Activation Dates")
        wizard_window.click_input(coords=(SET_ACTIVATION_DATES_REL[0], SET_ACTIVATION_DATES_REL[1]))
        time.sleep(DELAY)
        
        # Click OK in Activation Dates popup
        print(f"   Clicking OK in Activation Dates popup")
        wizard_window.click_input(coords=(ACTIVATION_DATES_OK_REL[0], ACTIVATION_DATES_OK_REL[1]))
        time.sleep(DELAY)
        
        # Click Next button
        print(f"   Clicking Next (Step 3/4  4/4)")
        wizard_window.click_input(coords=(NEXT_BUTTON_2_REL[0], NEXT_BUTTON_2_REL[1]))
        time.sleep(DELAY)
        
        # Click Next/Finish button
        print(f"   Clicking Finish")
        wizard_window.click_input(coords=(NEXT_BUTTON_2_REL[0], NEXT_BUTTON_2_REL[1]))
        time.sleep(DELAY)
        
        if KILL_SWITCH:
            print("Aborting...")
            break
        
        # Click OK in final confirmation popup
        print(f"   Clicking OK in confirmation popup")
        wizard_window.click_input(coords=(FINAL_OK_REL[0], FINAL_OK_REL[1]))
        time.sleep(DELAY)
        
        print(f" Completed {netid}\n")
        
    except Exception as e:
        print(f" Failed {netid}: {e}\n")
        import traceback
        traceback.print_exc()

print("All users processed.")