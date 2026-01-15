from pywinauto import Application
import time
import keyboard
import sys

# CONFIGURATION
APP_TITLE = "Area Access Manager"
NETIDS = ["adwitv", "bryan747"]
CLICK_DELAY = 0.25
KEY_DELAY = 0.05
BETWEEN_USERS_DELAY = 1.0

# Debug: prints each automated action so you can spot any extra click/keypress.
DEBUG_ACTIONS = True

action_counter = 0

# Main window coordinates (relative to main window)
ASSIGN_ACCESS_OFFSET = (465, 59)

# Wizard-step coordinates (relative to the MAIN window client area)
TAB_2_CLICK_REL = (665, 304)
NETID_FIELD_CLICK_REL = (1006, 441)

print("Connecting to application...")
print("(Press Ctrl+C anytime to abort the program)\n")

try:
    try:
        app_win32 = Application(backend="win32").connect(title_re=APP_TITLE)
        main_win32 = app_win32.window(title_re=APP_TITLE)
        print(f" Connected: {main_win32.window_text()}")
        main_win32.wait("ready", timeout=5)
        print(" Window is ready\n")

        def log_action(message: str):
            global action_counter
            if not DEBUG_ACTIONS:
                return
            action_counter += 1
            print(f"   [ACTION {action_counter:02d}] {message}")

        def click_main(coords):
            # Click using the SAME coordinate system as main_win32.click_input(coords=...)
            # (i.e., the one used for ASSIGN_ACCESS_OFFSET).
            try:
                rect = main_win32.rectangle()
                approx_screen = (rect.left + coords[0], rect.top + coords[1])
                log_action(f"CLICK main @ {coords} (approx screen {approx_screen})")
            except Exception:
                log_action(f"CLICK main @ {coords}")
            main_win32.click_input(coords=coords)

        def press(keys: str):
            log_action(f"KEY {keys}")
            keyboard.press_and_release(keys)
    except Exception as e:
        print(f" Failed to connect: {e}")
        exit(1)

    for netid in NETIDS:
        # Give the confirmation popup time to close before starting next user
        time.sleep(BETWEEN_USERS_DELAY)
        print(f"Processing {netid}")
        
        try:
            # Click Assign Access button
            print(f"   Clicking Assign Access button at {ASSIGN_ACCESS_OFFSET}")
            main_win32.set_focus()
            time.sleep(0.3)
            log_action(f"CLICK Assign Access @ {ASSIGN_ACCESS_OFFSET}")
            main_win32.click_input(coords=ASSIGN_ACCESS_OFFSET)
            time.sleep(CLICK_DELAY)
            
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
            click_main(TAB_2_CLICK_REL)
            time.sleep(CLICK_DELAY)
            
            # Click NetID field
            print(f"   Clicking NetID field and entering {netid}")
            click_main(NETID_FIELD_CLICK_REL)
            time.sleep(CLICK_DELAY)
            
                      
            # Type the NetID
            keyboard.write(netid)
            time.sleep(KEY_DELAY)
            
            # Next (Step 1/4 -> 2/4): Enter
            print(f"   Next (Step 1/4  2/4) via Enter")
            try:
                wizard_window.set_focus()
            except Exception:
                pass
            press('enter')
            time.sleep(KEY_DELAY)
            
            # Next (Step 2/4 -> 3/4): Enter
            print(f"   Next (Step 2/4  3/4) via Enter")
            try:
                wizard_window.set_focus()
            except Exception:
                pass
            press('enter')
            time.sleep(KEY_DELAY)
            

            # Set Activation Dates: Tab, then Enter
            print(f"   Set Activation Dates via Tab+Enter")
            try:
                wizard_window.set_focus()
            except Exception:
                pass
            press('tab')
            time.sleep(KEY_DELAY)
            press('enter')
            time.sleep(KEY_DELAY)
            
            # OK in Activation Dates popup: Enter
            print(f"   OK in Activation Dates popup via Enter")
            press('enter')
            time.sleep(KEY_DELAY)
            
            # Next (Step 3/4 -> 4/4): Tab x4, then Enter
            print(f"   Next (Step 3/4  4/4) via Tab x4 + Enter")
            try:
                wizard_window.set_focus()
            except Exception:
                pass
            for _ in range(4):
                press('tab')
                time.sleep(KEY_DELAY)
            press('enter')
            time.sleep(KEY_DELAY)
            
            # Finish: Enter
            print(f"   Finish via Enter")
            try:
                wizard_window.set_focus()
            except Exception:
                pass
            press('enter')
            time.sleep(KEY_DELAY)
            

            # Click OK in final confirmation popup
            print(f"   Clicking OK in confirmation popup")
            # Avoid window-title ambiguity (two windows named "Area Access Manager").
            # The OK button is the default, so Enter reliably confirms.
            try:
                wizard_window.set_focus()
            except Exception:
                pass
            time.sleep(0.2)
            press('enter')
            time.sleep(KEY_DELAY)
            
            print(f" Completed {netid}\n")
            
        except Exception as e:
            print(f" Failed {netid}: {e}\n")
            import traceback
            traceback.print_exc()

    print("All users processed.")
    
except KeyboardInterrupt:
    print("\n\nProgram interrupted by user.")
