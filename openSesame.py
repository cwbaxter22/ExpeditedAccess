from __future__ import annotations

import queue
import re
import threading
import time
import tkinter as tk
from dataclasses import dataclass
from tkinter import messagebox, ttk

import ctypes

import keyboard
from pywinauto import Application


# ----------------------------
# Defaults / Configuration
# ----------------------------

APP_TITLE = "Area Access Manager"


@dataclass
class Config:
    click_delay: float = 0.25
    key_delay: float = 0.05
    between_users_delay: float = 1.0

    assign_access_offset: tuple[int, int] = (465, 59)
    tab_2_click_rel: tuple[int, int] = (665, 304)
    netid_field_click_rel: tuple[int, int] = (1006, 441)

    debug_actions: bool = True


DISCLAIMER_TEXT = (
    "DISCLAIMER: THIS PROGRAM WILL ONLY WORK IF AREA ACCESS MANAGER IS SCALED CORRECTLY AND OPEN.\n"
    "Sometimes, on laptops for example, Area Access Manager will show up as smaller than usual. "
    "This will create a coordinate error and cause the program to fail.\n"
    "Sometimes, this can be corrected through adjustments made on the Advanced Tab.\n"
    "The program works through a combination of moving the mouse to coordinates and clicking, and sending keyboard input.\n"
    "DO NOT MOVE THE MOUSE DURING OPERATION\n\n"
    "SETUP:\n"
    "1) Open Area Access Manager\n"
    "2) Confirm scaling is correct\n"
    "3) Enter UW NetIDs each on a new line in the Input Box\n"
    "4) Note the location of the \"ABORT\" button incase steps 5 and 6 malfunction\n"
    "After STEP 5, do not touch the mouse unless you need to click \"ABORT\"\n"
    "5) Press \"Assign Access\"\n"
    "6) Supervise the first batch to confirm there are no coordinate errors\n"
    "7) Once access has been assigned to all users, click the top of the \"Activate\" column on Area Access Manager to sort by newest access granted first\n"
    "8) Select the NetIDs that were just added and set specific dates and times"
)


ADVANCED_HELP_TEXT = (
    "If the program runs into an error in the clicking portion, the culprit may be a differently scaled screen. "
    "You can press the \"Coord Check\" button to get a read-out of what the coordinates are of your mouse on Area Access Manager in real-time.\n"
    "1) Open Area Access Manager\n"
    "2) Press Coord Check button\n"
    "3) Hover over the \"Assign Access\" yellow button, \"UWID\" tab on the pop-up, and anywhere on the \"NetID\" text entry box.\n"
    "4) Note the coordinates for each and update them on the advanced tab\n\n"
    "Another error the program can encounter is moving more quickly than Area Access Manager can keep up.\n"
    "In this situation, try increasing the time for the Click, Key, and Between-Users delays.\n"
    "Start with the Between-Users, then Click, and then Key."
)


class AbortRequested(Exception):
    pass


class _POINT(ctypes.Structure):
    _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]


def _get_cursor_pos() -> tuple[int, int]:
    pt = _POINT()
    if ctypes.windll.user32.GetCursorPos(ctypes.byref(pt)) == 0:
        raise OSError("GetCursorPos failed")
    return int(pt.x), int(pt.y)


def _sleep_or_abort(seconds: float, abort_event: threading.Event):
    end_time = time.time() + max(0.0, seconds)
    while time.time() < end_time:
        if abort_event.is_set():
            raise AbortRequested()
        time.sleep(min(0.05, end_time - time.time()))


def run_assign_access(netids: list[str], cfg: Config, abort_event: threading.Event, log):
    if not netids:
        log("No NetIDs provided.\n")
        return

    action_counter = 0

    def log_action(message: str):
        nonlocal action_counter
        if not cfg.debug_actions:
            return
        action_counter += 1
        log(f"[ACTION {action_counter:02d}] {message}\n")

    def press(keys: str):
        if abort_event.is_set():
            raise AbortRequested()
        log_action(f"KEY {keys}")
        keyboard.press_and_release(keys)

    log("Connecting to application...\n")
    log("(Press ABORT anytime to stop)\n\n")

    app_win32 = Application(backend="win32").connect(title_re=APP_TITLE)
    main_win32 = app_win32.window(title_re=APP_TITLE)
    main_win32.wait("ready", timeout=10)
    log(f"Connected: {main_win32.window_text()}\n")

    def click_main(coords: tuple[int, int], label: str):
        if abort_event.is_set():
            raise AbortRequested()
        try:
            rect = main_win32.rectangle()
            approx_screen = (rect.left + coords[0], rect.top + coords[1])
            log_action(f"CLICK {label} @ {coords} (approx screen {approx_screen})")
        except Exception:
            log_action(f"CLICK {label} @ {coords}")
        main_win32.click_input(coords=coords)

    for netid in netids:
        if abort_event.is_set():
            raise AbortRequested()

        # Give the prior confirmation popup time to close before starting next user
        _sleep_or_abort(cfg.between_users_delay, abort_event)
        log(f"Processing {netid}\n")

        # Click Assign Access
        main_win32.set_focus()
        _sleep_or_abort(0.2, abort_event)
        click_main(cfg.assign_access_offset, "Assign Access")
        _sleep_or_abort(cfg.click_delay, abort_event)

        # Wait for wizard window
        log("Waiting for wizard window...\n")
        wizard_window = None
        for attempt in range(15):
            if abort_event.is_set():
                raise AbortRequested()
            try:
                wizard = Application(backend="win32").connect(title_re=".*Assignment Wizard.*")
                wizard_window = wizard.top_window()
                rect = wizard_window.rectangle()
                log(f"Wizard window found at ({rect.left}, {rect.top})\n")
                break
            except Exception:
                _sleep_or_abort(1.0, abort_event)
        if wizard_window is None:
            raise RuntimeError("Wizard window did not appear after 15 seconds")

        # Click 2nd tab (UWID) — coordinate-based
        log("Clicking UWID tab\n")
        click_main(cfg.tab_2_click_rel, "UWID tab")
        _sleep_or_abort(cfg.click_delay, abort_event)

        # Click NetID field — coordinate-based
        log(f"Clicking NetID field and entering {netid}\n")
        click_main(cfg.netid_field_click_rel, "NetID field")
        _sleep_or_abort(cfg.click_delay, abort_event)

        # Type NetID
        keyboard.write(netid)
        _sleep_or_abort(cfg.key_delay, abort_event)

        # Next: Enter
        log("Next (Step 1/4 -> 2/4) via Enter\n")
        try:
            wizard_window.set_focus()
        except Exception:
            pass
        press('enter')
        _sleep_or_abort(cfg.key_delay, abort_event)

        # Next: Enter
        log("Next (Step 2/4 -> 3/4) via Enter\n")
        try:
            wizard_window.set_focus()
        except Exception:
            pass
        press('enter')
        _sleep_or_abort(cfg.key_delay, abort_event)

        # Set Activation Dates: Tab then Enter
        log("Set Activation Dates via Tab+Enter\n")
        try:
            wizard_window.set_focus()
        except Exception:
            pass
        press('tab')
        _sleep_or_abort(cfg.key_delay, abort_event)
        press('enter')
        _sleep_or_abort(cfg.key_delay, abort_event)

        # Activation Dates popup OK: Enter
        log("OK in Activation Dates popup via Enter\n")
        press('enter')
        _sleep_or_abort(cfg.key_delay, abort_event)

        # Next: Tab x4 then Enter
        log("Next (Step 3/4 -> 4/4) via Tab x4 + Enter\n")
        try:
            wizard_window.set_focus()
        except Exception:
            pass
        for _ in range(4):
            press('tab')
            _sleep_or_abort(cfg.key_delay, abort_event)
        press('enter')
        _sleep_or_abort(cfg.key_delay, abort_event)

        # Finish: Enter
        log("Finish via Enter\n")
        try:
            wizard_window.set_focus()
        except Exception:
            pass
        press('enter')
        _sleep_or_abort(cfg.key_delay, abort_event)

        # Final confirmation OK: Enter (no coordinates)
        log("Clicking OK in confirmation popup (pressing Enter)\n")
        try:
            wizard_window.set_focus()
        except Exception:
            pass
        _sleep_or_abort(0.2, abort_event)
        press('enter')
        _sleep_or_abort(cfg.key_delay, abort_event)

        log(f"Completed {netid}\n\n")

    log("All users processed.\n")


def run_coord_check(cfg: Config, abort_event: threading.Event, log):
    app_win32 = Application(backend="win32").connect(title_re=APP_TITLE)
    main_win32 = app_win32.window(title_re=APP_TITLE)
    main_win32.wait("ready", timeout=10)
    rect = main_win32.rectangle()

    # Output should stay blank except for a single live x/y readout.
    last = None
    while not abort_event.is_set():
        x, y = _get_cursor_pos()
        rel_x = x - rect.left
        rel_y = y - rect.top
        line = f"x: {rel_x}\ny: {rel_y}\n"
        # Only update when it changes
        if line != last:
            log(line)
            last = line
        time.sleep(0.05)


# ----------------------------
# GUI
# ----------------------------


class TextSink:
    def __init__(self, text_widget: tk.Text):
        self.text = text_widget

    def append(self, s: str):
        self.text.configure(state="normal")
        self.text.insert("end", s)
        self.text.see("end")
        self.text.configure(state="disabled")

    def set(self, s: str):
        self.text.configure(state="normal")
        self.text.delete("1.0", "end")
        self.text.insert("end", s)
        self.text.see("end")
        self.text.configure(state="disabled")


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Expedited Access")
        self.geometry("900x650")

        self.cfg = Config()
        self.abort_event = threading.Event()
        self.worker_thread: threading.Thread | None = None

        self.log_queue: queue.Queue[tuple[str, str]] = queue.Queue()

        self._build_ui()
        self.after(50, self._drain_log_queue)

    # ---- UI ----

    def _build_ui(self):
        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True)

        self.tab_input = ttk.Frame(notebook)
        self.tab_adv = ttk.Frame(notebook)
        notebook.add(self.tab_input, text="Input")
        notebook.add(self.tab_adv, text="Advanced")

        self._build_input_tab(self.tab_input)
        self._build_advanced_tab(self.tab_adv)

    def _build_input_tab(self, parent: ttk.Frame):
        # NetIDs input
        ttk.Label(parent, text="NetIDs - each on a newline").pack(anchor="w", padx=10, pady=(10, 0))
        self.netids_text = tk.Text(parent, height=6)
        self.netids_text.pack(fill="x", padx=10, pady=5)

        # Buttons row
        btn_row = ttk.Frame(parent)
        btn_row.pack(fill="x", padx=10, pady=5)

        self.btn_assign = ttk.Button(btn_row, text="Assign Access", command=self.on_assign_access)
        self.btn_assign.pack(side="left")

        ttk.Button(btn_row, text="Help", command=lambda: messagebox.showinfo("Help", DISCLAIMER_TEXT)).pack(
            side="left", padx=(8, 0)
        )

        self.btn_abort_1 = ttk.Button(btn_row, text="ABORT", command=self.on_abort)
        self.btn_abort_1.pack(side="right")

        # Output
        ttk.Label(parent, text="Output").pack(anchor="w", padx=10, pady=(10, 0))
        out_frame = ttk.Frame(parent)
        out_frame.pack(fill="x", expand=False, padx=10, pady=5)
        self.output_text = tk.Text(out_frame, state="disabled", height=14)
        out_scroll = ttk.Scrollbar(out_frame, orient="vertical", command=self.output_text.yview)
        self.output_text.configure(yscrollcommand=out_scroll.set)
        self.output_text.pack(side="left", fill="both", expand=True)
        out_scroll.pack(side="right", fill="y")
        self.output_sink = TextSink(self.output_text)

        # Spacer so Output doesn't take the whole tab
        ttk.Frame(parent).pack(fill="both", expand=True)

    def _build_advanced_tab(self, parent: ttk.Frame):
        top = ttk.Frame(parent)
        top.pack(fill="x", padx=10, pady=10)

        # Delay table
        delays = ttk.LabelFrame(top, text="Delays")
        delays.pack(side="left", fill="both", expand=True, padx=(0, 10))

        self.var_click_delay = tk.StringVar(value=str(self.cfg.click_delay))
        self.var_key_delay = tk.StringVar(value=str(self.cfg.key_delay))
        self.var_between_users_delay = tk.StringVar(value=str(self.cfg.between_users_delay))

        self._add_float_row(delays, "Click_Delay", self.var_click_delay, 0)
        self._add_float_row(delays, "Key_Delay", self.var_key_delay, 1)
        self._add_float_row(delays, "Between_Users_Delay", self.var_between_users_delay, 2)

        # Coordinates table
        coords = ttk.LabelFrame(top, text="Coordinates")
        coords.pack(side="left", fill="both", expand=True)

        self.var_assign_x = tk.StringVar(value=str(self.cfg.assign_access_offset[0]))
        self.var_assign_y = tk.StringVar(value=str(self.cfg.assign_access_offset[1]))
        self.var_tab_x = tk.StringVar(value=str(self.cfg.tab_2_click_rel[0]))
        self.var_tab_y = tk.StringVar(value=str(self.cfg.tab_2_click_rel[1]))
        self.var_netid_x = tk.StringVar(value=str(self.cfg.netid_field_click_rel[0]))
        self.var_netid_y = tk.StringVar(value=str(self.cfg.netid_field_click_rel[1]))

        self._add_coord_row(coords, "ASSIGN_ACCESS_OFFSET", self.var_assign_x, self.var_assign_y, 0)
        self._add_coord_row(coords, "TAB_2_CLICK_REL", self.var_tab_x, self.var_tab_y, 1)
        self._add_coord_row(coords, "NETID_FIELD_CLICK_REL", self.var_netid_x, self.var_netid_y, 2)

        # Buttons
        btn_row = ttk.Frame(parent)
        btn_row.pack(fill="x", padx=10, pady=(0, 10))

        self.btn_coord = ttk.Button(btn_row, text="Coord Check", command=self.on_coord_check)
        self.btn_coord.pack(side="left")

        ttk.Button(btn_row, text="Advanced Help", command=lambda: messagebox.showinfo("Advanced Help", ADVANCED_HELP_TEXT)).pack(
            side="left", padx=(8, 0)
        )

        self.btn_abort_2 = ttk.Button(btn_row, text="ABORT", command=self.on_abort)
        self.btn_abort_2.pack(side="right")

        # Output
        ttk.Label(parent, text="Output").pack(anchor="w", padx=10)
        out_frame = ttk.Frame(parent)
        out_frame.pack(fill="x", expand=False, padx=10, pady=5)
        self.adv_output_text = tk.Text(out_frame, state="disabled", height=14)
        out_scroll = ttk.Scrollbar(out_frame, orient="vertical", command=self.adv_output_text.yview)
        self.adv_output_text.configure(yscrollcommand=out_scroll.set)
        self.adv_output_text.pack(side="left", fill="both", expand=True)
        out_scroll.pack(side="right", fill="y")
        self.adv_output_sink = TextSink(self.adv_output_text)

        # Spacer so Output doesn't take the whole tab
        ttk.Frame(parent).pack(fill="both", expand=True)

    # ---- Validation UI helpers ----

    def _flash_error(self, widget: tk.Widget):
        try:
            widget.configure(background="#ffcccc")
            self.after(150, lambda: widget.configure(background="white"))
        except Exception:
            pass

    def _add_float_row(self, parent: ttk.LabelFrame, label: str, var: tk.StringVar, row: int):
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=8, pady=6)

        entry = tk.Entry(parent, textvariable=var)

        def validate_float(new_value: str):
            if new_value == "":
                return True
            if re.fullmatch(r"\d+(\.\d+)?", new_value) or re.fullmatch(r"\d+\.?", new_value) or re.fullmatch(r"\d*\.\d+", new_value):
                return True
            self.bell()
            self._flash_error(entry)
            return False

        vcmd = (self.register(validate_float), "%P")
        entry.configure(validate="key", validatecommand=vcmd)
        entry.grid(row=row, column=1, sticky="ew", padx=8, pady=6)
        parent.grid_columnconfigure(1, weight=1)

    def _add_coord_row(self, parent: ttk.LabelFrame, label: str, var_x: tk.StringVar, var_y: tk.StringVar, row: int):
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=8, pady=6)

        cell = ttk.Frame(parent)
        cell.grid(row=row, column=1, sticky="w", padx=8, pady=6)

        x_entry = tk.Entry(cell, width=8, textvariable=var_x)
        y_entry = tk.Entry(cell, width=8, textvariable=var_y)

        def validate_int(new_value: str, which: tk.Entry):
            if new_value == "":
                return True
            if new_value.isdigit():
                return True
            self.bell()
            self._flash_error(which)
            return False

        x_vcmd = (self.register(lambda p: validate_int(p, x_entry)), "%P")
        y_vcmd = (self.register(lambda p: validate_int(p, y_entry)), "%P")
        x_entry.configure(validate="key", validatecommand=x_vcmd)
        y_entry.configure(validate="key", validatecommand=y_vcmd)
        ttk.Label(cell, text="x").pack(side="left")
        x_entry.pack(side="left", padx=(4, 10))
        ttk.Label(cell, text="y").pack(side="left")
        y_entry.pack(side="left", padx=(4, 0))

    def _read_config_from_advanced(self) -> Config:
        # Delays
        try:
            click_delay = float(self.var_click_delay.get())
            key_delay = float(self.var_key_delay.get())
            between_users_delay = float(self.var_between_users_delay.get())
        except Exception:
            raise ValueError("Delays must be numbers (e.g. 1.0, 0.25).")

        if click_delay < 0 or key_delay < 0 or between_users_delay < 0:
            raise ValueError("Delays must be non-negative.")

        # Coordinates
        def parse_pos_int(s: str, name: str) -> int:
            if not s.isdigit():
                raise ValueError(f"{name} must be a positive whole number.")
            v = int(s)
            if v <= 0:
                raise ValueError(f"{name} must be a positive whole number.")
            return v

        assign_x = parse_pos_int(self.var_assign_x.get(), "ASSIGN_ACCESS_OFFSET x")
        assign_y = parse_pos_int(self.var_assign_y.get(), "ASSIGN_ACCESS_OFFSET y")
        tab_x = parse_pos_int(self.var_tab_x.get(), "TAB_2_CLICK_REL x")
        tab_y = parse_pos_int(self.var_tab_y.get(), "TAB_2_CLICK_REL y")
        netid_x = parse_pos_int(self.var_netid_x.get(), "NETID_FIELD_CLICK_REL x")
        netid_y = parse_pos_int(self.var_netid_y.get(), "NETID_FIELD_CLICK_REL y")

        return Config(
            click_delay=click_delay,
            key_delay=key_delay,
            between_users_delay=between_users_delay,
            assign_access_offset=(assign_x, assign_y),
            tab_2_click_rel=(tab_x, tab_y),
            netid_field_click_rel=(netid_x, netid_y),
            debug_actions=True,
        )

    # ---- Logging ----

    def _log(self, target: str, s: str):
        self.log_queue.put((target, s))

    def _drain_log_queue(self):
        try:
            while True:
                target, s = self.log_queue.get_nowait()
                if target == "main":
                    self.output_sink.append(s)
                elif target == "adv":
                    self.adv_output_sink.append(s)
                elif target == "adv_set":
                    self.adv_output_sink.set(s)
        except queue.Empty:
            pass
        self.after(50, self._drain_log_queue)

    # ---- Actions ----

    def _set_running_state(self, running: bool):
        state = "disabled" if running else "normal"
        self.btn_assign.configure(state=state)
        self.btn_coord.configure(state=state)
        # ABORT should always be enabled

    def on_abort(self):
        self.abort_event.set()
        self._log("main", "\nABORT requested. Stopping...\n")
        self._log("adv", "\nABORT requested. Stopping...\n")

    def on_assign_access(self):
        if self.worker_thread and self.worker_thread.is_alive():
            messagebox.showwarning("Busy", "A task is already running. Press ABORT to stop it.")
            return

        # Parse NetIDs (each on a newline)
        raw = self.netids_text.get("1.0", "end").splitlines()
        netids = [line.strip() for line in raw if line.strip()]

        try:
            self.cfg = self._read_config_from_advanced()
        except ValueError as e:
            messagebox.showerror("Invalid Settings", str(e))
            return

        self.abort_event.clear()
        self._set_running_state(True)

        def worker():
            try:
                run_assign_access(netids, self.cfg, self.abort_event, lambda s: self._log("main", s))
            except AbortRequested:
                self._log("main", "\nStopped (ABORT).\n")
            except Exception as e:
                self._log("main", f"\nERROR: {e}\n")
            finally:
                self.after(0, lambda: self._set_running_state(False))

        self.worker_thread = threading.Thread(target=worker, daemon=True)
        self.worker_thread.start()

    def on_coord_check(self):
        if self.worker_thread and self.worker_thread.is_alive():
            messagebox.showwarning("Busy", "A task is already running. Press ABORT to stop it.")
            return

        try:
            self.cfg = self._read_config_from_advanced()
        except ValueError as e:
            messagebox.showerror("Invalid Settings", str(e))
            return

        self.abort_event.clear()
        self._set_running_state(True)

        # Clear coord output area; coord check will live-update x/y.
        self._log("adv_set", "")

        def worker():
            try:
                run_coord_check(self.cfg, self.abort_event, lambda s: self._log("adv_set", s))
            except AbortRequested:
                # Leave last x/y displayed
                pass
            except Exception as e:
                self._log("adv_set", f"ERROR: {e}\n")
            finally:
                self.after(0, lambda: self._set_running_state(False))

        self.worker_thread = threading.Thread(target=worker, daemon=True)
        self.worker_thread.start()


if __name__ == "__main__":
    App().mainloop()
