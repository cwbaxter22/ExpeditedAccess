from __future__ import annotations

import json
import os
import queue
import re
import sys
import threading
import time
import tkinter as tk
import webbrowser
from dataclasses import dataclass
from pathlib import Path
from tkinter import messagebox, ttk

import ctypes

import keyboard
from pywinauto import Application, findwindows

import statistics


# ----------------------------
# Defaults / Configuration
# ----------------------------

APP_TITLE = "Area Access Manager"

REPO_URL = "https://github.com/cwbaxter22/ExpeditedAccess"
APP_ICON_ICO = "genie_lamp_retro.ico"


def _resource_path(relative_name: str) -> str:
    """Return an absolute path to a bundled resource.

    Works for both normal `python openSesame.py` runs and PyInstaller onefile builds.
    """

    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, relative_name)

# Set by the GUI at runtime so the automation can request a user-driven Resume.
_RESUME_HOOKS = None


@dataclass
class Config:
    click_delay: float = 0.05
    key_delay: float = 0.05
    between_users_delay: float = 0.05

    assign_access_offset: tuple[int, int] = (465, 59)
    tab_2_click_rel: tuple[int, int] = (650, 300)
    netid_field_click_rel: tuple[int, int] = (1080, 444)

    debug_actions: bool = True


DISCLAIMER_TEXT = (
    "DISCLAIMER: THIS PROGRAM WILL ONLY WORK IF AREA ACCESS MANAGER IS SCALED CORRECTLY AND OPEN.\n"
    "Sometimes, on laptops for example, Area Access Manager will show up as smaller than usual, or tabs can block each other after clicking 'Assign Access'.\n"
    "This will create a coordinate error and cause the program to fail.\n"
    "Generally, this can be corrected through corrections on the Set-Up tab."
    "The program works through a combination of moving your cursor to coordinates and clicking, and sending keyboard input.\n"
    "DO NOT MOVE THE MOUSE DURING OPERATION\n\n"
    "OPERATION:\n"
    "1) Open Area Access Manager on one display and this program on another\n"
    "It is helpful to be able to see the output log, but not required."
    "2) Confirm scaling is correct, check coordinates on Set-Up tab\n"
    "3) Enter UW NetIDs each on a new line (Shift + Enter) in the Input Box\n"
    "4) Note the location of the \"ABORT\" button incase steps 5 and 6 malfunction\n"
    "After STEP 5, do not touch the mouse unless you need to click \"ABORT\", or the program prompts you.\n"
    "5) Press \"Assign Access\"\n"
    "6) Supervise the first batch to confirm there are no errors\n"
    "If a NetID has a typo, a message will appear in the log and the program will PAUSE until a correction has been made.\n"
    "7) Once access has been assigned to all users, click the top of the \"Activate\" column on Area Access Manager to sort by newest access granted first\n"
    "8) Select the NetIDs that were just added and set specific dates and times"
)


ADVANCED_HELP_TEXT = (
    "Complete the following steps when running the program for the first time on a new device, or if clicks are missing targets:\n\n"
    "1) Open Area Access Manager\n"
    "2) Press Coord Check button\n"
    "3) Hover over the \"Assign Access\" yellow button.\n"
    "4) Note the coordinates for this cursor location and update them in the Coordinates section.\n"
    "5) Repeat this process for the UWID tab and NetID field.\n\n"
    "Another error the program can encounter is moving too quickly- i.e. faster than Area Access Manager can keep up with.\n"
    "In this situation, try increasing the time for the Click, Key, and Between-Users delays.\n"
    "Start with the Click, then Between Users, and then Key delay.\n\n"
    "These settings can be saved for future runs by pressing the Save Settings button."
)


def _settings_file_path() -> Path:
    base = os.getenv("APPDATA") or str(Path.home())
    return Path(base) / "ExpeditedAccess" / "settings.json"

# Toggle for showing the snip/ocr debug plot; enable for current debugging session.
SHOW_DEBUG_PLOT = False



def _default_config() -> Config:
    # Always reflect the in-file defaults as they exist right now.
    return Config()


def _config_to_dict(cfg: Config) -> dict:
    return {
        "click_delay": cfg.click_delay,
        "key_delay": cfg.key_delay,
        "between_users_delay": cfg.between_users_delay,
        "assign_access_offset": list(cfg.assign_access_offset),
        "tab_2_click_rel": list(cfg.tab_2_click_rel),
        "netid_field_click_rel": list(cfg.netid_field_click_rel),
    }


def _dict_to_config(d: dict) -> Config:
    defaults = _default_config()
    return Config(
        click_delay=float(d.get("click_delay", defaults.click_delay)),
        key_delay=float(d.get("key_delay", defaults.key_delay)),
        between_users_delay=float(d.get("between_users_delay", defaults.between_users_delay)),
        assign_access_offset=tuple(d.get("assign_access_offset", list(defaults.assign_access_offset))),
        tab_2_click_rel=tuple(d.get("tab_2_click_rel", list(defaults.tab_2_click_rel))),
        netid_field_click_rel=tuple(d.get("netid_field_click_rel", list(defaults.netid_field_click_rel))),
        debug_actions=True,
    )


class AbortRequested(Exception):
    pass


class _POINT(ctypes.Structure):
    _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]


_EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)


def _get_pid_for_hwnd(hwnd: int) -> int:
    pid = ctypes.c_ulong(0)
    ctypes.windll.user32.GetWindowThreadProcessId(ctypes.c_void_p(hwnd), ctypes.byref(pid))
    return int(pid.value)


def _count_visible_toplevel_windows_for_pid(pid: int) -> int:
    count = 0

    @ _EnumWindowsProc
    def enum_cb(hwnd, lparam):
        nonlocal count
        try:
            hwnd_int = int(hwnd)
            if ctypes.windll.user32.IsWindowVisible(ctypes.c_void_p(hwnd_int)) == 0:
                return True

            w_pid = ctypes.c_ulong(0)
            ctypes.windll.user32.GetWindowThreadProcessId(ctypes.c_void_p(hwnd_int), ctypes.byref(w_pid))
            if int(w_pid.value) == pid:
                count += 1
        except Exception:
            pass
        return True

    ctypes.windll.user32.EnumWindows(enum_cb, 0)
    return count


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


def _connect_single_main_window(abort_event: threading.Event, log, timeout: float = 10.0, key_delay: float = 0.05):
    """Connect to Area Access Manager even when a second window briefly exists.

    Some UI transitions can momentarily create two top-level windows with the same title,
    which causes pywinauto's title-based lookups to raise an ambiguity error. This helper
    waits until exactly one match exists, then connects by handle.
    """

    deadline = time.time() + timeout
    last_count: int | None = None

    while time.time() < deadline:
        if abort_event.is_set():
            raise AbortRequested()

        try:
            # Returns a list of ElementInfo objects (may be empty).
            elements = findwindows.find_elements(title_re=APP_TITLE)
            count = len(elements)
            if count == 1:
                handle = elements[0].handle
                app = Application(backend="win32").connect(handle=handle)
                win = app.window(handle=handle)
                return app, win

            # If there are 2 windows with the same title, it's often because a popup
            # hasn't cleared yet. Press Enter and retry until only one remains.
            if count != last_count:
                if count > 1:
                    log(f"Waiting for a single '{APP_TITLE}' window (found {count}). Pressing Enter to clear popups...\n")
                else:
                    log(f"Waiting for a single '{APP_TITLE}' window (found {count})...\n")
                last_count = count

            if count > 1:
                keyboard.press_and_release('enter')
                _sleep_or_abort(key_delay, abort_event)
        except Exception:
            # Treat lookup failures as transient; keep retrying.
            pass

        time.sleep(0.1)

    raise RuntimeError(f"Timed out waiting for a single '{APP_TITLE}' window")


def run_assign_access(netids: list[str], cfg: Config, abort_event: threading.Event, log):
    if not netids:
        log("No NetIDs provided.\n")
        return

    run_start = time.perf_counter()
    processed_count = 0

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

    _EASYOCR_READER = None

    def _detect_text_regions(
        image_input,
    ) -> list[tuple]:
        """Detect text using EasyOCR. Return list of bboxes.

        Accepts numpy array (H x W) uint8 preferred. Returns list of
        (x_min, y_min, x_max, y_max, text, confidence).
        """

        try:
            import warnings
            import easyocr  # type: ignore
            import numpy as np  # type: ignore
        except Exception as e:
            raise RuntimeError(
                "Missing EasyOCR dependency for text detection. Install with `pip install easyocr`. "
                f"Details: {e}"
            )

        # EasyOCR uses torch under the hood; suppress a common CPU-only warning.
        warnings.filterwarnings(
            "ignore",
            message=r".*pin_memory.*no accelerator is found.*",
            category=UserWarning,
        )

        # Ensure numpy array input for EasyOCR
        if not isinstance(image_input, np.ndarray):
            image_np = np.array(image_input, dtype=np.uint8)
        else:
            image_np = image_input.astype(np.uint8, copy=False)

        # Initialize reader once (on first call, downloads ~200MB model)
        nonlocal _EASYOCR_READER
        if _EASYOCR_READER is None:
            _EASYOCR_READER = easyocr.Reader(['en'], gpu=False)
        reader = _EASYOCR_READER
        results = reader.readtext(image_np)

        # results is a list of: (bbox_coords, text, confidence)
        # bbox_coords is: [[x1,y1], [x2,y2], [x3,y3], [x4,y4]] (4 corners)
        detections = []

        for (bbox_coords, text, confidence) in results:
            xs = [pt[0] for pt in bbox_coords]
            ys = [pt[1] for pt in bbox_coords]
            x_min, x_max = min(xs), max(xs)
            y_min, y_max = min(ys), max(ys)
            detections.append((int(x_min), int(y_min), int(x_max), int(y_max), text, confidence))

        return detections

    def _evaluate_assignment_screen(detections: list[tuple]) -> tuple[str, list[tuple], dict]:
        """Evaluate OCR detections to decide single vs multiple user.

        Returns (status, bboxes_used, constraint_bounds) where status is one of:
        - 'title_missing'
        - 'label_missing'
        - 'single_user'
        - 'multiple_users'

        Uses word count: 2 words = 1 user, >2 words = multiple users.
        constraint_bounds is a dict with 'label_y', 'label_x_min', 'y_max', 'x_max' if applicable, else None.
        """

        def _norm(s: str) -> str:
            return " ".join((s or "").lower().strip().split())

        def _is_title(s: str) -> bool:
            t = _norm(s)
            if "select the people you" in t:
                return True
            if "access level assignment wizard" in t:
                return True
            if "access level assignment" in t:
                return True
            # Tolerate minor OCR variation: require core keywords.
            words = set(t.replace("-", " ").split())
            has_assign = any(w.startswith("assign") for w in words)
            has_wiz = any(w.startswith("wiz") for w in words)
            return ("access" in words) and ("level" in words) and has_assign and has_wiz

        # Locate title
        title_y = None
        for x_min, y_min, x_max, y_max, text, conf in detections:
            if _is_title(text):
                title_y = y_max
                break

        if title_y is None:
            return "title_missing", [], None

        # Find Last Name (preferred) or First Name below title
        label_y = None
        label_x_min = None
        label_bbox = None
        below_title = [d for d in detections if d[1] > title_y]
        below_title.sort(key=lambda d: d[1])

        def _is_label(s: str, kind: str) -> bool:
            t = _norm(s)
            return kind in t

        for x_min, y_min, x_max, y_max, text, conf in below_title:
            if _is_label(text, "last name"):
                label_y = y_max
                label_x_min = x_min
                label_bbox = (x_min, y_min, x_max, y_max, text, conf)
                break

        if label_y is None:
            for x_min, y_min, x_max, y_max, text, conf in below_title:
                if _is_label(text, "first name"):
                    label_y = y_max
                    label_x_min = x_min
                    label_bbox = (x_min, y_min, x_max, y_max, text, conf)
                    break

        if label_y is None:
            return "label_missing", [], None

        # Crop: remove everything with x-value less than label_x_min
        cropped_detections = [d for d in detections if d[0] >= label_x_min]

        # Rows below label: only search within 100px below label_y
        # Shift constraint by 5 pixels on Y-axis only
        constraint_y_max = label_y + 100 + 5
        constraint_x_max = label_x_min + 100
        constraint_bounds = {
            "label_y": label_y + 5,
            "label_x_min": label_x_min,
            "y_max": constraint_y_max,
            "x_max": constraint_x_max,
        }

        below = [
            d for d in cropped_detections
            if d[1] > label_y and d[1] <= constraint_y_max and d[0] <= constraint_x_max
        ]

        if not below:
            return "single_user", [label_bbox], constraint_bounds

        # Count words ONLY within the bounding box
        word_count = sum(len(d[4].split()) for d in below)

        if cfg.debug_actions:
            counted_text = " | ".join([d[4] for d in below])
            log_action(f"OCR word_count_in_box={word_count}; texts={counted_text}")

        # <= 2 words = 1 user (e.g., "Last Name"), > 2 words = multiple users
        status = "multiple_users" if word_count > 2 else "single_user"
        used = [label_bbox] + below
        return status, used, constraint_bounds

    def _grab_grayscale_intensity_array(
        bbox: tuple[int, int, int, int],
    ) -> tuple[list[list[int]], dict[str, float]]:
        """Grab a screen region and return (2D intensities, summary stats).

        Intensities are 0..255 (0=black, 255=white).
        `bbox` is (left, top, right, bottom) in screen coordinates.
        """

        try:
            from PIL import ImageGrab  # type: ignore
        except Exception as e:
            raise RuntimeError(
                "Missing Pillow dependency for screen capture. Install with `pip install pillow`. "
                f"Details: {e}"
            )

        if bbox == (0, 0, 0, 0):
            img = ImageGrab.grab().convert("L")
        else:
            img = ImageGrab.grab(bbox).convert("L")
        w, h = img.size

        # Use get_flattened_data when available to avoid Pillow deprecation warning.
        if hasattr(img, "get_flattened_data"):
            flat_seq = img.get_flattened_data()
        else:
            flat_seq = img.getdata()

        flat = list(flat_seq)
        # Reshape into row-major 2D list for easy indexing in hover callback.
        arr2d = [flat[i * w:(i + 1) * w] for i in range(h)]

        mean_v = float(sum(flat) / len(flat))
        median_v = float(statistics.median(flat))
        std_v = float(statistics.pstdev(flat))
        stats = {
            "mean": mean_v,
            "median": median_v,
            "std": std_v,
            "min": float(min(flat)),
            "max": float(max(flat)),
        }
        return arr2d, stats

    def _show_snip_debug_plot(
        arr2d: list[list[int]],
        stats: dict[str, float],
        *,
        netid: str,
        bbox: tuple[int, int, int, int],
        detections: list = None,
        status: str = "unknown",
        constraint_bounds: dict = None,
    ):
        """Show a hoverable Matplotlib plot of the snipped region with OCR bboxes and search constraint."""

        if not SHOW_DEBUG_PLOT:
            return

        try:
            import matplotlib.pyplot as plt  # type: ignore
            from matplotlib.gridspec import GridSpec  # type: ignore
            from matplotlib.patches import Rectangle  # type: ignore
            import warnings
        except Exception as e:
            raise RuntimeError(
                "Missing Matplotlib dependency for debug plots. Install with `pip install matplotlib`. "
                f"Details: {e}"
            )

        # Suppress main-thread GUI warning (we're intentionally in a worker thread).
        warnings.filterwarnings("ignore", message="Starting a Matplotlib GUI outside of the main thread")

        h = len(arr2d)
        w = len(arr2d[0]) if h else 0

        fig = plt.figure(figsize=(10, 5))
        gs = GridSpec(1, 2, width_ratios=[5, 2], figure=fig)
        ax = fig.add_subplot(gs[0, 0])
        axr = fig.add_subplot(gs[0, 1])

        im = ax.imshow(arr2d, cmap="gray", vmin=0, vmax=255, interpolation="nearest")
        ax.set_title(f"NetID snip debug: {netid}\nStatus: {status}")
        ax.set_xlabel("x")
        ax.set_ylabel("y")

        # Draw yellow constraint box if available
        if constraint_bounds:
            label_y = constraint_bounds.get("label_y", 0)
            label_x_min = constraint_bounds.get("label_x_min", 0)
            y_max = constraint_bounds.get("y_max", 0)
            x_max = constraint_bounds.get("x_max", 0)
            rect_constraint = Rectangle(
                (label_x_min, label_y),
                x_max - label_x_min,
                y_max - label_y,
                linewidth=3,
                edgecolor="yellow",
                facecolor="none",
            )
            ax.add_patch(rect_constraint)

        # Draw bounding boxes from OCR detections
        if detections:
            for x_min, y_min, x_max, y_max, text, conf in detections:
                rect = Rectangle(
                    (x_min, y_min),
                    x_max - x_min,
                    y_max - y_min,
                    linewidth=2,
                    edgecolor="red",
                    facecolor="none",
                )
                ax.add_patch(rect)
                ax.text(x_min, y_min - 5, f"{text[:15]}({conf:.2f})", color="red", fontsize=8)

        try:
            fig.colorbar(im, ax=ax, fraction=0.046, pad=0.04)
        except Exception:
            pass

        axr.axis("off")
        stats_text = (
            "Intensity stats (0..255)\n"
            f"mean:   {stats['mean']:.2f}\n"
            f"median: {stats['median']:.2f}\n"
            f"std:    {stats['std']:.2f}\n"
            f"min:    {stats['min']:.0f}\n"
            f"max:    {stats['max']:.0f}\n\n"
            f"Status: {status}\n\n"
            "Yellow box = search area\n"
            "Red boxes = detected text\n"
            "Close to continue"
        )
        axr_text = axr.text(0.0, 1.0, stats_text, va="top", ha="left", family="monospace")
        hover_text = axr.text(0.0, 0.05, "hover: (x=?, y=?)\nI=?", va="bottom", ha="left", family="monospace")

        def on_move(event):
            if event.inaxes is not ax or event.xdata is None or event.ydata is None:
                return
            x = int(event.xdata + 0.5)
            y = int(event.ydata + 0.5)
            if 0 <= x < w and 0 <= y < h:
                val = arr2d[y][x]
                hover_text.set_text(f"hover: (x={x}, y={y})\nI={val}")
                fig.canvas.draw_idle()

        fig.canvas.mpl_connect("motion_notify_event", on_move)
        plt.tight_layout()
        plt.show(block=True)

    log("Connecting to application...\n")
    log("(Press ABORT anytime to stop)\n\n")

    app_win32, main_win32 = _connect_single_main_window(abort_event, log, timeout=15.0, key_delay=cfg.key_delay)
    main_win32.wait("ready", timeout=10)
    log(f"Connected: {main_win32.window_text()}\n")

    def click_main(coords: tuple[int, int], label: str):
        if abort_event.is_set():
            raise AbortRequested()

        if label == "Assign Access":
            log("\n" + ("*" * 56) + "\n")
            log("********************  ASSIGN ACCESS CLICK  ********************\n")
            log(("*" * 56) + "\n")
        try:
            rect = main_win32.rectangle()
            approx_screen = (rect.left + coords[0], rect.top + coords[1])
            log_action(f"CLICK {label} @ {coords} (approx screen {approx_screen})")
        except Exception:
            log_action(f"CLICK {label} @ {coords}")
        main_win32.click_input(coords=coords)

    main_pid = _get_pid_for_hwnd(int(main_win32.handle))

    for idx, netid in enumerate(netids):
        if abort_event.is_set():
            raise AbortRequested()

        # Confirm there is only one Area Access Manager window present.
        # If not, press Enter repeatedly (up to 15s) to clear any lingering dialog.
        app_win32, main_win32 = _connect_single_main_window(abort_event, log, timeout=15.0, key_delay=cfg.key_delay)
        main_win32.wait("ready", timeout=10)
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

        # Click UWID tab — coordinate-based (explicitly required)
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

        # After Enter, evaluate full-screen OCR for ambiguity (multiple users).
        # Keep looping until we see the title; then decide single vs multiple users.
        import numpy as np  # type: ignore
        ocr_status = "title_missing"
        detections: list[tuple] = []
        first_attempt = True

        while True:
            if abort_event.is_set():
                raise AbortRequested()

            # Log status before capturing
            if ocr_status == "title_missing":
                if first_attempt:
                    log("Imaging\n")
                    first_attempt = False
                else:
                    log("Re-imaging\n")
            else:
                log("Analyzing\n")

            _sleep_or_abort(cfg.click_delay * 10, abort_event)
            # Snip only the top half of the screen (faster than full-screen OCR)
            import ctypes
            user32 = ctypes.windll.user32
            screen_w = int(user32.GetSystemMetrics(0))
            screen_h = int(user32.GetSystemMetrics(1))
            top_half_bbox = (0, 0, screen_w, max(1, screen_h // 2))
            full_arr2d, full_stats = _grab_grayscale_intensity_array(top_half_bbox)
            try:
                img_array = np.array(full_arr2d, dtype=np.uint8)
                detections = _detect_text_regions(img_array)
                ocr_status, used_boxes, constraint_bounds = _evaluate_assignment_screen(detections)
            except Exception as ocr_err:
                log(f"[OCR] Warning: {ocr_err}\n")
                ocr_status = "ocr_failed"
                used_boxes = []
                constraint_bounds = None

            if ocr_status == "title_missing":
                continue
            if ocr_status == "label_missing":
                # Keep trying; UI may still be rendering.
                continue

            # single_user or multiple_users or ocr_failed
            _show_snip_debug_plot(full_arr2d, full_stats, netid=netid, bbox=(0, 0, 0, 0), detections=used_boxes, status=ocr_status, constraint_bounds=constraint_bounds)

            if ocr_status == "multiple_users":
                if _RESUME_HOOKS is None:
                    raise RuntimeError("Resume requested but GUI hooks not initialized")

                log(
                    "\nPAUSED: More than one user detected.\n"
                    "Please select the correct user in Area Access Manager, press Enter there, then press Resume here.\n"
                    "(No keys/clicks will be sent while paused.)\n\n"
                )

                _RESUME_HOOKS['resume_event'].clear()
                _RESUME_HOOKS['enable_resume']()

                while not _RESUME_HOOKS['resume_event'].is_set():
                    if abort_event.is_set():
                        raise AbortRequested()
                    time.sleep(0.1)

                _RESUME_HOOKS['disable_resume']()
                # RESUME clicks shift focus to the GUI. Re-focus Area Access Manager
                # so the remaining keyboard automation goes to the correct window.
                try:
                    wizard_window.set_focus()
                except Exception:
                    try:
                        main_win32.set_focus()
                    except Exception:
                        pass
                _sleep_or_abort(0.2, abort_event)
                break  # user handled selection; proceed

            elif ocr_status == "single_user":
                # Proceed automatically
                break
            else:
                # ocr_failed: proceed without pausing
                break

        # Ensure keyboard focus is back on Area Access Manager before continuing.
        try:
            wizard_window.set_focus()
        except Exception:
            try:
                main_win32.set_focus()
            except Exception:
                pass
        _sleep_or_abort(0.1, abort_event)

        # Validate NetID using window-count rule:
        # After pressing Enter, if NetID is invalid AAM shows a popup, yielding 3 windows
        # for the AAM process instead of 2. Proceed only if count == 2.
        def popup_detected_after_enter(observe_seconds: float) -> bool:
            # IMPORTANT: don't short-circuit when we see 2 windows immediately.
            # The popup can appear a fraction of a second later.
            deadline = time.time() + max(0.0, observe_seconds)
            max_count = -1
            last_count = None

            while time.time() < deadline:
                if abort_event.is_set():
                    raise AbortRequested()
                c = _count_visible_toplevel_windows_for_pid(main_pid)
                last_count = c
                max_count = max(max_count, c)
                if c > 2:
                    return True
                time.sleep(0.05)

            if cfg.debug_actions and last_count is not None:
                log_action(f"AAM windows after Enter: last={last_count}, max={max_count}")
            return False

        def wait_for_popup_and_dismiss(timeout_seconds: float, label: str) -> bool:
            """Wait for an AAM popup (extra top-level window) then dismiss with Enter.

            Returns True if a popup was observed and dismissed, else False.
            """
            deadline = time.time() + max(0.0, timeout_seconds)
            saw_popup = False

            def _try_dismiss_extra_aam_window() -> bool:
                """If an extra AAM dialog exists, focus it and click OK (fallback: Enter)."""
                try:
                    elements = findwindows.find_elements(title_re=APP_TITLE)
                except Exception:
                    return False

                def _safe_handle(win) -> int:
                    try:
                        return int(getattr(win, "handle"))
                    except Exception:
                        return 0

                main_handle = _safe_handle(main_win32)
                wizard_handle = _safe_handle(wizard_window)

                for el in elements:
                    try:
                        h = int(el.handle)
                    except Exception:
                        continue

                    if h in (0, main_handle, wizard_handle):
                        continue

                    try:
                        dlg_app = Application(backend="win32").connect(handle=h)
                        dlg = dlg_app.window(handle=h)
                        try:
                            dlg.set_focus()
                        except Exception:
                            pass

                        # Prefer clicking OK (some dialogs ignore Enter if focus isn't right).
                        try:
                            ok_btn = dlg.child_window(title_re=r"^OK$", class_name="Button")
                            if ok_btn.exists(timeout=0.2):
                                ok_btn.click_input()
                                return True
                        except Exception:
                            pass

                        # Fallback: send Enter to dismiss.
                        press('enter')
                        return True
                    except Exception:
                        continue

                return False

            while time.time() < deadline:
                if abort_event.is_set():
                    raise AbortRequested()

                # Some popups don't reliably show up in the PID window-count heuristic.
                # Try direct window discovery first.
                if _try_dismiss_extra_aam_window():
                    saw_popup = True
                    return True

                c = _count_visible_toplevel_windows_for_pid(main_pid)
                if c > 2:
                    saw_popup = True
                    if cfg.debug_actions:
                        log_action(f"Popup detected ({label}); windows={c}; dismissing with Enter")
                    try:
                        wizard_window.set_focus()
                    except Exception:
                        try:
                            main_win32.set_focus()
                        except Exception:
                            pass
                    press('enter')

                    # Wait for popup to clear.
                    clear_deadline = time.time() + 5.0
                    while time.time() < clear_deadline:
                        if abort_event.is_set():
                            raise AbortRequested()
                        if _count_visible_toplevel_windows_for_pid(main_pid) <= 2:
                            return True
                        time.sleep(0.05)

                    return True

                time.sleep(0.05)

            return saw_popup

        # Start the exact key sequence here.
        # For multiple-user, the user already pressed the first Enter manually before hitting Resume.
        if ocr_status != "multiple_users":
            log("Enter\n")
            press('enter')

        _sleep_or_abort(cfg.key_delay, abort_event)
        if popup_detected_after_enter(1.5):
            if _RESUME_HOOKS is None:
                raise RuntimeError("Resume requested but GUI hooks not initialized")

            while True:
                if abort_event.is_set():
                    raise AbortRequested()

                current_count = _count_visible_toplevel_windows_for_pid(main_pid)
                log(
                    "\nPAUSED: NetID may be invalid (see pop-up on Area Access Manager).\n"
                    "Correct the NetID typo, \npress 'Next' in Area Access Manager to confirm typo is fixed, \nand then press 'Resume' Here.\n"
                    f"(Debug: observed {current_count} Area Access Manager window(s).)\n"
                    "(No keys/clicks will be sent while paused.)\n\n"
                )

                _RESUME_HOOKS['resume_event'].clear()
                _RESUME_HOOKS['enable_resume']()

                while not _RESUME_HOOKS['resume_event'].is_set():
                    if abort_event.is_set():
                        raise AbortRequested()
                    time.sleep(0.1)

                _RESUME_HOOKS['disable_resume']()

                # User-driven recovery: do not press Enter automatically.
                # We only proceed once the popup is gone (no extra window).
                # Keep checking until count is back to 2.
                deadline = time.time() + 15.0
                while time.time() < deadline:
                    if abort_event.is_set():
                        raise AbortRequested()
                    if _count_visible_toplevel_windows_for_pid(main_pid) == 2:
                        break
                    time.sleep(0.1)

                if _count_visible_toplevel_windows_for_pid(main_pid) == 2:
                    # RESUME clicks shift focus to the GUI. Re-focus the wizard so the
                    # remaining keyboard automation goes to Area Access Manager.
                    try:
                        wizard_window.set_focus()
                    except Exception:
                        try:
                            main_win32.set_focus()
                        except Exception:
                            pass
                    _sleep_or_abort(0.2, abort_event)
                    break

        if ocr_status == "single_user":
            # (First Enter already sent above, before popup detection)

            log("Tab\n")
            press('tab')
            _sleep_or_abort(cfg.key_delay, abort_event)

            log("Enter\n")
            press('enter')
            _sleep_or_abort(cfg.key_delay, abort_event)

            log("Enter (popup)\n")
            press('enter')
            _sleep_or_abort(cfg.key_delay, abort_event)

            # Tab x4 then Enter
            log("Tab x4 + Enter\n")
            for _ in range(4):
                press('tab')
                _sleep_or_abort(cfg.key_delay, abort_event)
            press('enter')
            _sleep_or_abort(cfg.key_delay, abort_event)

            # Enter
            log("Enter\n")
            press('enter')
            _sleep_or_abort(cfg.key_delay, abort_event)

            # Enter
            log("Enter\n")
            press('enter')
            _sleep_or_abort(cfg.key_delay, abort_event)

            # Wait key delay * 10
            log("Waiting (key_delay * 10)\n")
            _sleep_or_abort(cfg.key_delay * 10, abort_event)

            # Final Enter to trigger the final confirmation dialog (if any)
            log("Enter\n")
            press('enter')
            _sleep_or_abort(cfg.key_delay, abort_event)

            
        else:
            # Match the single-user workflow from this point forward.
            # (First Enter is skipped only for multiple_users; it was handled manually.)

            # Tab
            log("Tab\n")
            press('tab')
            _sleep_or_abort(cfg.key_delay, abort_event)

            # Enter
            log("Enter\n")
            press('enter')
            _sleep_or_abort(cfg.key_delay, abort_event)

            # New window pops up; OK with Enter
            log("Enter (popup)\n")
            press('enter')
            _sleep_or_abort(cfg.key_delay, abort_event)

            # Tab x4 then Enter
            log("Tab x4 + Enter\n")
            for _ in range(4):
                press('tab')
                _sleep_or_abort(cfg.key_delay, abort_event)
            press('enter')
            _sleep_or_abort(cfg.key_delay, abort_event)

            # Enter
            log("Enter\n")
            press('enter')
            _sleep_or_abort(cfg.key_delay, abort_event)

            # Enter
            log("Enter\n")
            press('enter')
            _sleep_or_abort(cfg.key_delay, abort_event)

            # Wait key delay * 10
            log("Waiting (key_delay * 10)\n")
            _sleep_or_abort(cfg.key_delay * 10, abort_event)

            # Enter
            log("Enter\n")
            press('enter')

        log(f"Completed {netid}\n\n")
        processed_count += 1

    elapsed = time.perf_counter() - run_start
    log("All users processed.\n")
    log(f"Total users processed: {processed_count}\n")
    log(f"Total processing time: {elapsed:.1f} seconds\n")


def run_coord_check(cfg: Config, abort_event: threading.Event, log):
    app_win32, main_win32 = _connect_single_main_window(abort_event, log, timeout=15.0, key_delay=cfg.key_delay)
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
        self.title("Open Sesame")
        self.geometry("900x650")

        # Set window/taskbar icon (Windows requires .ico).
        try:
            self.iconbitmap(_resource_path(APP_ICON_ICO))
        except Exception:
            # If the icon isn't found (or in unsupported environments), keep default.
            pass

        self.settings_path = _settings_file_path()
        self.cfg = _default_config()
        self.show_setup_reminder = True
        self._load_persisted_settings()
        self.abort_event = threading.Event()
        self.resume_event = threading.Event()
        self.worker_thread: threading.Thread | None = None
        self.current_task: str | None = None  # 'assign' | 'coord' | None

        self.log_queue: queue.Queue[tuple[str, str]] = queue.Queue()

        self._build_ui()
        self._apply_cfg_to_vars(self.cfg)
        self.after(50, self._drain_log_queue)

        # First-run / reminder popup
        if self.show_setup_reminder:
            self.after(200, self._show_setup_reminder_dialog)

    # ---- UI ----

    def _build_ui(self):
        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True)

        self.tab_input = ttk.Frame(notebook)
        self.tab_adv = ttk.Frame(notebook)
        notebook.add(self.tab_input, text="Input")
        notebook.add(self.tab_adv, text="Set-up")

        self._build_input_tab(self.tab_input)
        self._build_advanced_tab(self.tab_adv)

    def _add_repo_footer(self, parent: ttk.Frame):
        footer = ttk.Frame(parent)
        footer.pack(side="bottom", fill="x", padx=10, pady=(0, 10))

        ttk.Label(footer, text="Software Repository:").pack(side="left")

        link = tk.Label(
            footer,
            text=REPO_URL,
            fg="blue",
            cursor="hand2",
            font=("TkDefaultFont", 9, "underline"),
        )
        link.pack(side="left", padx=(6, 0))
        link.bind("<Button-1>", lambda _e: self._open_repo_url())

    def _open_repo_url(self):
        try:
            webbrowser.open_new_tab(REPO_URL)
        except Exception:
            messagebox.showerror("Error", f"Could not open URL:\n{REPO_URL}")

    def _build_input_tab(self, parent: ttk.Frame):
        # NetIDs input
        ttk.Label(parent, text="NetIDs - Each on a Newline (Shift + Enter)").pack(anchor="w", padx=10, pady=(10, 0))
        self.netids_text = tk.Text(parent, height=6)
        self.netids_text.pack(fill="x", padx=10, pady=5)

        # Buttons row
        btn_row = ttk.Frame(parent)
        btn_row.pack(fill="x", padx=10, pady=5)

        self.btn_assign = ttk.Button(btn_row, text="Assign Access", command=self.on_assign_access)
        self.btn_assign.pack(side="left")

        self.btn_resume = ttk.Button(btn_row, text="RESUME", command=self.on_resume, state="disabled")
        self.btn_resume.pack(side="left", padx=(8, 0))

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
        self._add_repo_footer(parent)

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

        self.btn_coord = ttk.Button(btn_row, text="Start/Stop Coord Check", command=self.on_coord_check)
        self.btn_coord.pack(side="left")

        self.btn_save_settings = ttk.Button(btn_row, text="Save Settings", command=self.on_save_settings)
        self.btn_save_settings.pack(side="left", padx=(8, 0))

        self.btn_restore_defaults = ttk.Button(btn_row, text="Restore Defaults", command=self.on_restore_defaults)
        self.btn_restore_defaults.pack(side="left", padx=(8, 0))

        ttk.Button(btn_row, text="Set-up Help", command=lambda: messagebox.showinfo("Set-up Help", ADVANCED_HELP_TEXT)).pack(
            side="left", padx=(8, 0)
        )

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
        self._add_repo_footer(parent)

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

    def _apply_cfg_to_vars(self, cfg: Config):
        # Delays
        self.var_click_delay.set(str(cfg.click_delay))
        self.var_key_delay.set(str(cfg.key_delay))
        self.var_between_users_delay.set(str(cfg.between_users_delay))

        # Coordinates
        self.var_assign_x.set(str(cfg.assign_access_offset[0]))
        self.var_assign_y.set(str(cfg.assign_access_offset[1]))
        self.var_tab_x.set(str(cfg.tab_2_click_rel[0]))
        self.var_tab_y.set(str(cfg.tab_2_click_rel[1]))
        self.var_netid_x.set(str(cfg.netid_field_click_rel[0]))
        self.var_netid_y.set(str(cfg.netid_field_click_rel[1]))

    def _load_persisted_settings(self):
        try:
            if not self.settings_path.exists():
                return
            data = json.loads(self.settings_path.read_text(encoding="utf-8"))
            cfg_data = data.get("config") or {}
            self.cfg = _dict_to_config(cfg_data)
            ui_data = data.get("ui") or {}
            self.show_setup_reminder = bool(ui_data.get("show_setup_reminder", True))
        except Exception:
            # Corrupt/invalid settings should not prevent launching.
            self.cfg = _default_config()
            self.show_setup_reminder = True

    def _save_persisted_settings(self, cfg: Config | None = None, show_setup_reminder: bool | None = None):
        try:
            data = {}
            if self.settings_path.exists():
                try:
                    data = json.loads(self.settings_path.read_text(encoding="utf-8"))
                except Exception:
                    data = {}

            if cfg is not None:
                data["config"] = _config_to_dict(cfg)

            if show_setup_reminder is not None:
                data.setdefault("ui", {})
                data["ui"]["show_setup_reminder"] = bool(show_setup_reminder)

            self.settings_path.parent.mkdir(parents=True, exist_ok=True)
            self.settings_path.write_text(json.dumps(data, indent=2), encoding="utf-8")
        except Exception:
            # Saving is best-effort.
            pass

    def _show_setup_reminder_dialog(self):
        dlg = tk.Toplevel(self)
        dlg.title("Set-up Reminder")
        dlg.transient(self)
        dlg.grab_set()
        dlg.resizable(False, False)

        msg = (
            "Before proceeding, please check the Set-up tab to confirm the delays and coordinates are correct.\n\n"
            "Tip: Use 'Coord Check' if clicks are missing targets."
        )

        ttk.Label(dlg, text=msg, justify="left", wraplength=520).pack(padx=14, pady=(14, 8), anchor="w")

        dont_show_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(dlg, text="Don't show this message again", variable=dont_show_var).pack(padx=14, pady=(0, 12), anchor="w")

        btn_row = ttk.Frame(dlg)
        btn_row.pack(fill="x", padx=14, pady=(0, 14))

        def on_ok():
            if dont_show_var.get():
                self.show_setup_reminder = False
                self._save_persisted_settings(show_setup_reminder=False)
            dlg.destroy()

        ttk.Button(btn_row, text="OK", command=on_ok).pack(side="right")

        # Center dialog over the main window
        self.update_idletasks()
        try:
            x = self.winfo_rootx() + (self.winfo_width() // 2) - 250
            y = self.winfo_rooty() + (self.winfo_height() // 2) - 120
            dlg.geometry(f"+{max(0, x)}+{max(0, y)}")
        except Exception:
            pass

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

        # Coord check uses the same button to stop.
        if running and self.current_task == "coord":
            self.btn_coord.configure(state="normal")
        else:
            self.btn_coord.configure(state=state)

        self.btn_save_settings.configure(state=state)
        self.btn_restore_defaults.configure(state=state)
        # Resume only enabled when a pause is requested
        if not running:
            self.btn_resume.configure(state="disabled")
        # ABORT should always be enabled

    def _set_resume_state(self, enabled: bool):
        self.btn_resume.configure(state=("normal" if enabled else "disabled"))

    def on_abort(self):
        self.abort_event.set()
        # If we're paused waiting on Resume, unblock it.
        self.resume_event.set()
        self._log("main", "\nABORT requested. Stopping...\n")
        self._log("adv", "\nABORT requested. Stopping...\n")

    def on_resume(self):
        self.resume_event.set()
        self._set_resume_state(False)
        self._log("main", "\nRESUME pressed.\n")

    def on_save_settings(self):
        try:
            cfg = self._read_config_from_advanced()
        except ValueError as e:
            messagebox.showerror("Invalid Settings", str(e))
            return

        self.cfg = cfg
        self._save_persisted_settings(cfg=cfg)
        self._log("adv", f"\nSaved settings to: {self.settings_path}\n")
        self._log("main", "\nSettings saved.\n")

    def on_restore_defaults(self):
        defaults = _default_config()
        self.cfg = defaults
        self._apply_cfg_to_vars(defaults)
        self._save_persisted_settings(cfg=defaults)
        self._log("adv", f"\nRestored defaults and saved to: {self.settings_path}\n")
        self._log("main", "\nDefaults restored and saved.\n")

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
        self.resume_event.clear()
        self.current_task = "assign"
        self._set_running_state(True)

        # Expose resume hooks to the automation engine
        global _RESUME_HOOKS
        _RESUME_HOOKS = {
            'resume_event': self.resume_event,
            'enable_resume': lambda: self.after(0, lambda: self._set_resume_state(True)),
            'disable_resume': lambda: self.after(0, lambda: self._set_resume_state(False)),
        }

        def worker():
            try:
                run_assign_access(netids, self.cfg, self.abort_event, lambda s: self._log("main", s))
            except AbortRequested:
                self._log("main", "\nStopped (ABORT).\n")
            except Exception as e:
                self._log("main", f"\nERROR: {e}\n")
            finally:
                # Clear hooks when done
                global _RESUME_HOOKS
                _RESUME_HOOKS = None
                self.current_task = None
                self.after(0, lambda: self._set_running_state(False))

        self.worker_thread = threading.Thread(target=worker, daemon=True)
        self.worker_thread.start()

    def on_coord_check(self):
        if self.worker_thread and self.worker_thread.is_alive():
            if self.current_task == "coord":
                # Toggle stop
                self.abort_event.set()
                self._log("adv", "\nStopping coord check...\n")
                return
            messagebox.showwarning("Busy", "A task is already running. Press ABORT to stop it.")
            return

        try:
            self.cfg = self._read_config_from_advanced()
        except ValueError as e:
            messagebox.showerror("Invalid Settings", str(e))
            return

        self.abort_event.clear()
        self.current_task = "coord"
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
                self.current_task = None
                self.after(0, lambda: self._set_running_state(False))

        self.worker_thread = threading.Thread(target=worker, daemon=True)
        self.worker_thread.start()


if __name__ == "__main__":
    App().mainloop()
