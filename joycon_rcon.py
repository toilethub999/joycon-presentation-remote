#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
PowerPoint Remote – Joy‑Con controlled laser pointer
Author: T
"""

import sys
import time
from pyjoycon import JoyCon, get_R_id, GyroTrackingJoyCon
import pyautogui
import win32com.client
import pythoncom
import logging

# ------------------------------------------------------------
# Configuration – tweak these if you like
# ------------------------------------------------------------
MOVE_SPEED = 1500          # how fast the mouse pointer moves
MODE_LASER = 0
MODE_MOUSE = 1
# ------------------------------------------------------------

logging.basicConfig(
    level=logging.INFO,
    format="[%(asctime)s] %(levelname)s – %(message)s",
    datefmt="%H:%M:%S",
)

# ------------------------------------------------------------
# 1)  Find the Joy‑Con and initialise gyroscope
# ------------------------------------------------------------
try:
    joycon_id = get_R_id()
    joycon = JoyCon(*joycon_id)
    time.sleep(1)                                   # ensure connection is established, wait for ACK
    joycon_gyro = GyroTrackingJoyCon(*joycon_id)
    joycon_gyro.reset_orientation()

    state = joycon.get_status()
    pre_pos_x = joycon_gyro.pointer[0]
    pre_pos_y = -joycon_gyro.pointer[1]             # note the sign flip

    # keep track of button states to detect a *new* press
    prev = {
        "x": 0, "b": 0, "y": 0, "a": 0,
        "sr": 0, "plus": 0, "home": 0,
        "r": 0, "zr": 0,
    }
    mode = MODE_LASER
except Exception as e:
    logging.error("Could not initialise Joy‑Con: %s", e)
    sys.exit(1)

# ------------------------------------------------------------
# 2)  Connect to the *running* PowerPoint instance
# ------------------------------------------------------------
try:
    ppt = pythoncom.GetActiveObject("PowerPoint.Application")
    logging.info("Using existing PowerPoint instance")
except Exception:
    # PowerPoint not running – start a new one
    ppt = win32com.client.Dispatch("PowerPoint.Application")
    logging.info("Started new PowerPoint instance")

# Grab the SlideShow window – we expect that the user is already
# running a slide‑show (F5).  If not, we will get an exception.
try:
    ss_view = ppt.ActivePresentation.SlideShowWindows(1).View
    logging.info("Slide‑Show window found – we can toggle tools")
except Exception:
    logging.error(
        "No SlideShow window detected.  Start a presentation with F5 "
        "before using the remote."
    )
    sys.exit(1)

# ------------------------------------------------------------
# 3)  Helper functions to *toggle* the built‑in tools
# ------------------------------------------------------------
def laser_on():
    """Turn the PowerPoint laser pointer on (the L key toggles it)."""
    logging.info("Laser pointer ON")
    pyautogui.press("l")  # keycode 'L' toggles the laser

def laser_off():
    """Turn the PowerPoint laser pointer off."""
    logging.info("Laser pointer OFF")
    pyautogui.press("l")  # press L again

def highlighter_on():
    """Start the PowerPoint highlighter (yellow annotation pen)."""
    logging.info("Highlighter ON")
    ss_view.StartHighlighter(ppt.constants.msoColorAutomatic)

def highlighter_off():
    """End any current annotation (including highlighter)."""
    logging.info("Highlighter OFF")
    ss_view.EndAnnotation()

def toggle_mode():
    """Called whenever SR is pressed – switches between laser & highlighter."""
    global mode, ss_view
    if mode == MODE_LASER:
        # switch to highlighter
        laser_off()          # make sure laser is not shown
        highlighter_on()
        mode = MODE_MOUSE
    else:
        # switch to laser
        highlighter_off()    # finish the annotation first
        laser_on()
        mode = MODE_LASER

# ------------------------------------------------------------
# 3)  Main loop – read the joystick, move the cursor, toggle tools
# ------------------------------------------------------------
logging.info("Remote is now running.  Press HOME on the Joy‑Con to exit.")
while True:
    try:
        state = joycon.get_status()
    except Exception as e:
        logging.warning("Lost Joy‑Con connection: %s", e)
        time.sleep(0.5)
        continue

    # ------------------------------------------------------------
    # 3a)  Gyroscope → mouse movement
    # ------------------------------------------------------------
    gyro = joycon_gyro.pointer
    cur_x, cur_y = gyro[0], -gyro[1]  # note the sign flip
    dx, dy = cur_x - pre_pos_x, cur_y - pre_pos_y
    if state["right"] or state["zr"]:          # only move when pressed
        pyautogui.moveRel(dx * MOVE_SPEED, dy * MOVE_SPEED, duration=0)
    pre_pos_x, pre_pos_y = cur_x, cur_y

    # ------------------------------------------------------------
    # 3b)  Button handling – detect *new* presses
    # ------------------------------------------------------------
    #  X  – page‑up
    if state["x"] and not prev["x"]:
        logging.info("Page‑up")
        pyautogui.press("pageup")
    #  B  – page‑down
    if state["b"] and not prev["b"]:
        logging.info("Page‑down")
        pyautogui.press("pagedown")
    #  Y  – left click
    if state["y"] and not prev["y"]:
        logging.info("Left click")
        pyautogui.click()
    #  A  – right click
    if state["a"] and not prev["a"]:
        logging.info("Right click")
        pyautogui.click(button="right")

    #  SR – toggle mode (laser ↔ highlighter)
    if state["sr"] and not prev["sr"]:
        logging.info("Mode toggle (SR pressed)")
        toggle_mode()

    #  PLUS – reset gyroscope orientation
    if state["plus"] and not prev["plus"]:
        logging.info("Reset gyroscope orientation")
        joycon_gyro.reset_orientation()

    #  HOME – exit
    if state["home"] and not prev["home"]:
        logging.info("EXIT – cleaning up")
        highlighter_off()  # make sure no annotation is left on the slide
        sys.exit(0)

    # Update button history
    prev = {
        k: state[k] for k in prev
    }

    # Small delay – you can raise this if you find the loop too fast
    time.sleep(0.02)