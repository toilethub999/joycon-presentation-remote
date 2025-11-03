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
import logging

# ------------------------------------------------------------
# Configuration – tweak these if you like
# ------------------------------------------------------------
MOVE_SPEED = 1500          # how fast the mouse pointer moves
MODE_LASER = 0
MODE_CURSOR = 1
# ------------------------------------------------------------

logging.basicConfig(
    level=logging.INFO,
    format="[%(asctime)s] %(levelname)s – %(message)s",
    datefmt="%H:%M:%S",
)

# ------------------------------------------------------------
# Initialize Joy-Cons
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
    mode = MODE_LASER
except Exception as e:
    logging.error("Could not initialise Joy‑Con: %s", e)
    sys.exit(1)

# ------------------------------------------------------------
# Helper functions to *toggle* the built‑in tools
# TODO: COM connection not working - try different library? Highlighter function without COM?
# ------------------------------------------------------------
def laser_on():
    """Turn the PowerPoint laser pointer on (the L key toggles it)."""
    logging.info("Laser pointer ON")
    pyautogui.press("l")  # keycode 'L' toggles the laser

def laser_off():
    """Turn the PowerPoint laser pointer off."""
    logging.info("Laser pointer OFF")
    pyautogui.press("l")  # press L again

def toggle_mode():
    """Called whenever SR is pressed – switches between laser & cursor."""
    global mode
    if mode == MODE_LASER:
        # switch to cursor
        laser_off()
        mode = MODE_CURSOR
    else:
        # switch to laser
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
    if state["buttons"]["right"]["r"] or state["buttons"]["right"]["zr"]:          # only move when pressed
        pyautogui.moveRel(dx * MOVE_SPEED, dy * MOVE_SPEED, duration=0)
    pre_pos_x, pre_pos_y = cur_x, cur_y

    # ------------------------------------------------------------
    # Button handling – detect *new* presses
    # ------------------------------------------------------------
    #  X  – page‑up
    if state["buttons"]["right"]["x"] and not prev["buttons"]["right"]["x"]:
        logging.info("Page‑up")
        pyautogui.press("pageup")
    #  B  – page‑down
    if state["buttons"]["right"]["b"] and not prev["buttons"]["right"]["b"]:
        logging.info("Page‑down")
        pyautogui.press("pagedown")
    #  Y  – left click
    if state["buttons"]["right"]["y"] and not prev["buttons"]["right"]["y"]:
        logging.info("Left click")
        pyautogui.click()
    #  A  – right click
    if state["buttons"]["right"]["a"] and not prev["buttons"]["right"]["a"]:
        logging.info("Right click")
        pyautogui.click(button="right")

    #  SR – toggle mode (laser ↔ highlighter)
    if state["buttons"]["right"]["sr"] and not prev["buttons"]["right"]["sr"]:
        logging.info("Mode toggle (SR pressed)")
        toggle_mode()

    #  PLUS – reset gyroscope orientation
    if state["buttons"]["shared"]["plus"] and not prev["buttons"]["shared"]["plus"]:
        logging.info("Reset gyroscope orientation")
        joycon_gyro.reset_orientation()

    #  HOME – exit
    if state["buttons"]["shared"]["home"] and not prev["buttons"]["shared"]["home"]:
        logging.info("EXIT – cleaning up")
        sys.exit(0)

    # Update button history
    prev = state

    # Small delay – you can raise this if you find the loop too fast
    time.sleep(0.02)