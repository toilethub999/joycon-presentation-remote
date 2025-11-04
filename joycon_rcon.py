#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
PowerPoint Remote – Joy‑Con controlled laser pointer
Author: T
"""
# ------------------------------------------------------------
# Import and Dependencies 
# ------------------------------------------------------------
ERR_IMPORT = False
import sys
import time
import logging
logging.basicConfig(
    level=logging.INFO,
    format="[%(asctime)s] %(levelname)s – %(message)s",
    datefmt="%H:%M:%S",
)
try:
    from pyjoycon import JoyCon, get_R_id, GyroTrackingJoyCon
except ImportError:
    logging.error("❌  Module pyjoycon not found – run `pip install joycon-python hidapi pyglm`")
    ERR_IMPORT = True
try:
    import pyautogui
except ImportError:
    logging.error("❌  Module pyautogui not found – run `pip install pyautogui`")
    ERR_IMPORT = True
if ERR_IMPORT:
    logging.error("Missing Dependencies - please install the required modules and try again")
    sys.exit()
else:
    logging.info("Modules successfully imported.")

# ------------------------------------------------------------
# Configuration – tweak these if you like
# ------------------------------------------------------------
MOVE_SPEED = 1500          # how fast the mouse pointer moves
pyautogui.FAILSAFE = False

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
except Exception as e:
    logging.error("Could not initialise Joy‑Con: %s", e)
    sys.exit(1)

# ------------------------------------------------------------
# Helper functions to *toggle* the built‑in tools
# TODO: COM connection not working - try different library? Highlighter function without COM?
# ------------------------------------------------------------


# ------------------------------------------------------------
# Main loop – read the joystick, move the cursor, toggle tools
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
    # Gyroscope → mouse movement
    # ------------------------------------------------------------
    gyro = joycon_gyro.pointer
    try:
        cur_x, cur_y = gyro[0], -gyro[1]  # note the sign flip
    except TypeError:
        # Catch gyro pointer returning None error, temporarily using previous values
        logging.warning("Lost Gyro, using previous values. Try recalibrating Joy-Con.")
        cur_x, cur_y = pre_pos_x, pre_pos_y

    dx, dy = cur_x - pre_pos_x, cur_y - pre_pos_y
    if state["buttons"]["right"]["r"] or state["buttons"]["right"]["zr"]:          # only move when pressed
        pyautogui.moveRel(dx * MOVE_SPEED, dy * MOVE_SPEED, duration=0, _pause=False)
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
    #  SR – toggle mode (laser <-> cursor)
    if state["buttons"]["right"]["sr"] and not prev["buttons"]["right"]["sr"]:
        logging.info("Toggle Laser Pointer")
        pyautogui.hotkey("ctrl","l")  # hotkey 'CTRL+L' toggles the laser
    #  PLUS – reset gyroscope orientation
    if state["buttons"]["shared"]["plus"] and not prev["buttons"]["shared"]["plus"]:
        logging.info("Reset gyroscope orientation")
        joycon_gyro.reset_orientation()
    #  HOME – exit
    if state["buttons"]["shared"]["home"] and not prev["buttons"]["shared"]["home"]:
        logging.info("EXIT – cleaning up")
        pyautogui.press("esc")
        sys.exit(0)

    # Update button history
    prev = state