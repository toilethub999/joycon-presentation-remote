#!/usr/bin/env python3
# coding: utf-8

"""
PowerPoint Remote - Joy-Con controlled laser pointer
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
    format="[%(asctime)s] %(levelname)s - %(message)s",
    datefmt="%H:%M:%S",
)
try:
    from pyjoycon import JoyCon, get_R_id, GyroTrackingJoyCon
except ImportError:
    logging.error("❌  Module pyjoycon not found - run `pip install joycon-python hidapi pyglm`")
    ERR_IMPORT = True
try:
    import pyautogui
except ImportError:
    logging.error("❌  Module pyautogui not found - run `pip install pyautogui`")
    ERR_IMPORT = True
if ERR_IMPORT:
    logging.error("Missing Dependencies - please install the required modules and try again")
    sys.exit()
else:
    logging.info("Modules successfully imported.")

# ------------------------------------------------------------
# Configuration - tweak these if you like
# ------------------------------------------------------------
MOVE_SPEED = 1500           # mouse pointer movement speed
keybind = {                 # modify keybinds here
    "next"          :"a",
    "prev"          :"y",
    "toggle_mode"   :"b",
    "move_pointer"  :"x",
    "lmb"           :"r",
    "rmb"           :"zr",
}
pyautogui.FAILSAFE = False  # pyautogui's safeguard against edges - enabling it may cause errors on edges

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
    logging.info("Found Joy-Con (R). Battery status: %d", state["battery"]["level"])
    if state["battery"]["level"] == 1:
        logging.warning("Low battery detected. Consider recharging Joy-Con before use.")
except Exception as e:
    logging.error("Could not initialise Joy-Con: %s", e)
    sys.exit(1)

# ------------------------------------------------------------
# Helper functions to *toggle* the built-in tools
# TODO: COM connection not working - try different library? Highlighter function without COM?
# ------------------------------------------------------------


# ------------------------------------------------------------
# Main loop - read the joystick, move the cursor, toggle tools
# ------------------------------------------------------------
logging.info("Remote is now running.  Press HOME on the Joy-Con to exit.")
while True:
    try:
        state = joycon.get_status()
        gyro = joycon_gyro.pointer
    except Exception as e:
        logging.warning("Lost Joy-Con connection: %s", e)
        time.sleep(0.5)
        continue

    # ------------------------------------------------------------
    # Button handling - detect *new* presses
    # ------------------------------------------------------------
    # Gyroscope -> mouse movement
    if state["buttons"]["right"][keybind["move_pointer"]]:
        # only move when pressed
        if not prev["buttons"]["right"][keybind["move_pointer"]]:
            # reset gyro orientation on every motion start - more reliable tracking
            joycon_gyro.reset_orientation()
        try:
            cur_x, cur_y = gyro[0], -gyro[1]  # note the sign flip
            dx, dy = cur_x - pre_pos_x, cur_y - pre_pos_y
            pyautogui.moveRel(dx * MOVE_SPEED, dy * MOVE_SPEED, duration=0, _pause=False)
            pre_pos_x, pre_pos_y = cur_x, cur_y
        except TypeError:
            # Catch gyro pointer returning None error, temporarily using previous values
            logging.warning("Lost gyroscope data - using previous values. Try recalibrating Joy-Con by pressing '+'.")
            cur_x, cur_y = pre_pos_x, pre_pos_y     
    #  Next  - page down
    if state["buttons"]["right"][keybind["next"]] and not prev["buttons"]["right"][keybind["next"]]:
        logging.info("Next")
        pyautogui.press("pagedown")
    #  Previous  - page up
    if state["buttons"]["right"][keybind["prev"]] and not prev["buttons"]["right"][keybind["prev"]]:
        logging.info("Previous")
        pyautogui.press("pageup")
    #  Left mouse button
    # if state["buttons"]["right"][keybind["lmb"]] and not prev["buttons"]["right"][keybind["lmb"]]:
    #     logging.info("Left click")
    #     pyautogui.click()
    #  Right mouse button
    # if state["buttons"]["right"][keybind["rmb"]] and not prev["buttons"]["right"][keybind["rmb"]]:
    #     logging.info("Right click")
    #     pyautogui.click(button="right")
    #  Toggle mode - (laser <-> cursor, hotkey CTRL+L)
    if state["buttons"]["right"][keybind["toggle_mode"]] and not prev["buttons"]["right"][keybind["toggle_mode"]]:
        logging.info("Toggle Laser Pointer")
        pyautogui.hotkey("ctrl","l")
    #  PLUS - reset gyroscope orientation
    if state["buttons"]["shared"]["plus"] and not prev["buttons"]["shared"]["plus"]:
        logging.info("Reset gyroscope orientation")
        joycon_gyro.reset_orientation()
        joycon_gyro.calibrate()
    #  HOME - exit
    if state["buttons"]["shared"]["home"] and not prev["buttons"]["shared"]["home"]:
        logging.info("EXIT - cleaning up")
        pyautogui.press("esc")
        sys.exit(0)

    # Update button history
    prev = state