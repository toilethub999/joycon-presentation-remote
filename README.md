# JoyCon Presentation Remote
Control PowerPoint presentations with a Nintendo Switch Joy-Con as a presenter remote control.

Adapted from [Joycon_Presentation_Remote](https://github.com/leoshome/Joycon_Presentation_Remote) using [joycon-python](https://github.com/tocoteron/joycon-python) and [pyautogui](https://github.com/asweigart/pyautogui).

## Features
- [x] Next/Previous Slides
- [x] Laser/Cursor Toggle
- [x] Gyro-based Pointer Control
- [ ] Left/Right Joy-Con Option
- [x] Keybind Remapping
- [ ] PowerPoint Highlighter
- [ ] Joystick-based Pointer Control

## Usage Instructions
1. Download `joycon_rcon.py` script or clone this repo.
2. Install the required additional libraries through `pip`
```pip install joycon-python hidapi pyglm pyautogui```
3. Pair a Nintendo Switch Right Joy-Con (Joy-Con (R)) via Bluetooth to the computer. The script currently only supports Joy-Con (R)
4. Start a PowerPoint presentation
5. Run the script: `python joycon_rcon.py`
6. Return the curson and window to the PowerPoint presentation
7. Control the presentation through the following keybinds:
### Default Keybinds
- Press ➕ to reset Joy-Con gyro calibration - do this when having issues with motion control
- Press `B` to toggle between cursor and laser pointer
- Hold `X` to move the cursor/laser pointer around
- Press `A` for next slide (Page Down)
- Press `Y` for previous slide (Page Up)
- Press `R` for right mouse click (disabled by default, enabled by uncommenting lines 116-118)
- Press `ZR` for left mouse click (disabled by default, enabled by uncommenting lines 120-122)
- Press `HOME`🏠 to end the slideshow and the script

To change the cursor movement speed or modify the keybinds, edit the configuration section on the `joycon_rcon.py` script (line 37) with any text/code editor. Save the file and rerun the script (Step 5) with your own configuration.

## Future Ideas
- Joy-Con 2 Compatibility (incl. Mouse Mode)
