# JoyCon Presentation Remote
Control your PowerPoint presentations with a Nintendo Switch Joy-Con.

Adapted from [Joycon_Presentation_Remote](https://github.com/leoshome/Joycon_Presentation_Remote) using [joycon-python Library](https://github.com/tocoteron/joycon-python).

## Features
- [x] Next/Previous Slides
- [x] Laser/Cursor Toggle
- [x] Gyro-based Pointer Control
- [ ] Left/Right Joy-Con Option
- [ ] Keybind Remapping
- [ ] PowerPoint Highlighter
- [ ] Joystick-based Pointer Control

## Usage Instructions
1. Download `joycon_rcon.py` script or clone this repo.
2. Install the required additional libraries through `pip`: `pip install joycon-python hidapi pyglm pyautogui`
3. Pair a Nintendo Switch Right Joy-Con (Joy-Con (R)) via Bluetooth to the computer. The script currently only supports Joy-Con (R)
4. Start a PowerPoint presentation
5. Run the script: `python joycon_rcon.py`
6. Return the curson and window to the PowerPoint presentation
7. Control the presentation through the following keybinds:
### Keybinds
- Press ➕ to reset Joy-Con gyro calibration - do this when having issues with motion control
- Press `SR` to toggle between cursor and laser pointer
- Hold `ZR` or `R` to move the cursor/laser pointer around
- Press `B` for next slide (Page Down)
- Press `X` for previous slide (Page Up)
- Press `A` for right mouse click
- Press `Y` for left mouse click
- Press `HOME`🏠 to end the slideshow and the script

## Future Ideas
- Joy-Con 2 Compatibility (incl. Mouse Mode)
