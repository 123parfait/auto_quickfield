# QuickField Automation

QuickField version: 6.2
Python version: 3.14.2 (32-bit)

## GUI
```
*test version of python：*
python -c "import struct;print(8*struct.calcsize('P'))"

*select path*
cd /d C:\Users\name\Desktop\QFAuto

*launch the code*
python -m pip install -r requirements.txt
python src\main.py gui

run_py32.bat src/main.py gui

*package：*
run_py32.bat -m pip install pyinstaller
run_py32.bat -m PyInstaller --noconfirm --clean --onedir --windowed --name QFAuto_V5 --paths src --hidden-import pythoncom --hidden-import pywintypes --hidden-import win32timezone gui_entry.py


```

## Usage
1. Open QuickField, prepare your model, and assign materials.
   Note: give separate labels to moving parts and parts whose physical values will change, so they can be controlled as variables later. The model's current position is treated as the start position, and placing the mover at the center is recommended.
2. Click `Load labels` to scan and load all labels from the current model.
3. In `Input`, choose the physical values to vary.
   Currently, only coil current in ampere-turns is supported.
   
   `all` means Cartesian product. For example, with 2 quantities and 2 values for each, you get 4 simulations.
   `pair` means one-to-one pairing. For example, with 2 quantities and 2 values for each, you get 2 simulations.
   
   Example:
   `bobine1=100,200` and `bobine2=300,400`
   `pair` => `(100,300)`, `(200,400)` (2 cases)
   `all` => `(100,300)`, `(100,400)`, `(200,300)`, `(200,400)` (4 cases)

   For a single label, these two input styles are equivalent:
   - one value each time: `bobine=100`, `bobine=200`, `bobine=300` (click `Confirm` three times)
   - multiple values once: `bobine=100,200,300` (click `Confirm` once)
4. Select the labels to move and the output quantities to export.
5. In `Motion`, set the start and end positions of the moving part.
   Only linear motion is currently supported.
6. Select where to save the output table, then click `Run` to start batch simulation.

Goal: build a simple, convenient plugin for batch topology optimization of linear motors.

## Next steps
- [ ] Polish the UI.
- [ ] Optimize the process to prevent QuickField from freezing during batch simulations.
- [ ] Add more features and extend support to rotating motors.
