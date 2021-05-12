# README  

## How to run
First, install python3 and necessary python package.
``` shell
pip install pandas
pip install pyqt5
```
Run `view.py` file
``` shell
py view.py
```

## How to use
After open the user interface, select source file and target file. Type in key column name, then hit the **`Start`** button. The program would generate `.auto.xls` or `.auto.xlsx` file in the same directory as target file.  
All the program runtime information will be presented in the text box at bottom.

## Other
The .ui file is exported from Qt Designer.  
Use following command to build .py file through .ui file.
``` shell
pyuic5 -x .\test.ui -o ui.py
```

Python freezing.
``` shell
py setup.py
```