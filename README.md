# README  

## How to Run in Python Enviroment
First, install python3 and necessary python package.
``` shell
pip install pandas
pip install pyqt5
```
Run `view.py` file
``` shell
py view.py
```

## How to Use
After open the user interface, select source file and target file. Type in key column header name, then hit the **`Start`** button. The program would generate `.auto.xls` or `.auto.xlsx` file in the same directory as target file.  
All the program runtime information will be presented in the text box at bottom.

## How to Build
Use `setup.py` file to help build.
``` shell
py setup.py
```
After processing, `view.exe` would be generated in `dist` folder.

## Other
**Qt Designer** is a helpful tool to build GUI visually. 
The `.ui` file can be exported from Qt Designer.  
Use following command to convert through `.ui` file into `view.js`
``` shell
pyuic5 -x <filename>.ui -o view.py
```