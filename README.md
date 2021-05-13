# README  

## How to Run in Python Enviroment
First, install python3 and necessary python package. (`xlwt` is used to output .xls file, now is deprecated)
``` shell
pip install pandas
pip install pyqt5
pip install xlrd
pip install xlwt
pip install openpyxl
```
Run `view.py` file.
``` shell
py src/view.py
```

## How to Use
After open the user interface, select source file and target file. Type in key column header name, then hit the **`Start`** button. The program would generate `.auto.xls` or `.auto.xlsx` file in the same directory as target file.  
All the program runtime information will be presented in the text box at bottom.

## How to Build Distribution
First, enter python virtual environment.
``` shell
# This will generate a new folder called <env_name>
python -m venv <env_name>   

# Activate venv
.\<env_name>\Scripts\Activate.ps1

# Quit venv
.\<env_name>\Scripts\deactivate.bat
```

Run `setup.py` in venv.
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