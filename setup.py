import PyInstaller.__main__

# Build Python program to .exe file
PyInstaller.__main__.run([
    'view.py',
    '--onefile',
    '--windowed',
    '-p C:\Windows\System32\downlevel'
    '--name ExcelAutoProcess'
])