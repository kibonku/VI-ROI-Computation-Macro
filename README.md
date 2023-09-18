# NDVI RoI Macro

### 1. Intro
* The macro was developed utilizing PyQt5 1, a Python-based binding for the Qt cross-platform application toolkit, which enables batch image processing and data transformation capabilities. 
* This novel macro designed for automated batch processing, ultimately streamlining the series of tasks involved and offering a more effective method for managing and analyzing grass bed VI data.

### 2. Requirements
* python >= 3.5
* PyQT5 >= 5
* opencv-contrib-python >= 4
* XlsxWriter: latest one would be preferred
* pandas: latest one would be preferred

For example:

```
Python == 3.8.12
PyQt5 == 5.15.4
opencv-contrib-python == 4.5.2.54
XlsxWriter == 1.2.6
pandas == 1.2.4
Your OS: Windows
Package: Pip
Language: Python
CUDA: None
```

### 3. Usage as a cli
```bash
JSON_to_RoI -j "/path/to/jsonfile.json" -i "path/to/inputfolder" -o "path/to/outputfolder"
```

### References
- https://doc.qt.io/qt-5/qtexamplesandtutorials.html
- https://docs.opencv.org/4.5.2/
- https://docs.python.org/ko/3.7/library/argparse.html
- https://xlsxwriter.readthedocs.io/
