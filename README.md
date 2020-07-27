
# CPR Tools File Checker

This is a program that I created for my company in order to check for corrupt media files in a directory. Supported file types are:

 - **Image:** JPEG, JPG, PNG, GIF, BMP
 - **Document:** PDF, DOCX, XLSX, PPTX
 - **Audio:** MP3, OGG, FLAC
 - **Video:** MP4, AVI, MOV, WMV, MTS, MPG

## Installation
A minimum of Python 3.7 is recommended for maximum compatibility.

Use the package manager  [pip](https://pip.pypa.io/en/stable/)  to install this program's dependencies:

```bash
pip install Pandas
pip install PySide2
pip install Mutagen
pip install Moviepy
pip install PyPDF2
pip install Docx2txt
pip install Func_Timeout
pip install XLRD
pip install soundfile
pip install Python-PPTX
```

## Usage

**In File Path:**

This is the directory you wish to check for corrupt files anything within this directory or any of its sub-folders will be evaluated for corruption.

**Out File Path**

This is the directory in which you wish to save your results file and store your copied/moved files if applicable.

**Copy/Move/NA Radio Options:**

These are used to determine whether you wish to copy or move your files to the out file path. 

**Generate Extension Report**

Checking this option will tell the program to create an extension report which is a breakdown of of the good and bad files by extension. 

