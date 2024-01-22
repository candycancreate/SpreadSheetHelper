# openpyxl-helper

openpyxl-helper is a simple extension of [openpyxl](https://openpyxl.readthedocs.io/en/stable/) which is a Python library to read/write Excel xlsx/xlsm/xltx/xltm files.

This project collects some usually used operations for Excel files.

## Install and setup poetry

### Install poetry

Install By officall installer

##### macOS / Linux / WSL（Windows Subsystem for Linux）
```
curl -sSL https://install.python-poetry.org | python3 -
```
##### Windows
```
(Invoke-WebRequest -Uri https://install.python-poetry.org -UseBasicParsing).Content | python -
```

### Setup poetry

#### set path

##### macOS / Linux / WSL（Windows Subsystem for Linux）
```
export PATH=$PATH:$HOME/.local/bin
```

##### Windows
```
$Env:Path += ";C:\Users\<user name>\AppData\Roaming\Python\Scripts"; setx PATH "$Env:Path"
```

#### Set config of env in project folder
```
poetry config virtualenvs.in-project true
```

## Create environment

```
cd <project_path>
poetry lock
poetry install --no-root
poetry env use <python version>
```
For example
```
cd ~/SpreadSheetHelper/src/openpyxl-helper
poetry lock
poetry install --no-root
poetry env use 3.8
```

## Active enviroment

```
poetry shell
```

### exit

```
exit
```

## Execute

### Run python
run python in the virtual enviroment

```
poetry run python
```
then you can use the package, for example
```
from openpyxl_helper.openpyxlhelper import OpenpyxlHelper
OpenpyxlHelper.get_sheet_names_by_file_path("../../test/file/ReadTest.xlsx")
```

### Run test

```
poetry run python -m unittest
```
verbosity mode
```
poetry run python -m unittest -v
```
