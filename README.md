# Desktop App: Breakdown Dashboard

## [Video Demo](https://youtu.be/suFbsJLzZP4)

### Description:
This application is a desktop application (standalone exe) used to display data from machine breakdown reports extracted from excel files. The extracted data is in the form of breakdown information, countermeasures, owners and when the task must be completed. This application can monitor whether the task has been completed or not.

#### How the Program Works?

Tkinter serves as Python’s primary binding to the Tk GUI toolkit, providing a standardized interface for graphical user interfaces (GUIs) across Linux, Microsoft Windows, and macOS platforms. However, its GUI elements often appear outdated compared to other frameworks. By incorporating ttkbootstrap, developers can modernize their Tkinter applications with sleek themes inspired by Bootstrap.

Xlwings is a Python library that makes it easy to call Python from Excel and vice versa. It creates reading and writing to and from Excel using Python easily. It can also be modified to act as a Python Server for Excel to synchronously exchange data between Python and Excel. Xlwings makes automating Excel with Python easy and can be used for- generating an automatic report, creating Excel embedded functions, manipulating Excel or CSV databases etc.

Pandas is a Python package that provides fast, flexible, and expressive data structures designed to make working with "relational" or "labeled" data both easy and intuitive. It aims to be the fundamental high-level building block for doing practical, real world data analysis in Python. Additionally, it has the broader goal of becoming the most powerful and flexible open source data analysis / manipulation tool available in any language. It is already well on its way towards this goal.

Matplotlib is a powerful plotting library in Python used for creating static, animated, and interactive visualizations.

SQLite is a C library that provides a lightweight disk-based database that doesn’t require a separate server process and allows accessing the database using a nonstandard variant of the SQL query language.

Another easy and unique method to convert python to exe file is a Python utility called Auto-py-to-exe which can easily transform a Python.py file into an executable that has all of its dependencies packed. Another advantage is that Auto-py-to-exe creates an executable file that is a built version of the source code rather than the original source code. This makes it harder for others to steal your code.

### TODO:
#### Download
Download the Repository through Clone Repository or Download Zip
```
git clone https://github.com/riyanardyanto/cs50-final-project.git
```
#### Installation
After download, go to `cmd` and navigate to the project folder directory.
```
cd cs50-final-project
``` 
Next, create virtual environment.
```
python -m venv .venv
``` 
Great. Now you have a virtual environment. The next step is to activate it.
```
.venv\Scripts\activate
``` 
Use [pip](https://pip.pypa.io/en/stable/) to install needed libraries.
```
$ pip install -r requirements.txt
```
#### Usage
Run the program python script `project.py` with [python](https://www.python.org/).
```
python project.py
```
Test the program python script `test_project.py` with [pytest](https://docs.pytest.org/en/7.2.x/).
```
pytest test_project.py
```
Run auto-py-to-exe to create standalone exe. 
```
auto-py-to-exe
```
After successfully running the program, it will generate exe file in output folder. 

![exe file output](https://raw.githubusercontent.com/riyanardyanto/cs50-final-project/main/output_file.png)

The Program output will be like this:

![App Image](https://raw.githubusercontent.com/riyanardyanto/cs50-final-project/main/app%20image.png)

