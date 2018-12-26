# AutoSigns

AutoSigns is an application designed to automate the following:

1. Downloading the Section Schedule Daily Summary from Destiny and creating a sorted/formatted report.
2. Creating the signage for classes running. This includes:
	* Classroom signs
        * Daily Schedules
        * PowerPoint slide deck for TV display

## Getting Started
### Prerequisites

The following must be installed:

* Google Chrome
* Microsoft Excel
* Microsoft PowerPoint
* Microsoft Word

### Usage

1. Check the box next to the function(s) you would like to use
2. Fill in the require fields.
3. Click "Start" when ready. The output files will be saved to your "Save Path" location.
4. Click "Exit" to close the application.

Note: Runtime may vary depending on the number of days/classes that need the signs to be created for.


## Development
### Built With

* [Python 3.6](https://docs.python.org/3/) - The scripting language used.
* [Pandas](https://pandas.pydata.org/) - Data structure/anaylsis tool used.
* [python-docx](https://python-docx.readthedocs.io/en/latest/) - Used to create Microsoft Word documents (Classroom signs)
* [Selenium](https://selenium-python.readthedocs.io/) - Web crawling automation framework.
* [xlsxwriter](https://xlsxwriter.readthedocs.io/) - Used to create Microsoft Excel documents (Daily Schedule)
* [python-pptx](https://python-pptx.readthedocs.io/en/latest/) - Used to create Microsoft PowerPoint documents
* [PyQt5](https://pypi.org/project/PyQt5/) - Framework used to create GUI.
* [QtDesigner](http://doc.qt.io/qt-5/qtdesigner-manual.html) - GUI builder tool.
* [PyInstaller](https://www.pyinstaller.org/) - Used to create executable for release.

### Running the Script
Run the following command to installer all the required Python modules:
```
pip install -r requirements.txt
```
To run the application:
```
.\AutoSigns.py
```

## Authors
* **Patrick Yu** - *Initial work* - [patrickgod1](https://github.com/patrickgod1)
