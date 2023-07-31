# AutoSigns

AutoSigns is an application designed to automate the following:

1. Extract/loading the Section Schedule Daily Summary report.
2. Transform the data by sorting/formating the report.
3. Creating the signage from the final report. This includes:
	* Classroom signs
    * Daily Schedules
    * PowerPoint slide deck for TV display
4. If enabled, write the daiy schedule report to Google Sheets for use with linked table in Google Slides and updated using Google App Script.

## Getting Started
These instructions will help you get started with using the application.

### Prerequisites
The following must be installed:
* Google Chrome
* Microsoft Excel
* Microsoft PowerPoint
* Microsoft Word
* Enable Google Drive API and Sheets API in Google Developer Console
    * Create a service account and save credential file as service_file.json in project root directory
    * Share/allow the service account access to Google Calendar and Google Drive
* config.ini file with Google Sheets URL to write to

### config.ini file
```
[General]
genReport=true
saveReportToPath=
createSigns=true
useExistingReport=false
classroomSignsOutput=true
dailyScheduleOutput=true
powerpointOutput=true
saveSignsDirectory=
uploadGBCSchedule=true
GBCScheduleURL=
uploadSFCSchedule=true
SFCScheduleURL=
```

### Usage
1. Check the box next to the function(s) you would like to use
2. Fill in the require fields.
3. Click "Start" when ready. The output files will be saved to your "Save Path" location.
4. Click "Exit" to close the application.

Note: Runtime may vary depending on the number of days/classes that need the signs to be created for.

## Development
These instructions will get you a copy of the project up and running on your local machine for development.

### Built With
* [Python 3.11](https://docs.python.org/3/) - The scripting language used.
* [Pandas](https://pandas.pydata.org/) - Data structure/anaylsis tool used.
* [python-docx](https://python-docx.readthedocs.io/en/latest/) - Used to create Microsoft Word documents (Classroom signs)
* [Selenium](https://selenium-python.readthedocs.io/) - Web crawling automation framework and Chrome webdriver manager.
* [Chrome Webdriver](http://chromedriver.chromium.org/downloads) - Webdriver for Chrome browser. Use to control automation with Selenium.
* [xlsxwriter](https://xlsxwriter.readthedocs.io/) - Used to create Microsoft Excel documents (Daily Schedule)
* [python-pptx](https://python-pptx.readthedocs.io/en/latest/) - Used to create Microsoft PowerPoint documents
* [PyQt5](https://pypi.org/project/PyQt5/) - Framework used to create GUI.
* [QtDesigner](http://doc.qt.io/qt-5/qtdesigner-manual.html) - GUI builder tool.
* [PyInstaller](https://www.pyinstaller.org/) - Used to create executable for release.
* [pygsheets](https://pygsheets.readthedocs.io/en/latest/) - Google Sheets API wrapper used.

### Running the Script
Run the following command to installer all the required Python modules:
```
pip install -r requirements.txt
```
To run the application:
```
python .\AutoSigns.py
```

### Compiling using PyInstaller

The project files includes a batch file (Windows platform only) with commands to run to compile into an executable. 

Other development platforms can run the following command in Terminal:

```
pyinstaller AutoSigns.spec .\AutoSigns.py
```
You may need to modify the file paths if not in same current working directory.

## Screenshot
![autosigns](https://user-images.githubusercontent.com/41496510/50427544-e0ec2b00-085f-11e9-9afe-e2a2f0a2d55b.png)

## Authors
* **Patrick Yu** - *Initial work* - [patrickgod1](https://github.com/patrickgod1) - UC Berkeley Extension
