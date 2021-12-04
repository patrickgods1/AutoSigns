======== AutoSigns ========

AutoSigns is an application designed to automate the following:

1. Extract/loading the Section Schedule Daily Summary report.
2. Transform the data by sorting/formating the report.
3. Creating the signage from the final report. This includes:
    a. Classroom signs
    b. Daily Schedules
    c. PowerPoint slide deck for TV display

-------- Prerequisites --------

The following must be installed:

1. Google Chrome
2. Microsoft Excel
3. Microsoft PowerPoint
4. Microsoft Word

-------- Usage --------

1. Check the box next to the function(s) you would like to use
2. Fill in the require fields.
3. Click "Start" when ready. The output files will be saved to your "Save Path" location.
4. Click "Exit" to close the application.

Note: Runtime may vary depending on the number of days/classes that need the signs to be created for.


======== Development ========

-------- Built With --------

1. Python 3.X standard library - Used for basic logic flow.
2. Python Pandas library - Data framework for reading and manipulating data.
3. python-docx library - Used to create Microsoft Word documents (Classroom signs)
4. Python Selenium library - Web crawling automation framework.
5. Webdriver-Manager - Framework to manage webdriver version compatible with version of browser
6. xlsxwriter - Used to create Microsoft Excel documents (Daily Schedule)
7. python-pptx library - Used to create Microsoft PowerPoint documents
8. PyQt5 - GUI framework.
9. QtDesigner - GUI builder tool.
10. PyInstaller - Bundles Python applications and all its dependencies into an executable.

-------- Authors --------

Patrick Yu - Initial work - UC Berkeley Extension
Unknown - Creators of the original VBA macros/scripts
