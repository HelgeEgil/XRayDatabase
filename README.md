# XRayDatabase
A helpful Python-based database and QA report generator for X-ray systems.

# Norwegian:
This is a software tools consisting of several parts. It contains:
* An sqlite3 database to keep track of X-ray and CT machines, with information about serial numbers, age, QA events, etc.
* It is GUI-based to perform all required tasks
* It can be populated automatically from a folder of DICOM images
* It can load and save information via CSV files
* With the supplied Excel worksheets, any QA measurements stored properly can be easily transferred into the SQL database
*	An automatic PDF report generator to create reports from the SQL database, from a single QA event or all events adhering to a given filter.

The software is written in Norwegian, with documentation.
