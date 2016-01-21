xy2kml v1.0
===========

Converts XLSX to KML and transforms point coordinates from one coordinate
system to another if needed. Windows executable can be downloaded from
http://yilmazturk.info/xy2kml/

CONTENTS OF THIS FILE
---------------------
* Introduction
* Requirements
* Usage
* FAQ

INTRODUCTION
------------
xy2kml converts x, y point coordinate values which are included in XLSX
document to KML document to view in Google Earth. 
If coordinate transformation option is chosen, the output (transformed values) 
will be written to another XLSX file named 'transformed.xlsx'.

Excel sample data named 'coord_sample.xlsx' that you can use for testing or 
comparison for results is included in distribution. 
This data consists of two sheets; 'UTM35N_WGS84' and 'TM30_ED50'. 
EPSG values for easting & northing are 32635 and 2320, respectively.
In addition, EPSG value for lon & lat (GCS_WGS84) is 4326.

Coordinate transformation will be performed based on PROJ.4 library
with SpatiaLite approach. You need to define your input & output coordinates'
EPSG value to accomplish transformation successfully. 

If any further information about EPSG is needed, please visit the following
websites:

* http://www.epsg.org/
* http://spatialreference.org/ref/epsg/
* http://epsg.io/

REQUIREMENTS
------------
This software requires the following modules:
* openpyxl (https://pypi.python.org/pypi/openpyxl)
* pysqlite2 (https://pypi.python.org/pypi/pysqlite)

USAGE
-----
* File > Open (reads your XLSX file, first row should contain column names)
* Select sheet name that includes coordinate values
* Select column names for x, y values
* Press 'Set CRS' button to specify EPSG value
* If you press 'Generate KML' button, will generate KML 
* If you choose 'Transform & Write to Excel?' option and press 
'Generate KML' button, will generate KML with an XLSX document in the same path
which includes transformed values (please note that when you choose 
'Transform & Write to Excel?' option, output CRS combobox will be active and
needed to specify new CRS for transformation

FAQ
---
Q: I press 'Generate KML' button and nothing happened, what's the matter?<br />
A: You're sure to specify sheet name, x & y column names and EPSG value.

Q: KML has been generated but shows wrong location...<br />
A: You need to define correct EPSG value to perform operation properly.


