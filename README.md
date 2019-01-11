# Media Organizer
Find video media files in a messy directory.  
![Excel](http://i67.tinypic.com/b5sk8k.png)

## Installation/Usage:
Reqs:  
- Python 3.6
- Excel 2013 (any version should do)  

Install:  
Option 1 (.exe):  
Download the zip file of the repo, unzip, navigate to to '\movie_organizer-master\dist\movie_scanner' and find the .exe file, 'movie_scanner.exe' and run it

Option 2 (from the command line):  
Download the zip file of the repo, open a terminal in that directory.
```
# Install the required packages   
$ pip install -r requirements.txt

# Run the python file
$ python movie_scanner.py
```

The program will then prompt you for a Windows absolute path, i.e. 'C:\user\joe\downloads'

## TODO:
* The title recognition in the info_extract function still needs work in order to more effectively deal with oddly named video files.
