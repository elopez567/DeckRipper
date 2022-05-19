# DeckRipper

Automated web browser and Tkinter GUI used to download online presentations that are non-downloadable.

## Description

DeckRipper.py is a Tkinter GUI python script that is used to save down online presentations that do not give the user the option to download. The application first opens an automated web browser from where the user can enter their credentials. The user then can enter the number of slides and record the screen dimensions for the program to screenshot. The program clicks through the slides, capturing a screenshot of each slide into a single folder. Finally, all the screenshots are compiled into a single powerpoint file where the user can then save down on their computer.

## Getting Started

### Dependencies

- Libraries needed: pyscreenshot, time, selenium, tkinter, pillow, and pptx
- Files needed: chromedriver

### Installing

- pip install the libraries stated under dependencies
- download the appropriate version of chromedriver compatible with your version of google chrome (https://chromedriver.chromium.org/downloads)
- within the script change the driver path to where you saved chromedriver
- within the script change the output path to desired path to save screenshot images and powerpoint file
- move script into desired IDE

### Executing program

- Run script within desired IDE
- Chrome will automatically open to Deal Roadshow
- Enter the Roadshow credenntials and log in 
- After you log in, enter the number of slides shown into the Tkinter entry box labeled "# of Slides"
- Click "Select Area", this will open a screenshot of the current page you are on
- Drag a retangle around the presentation screensize
- Click "Record Dimensions", this will record the presentation screen dimensions and close the screenshot
- Next click "Rip Deck", this will cycle through the presentation and take screenshots of every slide.
- Finally, click "Create PP" to compile all the screenshots into a single powerpoint presentation.
- Close the Tkinter app and navigate to your output folder to find the saved powerpoint.

## Help

Common Issue: Be cautious to make sure the automated web browser is full screen when selecting the area and recording dimensions. This will help to get the most accurate dimensions when the program cycles through the powerpoint and screenshots each slide.

## Authors

Emmanuel Lopez - LinkedIn: (https://www.linkedin.com/in/lopez-emmanuel/)

## Version History

* 1.0
    * Initial Release

