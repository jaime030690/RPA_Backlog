# RPA Backlog
The purpose of this project is to automate the creation of the Electronic Backlog Report. This report uses an Excel Macros to summary and visualize the data from the ticketing application. Getting the raw data from the ticketing application is a dull, repetitive process which is prone to user error. Using RPA (Robotic Process Automation) principles we identified this process as a perfect candidate for automation.
## Getting Started
The project is written in Python and relies on the [Pyautogui](https://pyautogui.readthedocs.io/en/latest/) module to automate the tasks. Pyautogui allows the programmer to interact with elements on the screen, input text, move the mouse, perform clicks, as well as pattern matching of screen contents. To install Pyautogui on your Python environment use the command below.
```
    pip install pyautogui
```
## Resources
Due to security concerns we had to ignore the `Resources` directory and `params.py` file. The `Resources` directory contain screenshots of the client application which is confidential; the `params.py` file contains a list with the names of work queues.
## How It Works
Below is a rough breakdown of the tasks needed to generate the Electronic Backlog Report.
1. Run query in client appication with all Electronic queue names
2. Export query results as an HTML file
3. Copy HTML content
4. Generate report using Excel file and macros

All of these functions could be performed using hotkey presses, but the client application has few hotkey functions. The ones that are available are of no use to this project. Therefore to interact with the client application we use `pyautogui.locateOnScreen()`. The function takes a file name as its argument. This file contains a screenshot of a region in the application where we'd like to move the mouse to. When the function is called a screenshot of the primary monitor is taken and Pyautogui looks for a matching section.

Electronic cases (or tickets) go to work queues in the client application. Each queue has a unique identifier which is saved in the `params.py` file as a list. This list is used to pass on paremeters for the query in the client application. This query generates a report which can be exported to HTML.

The query in HTML is then copied to the clipboard. The application opens the Excel document containing the report and runs the macro.

## Limitations
* Multiple monitor setups are not supported by the function `pyautogui.locateOnScreen()`, only the primary monitor will be seen
* The primary monitor needs to be "clean", meaning that other open applications should be minimized when possible
* The zoom level in the client application needs to be set at 111%, this number was chosen as it is easier to read for the developer :)
