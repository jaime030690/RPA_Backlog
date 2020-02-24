# Import dependencies
import pyautogui
import os
from time import sleep
from params import queue_list

#  Setup operations for pyautogui
def setup_pyautogui():
    
    # Enable fail safe, move mouse to top left corner of screen
    pyautogui.FAILSAFE = True

    # Step by 1 sec
    pyautogui.PAUSE  = 1.0

# Locates element on screen returning location if found, raises exception if not found
def locate_element(fileName):
    onScreen = pyautogui.locateOnScreen(fileName)

    if onScreen == None:
        raise pyautogui.PyAutoGUIException

    return onScreen

# Function for gathering data from application and copying to clipboard
def scrape_data(data):

    # Click TSP button
    pyautogui.click(locate_element(os.path.join('Resources', 'tsp_button.png')))
    sleep(3)

    # Click Query button
    pyautogui.click(locate_element(os.path.join('Resources', 'query_button.png')))
    sleep(3)

    # Click queue list
    pyautogui.click(locate_element(os.path.join('Resources', 'queue_list.png')))
    sleep(3)

    # Paste data list
    for line in data:
        pyautogui.typewrite(line)
        pyautogui.press('enter')

    # Press F2 to search
    pyautogui.press('f2')
    sleep(10)
    
    # Click 'Case HTML' button
    pyautogui.click(locate_element(os.path.join('Resources', 'case_html_button.png')))
    sleep(10)
    
    # Popup window will open, click clear space, select all and copy
    pyautogui.click(locate_element(os.path.join('Resources', 'query_results.png')))
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.hotkey('ctrl', 'c')
    
    # Close windows previously opened
    for _ in range(3):
        pyautogui.hotkey('esc')
        sleep(3)

# Function for opening the report file and running the report
def open_file():

    # Search for file in start menu
    pyautogui.press('win')
    pyautogui.typewrite('Electronic Backlog Report')
    pyautogui.press('enter')
    sleep(5)

    # Maximize window
    pyautogui.hotkey('win', 'up')
    sleep(5)

    # Run report
    pyautogui.click(locate_element(os.path.join('Resources', 'run_report.png')))

# Main function
def main():

    # Setup pyautogui module
    setup_pyautogui()

    # User message
    user_message = 'Application will start after you click OK!\n\nClick Cancel to end application. While application runs you may abort by moving mouse to the top left corner of the screen.'
    
    # Popup message with intructions, retrieve user's button choice
    user_input = pyautogui.confirm(text = user_message, title = 'Message', buttons = ['OK', 'Cancel'])

    # Check user input
    if user_input == 'OK':

        # Function calls to start process, exception handling
        try:
            scrape_data(queue_list)
            open_file()

        except pyautogui.FailSafeException:
            pyautogui.alert(text = 'Application ended by user!', title = 'Message')
    
        except pyautogui.PyAutoGUIException:
            pyautogui.alert(text = 'Element not found on screen!', title = 'Message')
        
        except OSError as e:
            pyautogui.alert(text = 'Error!\n\n' + str(e), title = 'Message')

    # Message to signal end
    pyautogui.alert(text = 'Application finished!', title = 'Message')

# Run main
if __name__ == "__main__":
    main()