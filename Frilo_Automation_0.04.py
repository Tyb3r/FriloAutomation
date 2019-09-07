import pyautogui
import sys
import pandas as pd
import numpy as np
import time
import win32gui
import win32con
import getpass
import datetime

"""It's basicly just a simple autoclicker for Frilo Module B5. It's intended to work in every language. Various improvements could be implemented, both in terms of speed and reliability. For the sake of time and simplicity here they were omitted"""


def check_access():
    """A function that checks, if the user is authorized to use the software. It's super easy really, not intended to really stop someone who tries. If a person is not authorized, a trial period is started"""
    username = getpass.getuser()
    list_of_users = ['Kuba', 'jturbaki']  # a list of authorized users
    if username not in list_of_users:
        today = datetime.date.today()
        licence_end = datetime.date(2018, 12, 31)  # licence end time
        if licence_end - today < datetime.timedelta(seconds=0):
            pyautogui.alert(text='Trial for software has expired. Contact turbakiewicz.jakub@gmail.com', title='Not autorized', button='Close')
            sys.exit(0)


def maximize_window():
    """A funtion that makes the user maximize the Frilo window"""
    while True:
        """a mechanism implemented in various spots in the software. Basicly what it does it tries to find a specific .png file for x seconds. If it does, the loops end. If it doesn't, a user is asked to locate the image or the software quits"""
        t_end = time.time() + 20
        while time.time() < t_end:
            try:
                is_logo_visable = pyautogui.locateCenterOnScreen('src/logo.png')  # None = not visable
                if is_logo_visable is None:
                    raise  # if the logo is not found it asks to open and show Frilo soft
                pyautogui.click(is_logo_visable[0] + 20, is_logo_visable[1])  # if visable, clicks next to the logo to make sure its the active window
                hwnd = win32gui.GetForegroundWindow()  # gets the active window
                win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)  # maximizes the active window
                break  # if succeeded to this part, stop the loop
            except:
                button = pyautogui.confirm(text='Please open Frilo B5 and bring it to foreground', title='Frilo is not opened', buttons=['Done', 'Exit software'])
                if button == 'Exit software':
                    sys.exit(0)  # Shuts the software
        if is_logo_visable is None:
            not_found_button = pyautogui.confirm(text='logo.png not found. Please make it visible and let the software click it or close the software', title='Button not found', buttons=['Done', 'Exit software'])
            if not_found_button == 'Exit software':
                sys.exit(0)
        else:
            break
    button = pyautogui.confirm(text='Frilo auto-calculation will begin. Please save and backup all of relevant data.', title='Warining', buttons=['Done', 'Exit software'])
    if button == 'Exit software':
        sys.exit(0)
    else:
        return (pyautogui.size())


def clean_sidebar():
    """A function that maximalizes the left sidebar of Frilo"""
    if len(blue_squares) == 1:
        pyautogui.click(blue_squares[0][0], blue_squares[0][1], duration=time_to_move_mouse)
        pyautogui.moveRel(100)
        find_blue_squares()
    try:
        proj_location = pyautogui.locateCenterOnScreen('src/Proje.png', region=(0, 0, 100, height - 1))
        pyautogui.moveTo((proj_location[0], proj_location[1] + 11), duration=time_to_move_mouse)
        pyautogui.moveTo((proj_location[0] + 2, proj_location[1] + 11), duration=time_to_move_mouse)
        pyautogui.dragTo(proj_location[0] + 2, height, duration=0.2)  # drags the sidebar way down
    except:
        pass  # resolution that doesn't show the icons. Might cause problems in the later parts, which are caught.
    pluses = pyautogui.locateAllOnScreen('src/+.png')
    how_many_pluses = 0
    for plus in pluses:
        how_many_pluses += 1  # a remnant of pyautogui.locateAllOnScreen being a generator. Would have been better just to convert to a list. Next time maybe
    for plus in range(how_many_pluses):
        plus_location = pyautogui.locateCenterOnScreen('src/+.png', region=(0, 0, 50, height - 1))
        pyautogui.click(plus_location, duration=time_to_move_mouse)  # clicks all of the 'pluses' to have everything visable


def def_values(element):
    """A function that sets up a bulk of calculation parameters"""
    global is_rectangle_set  # variables that check which cross-section type is used
    global is_circle_set
    global is_ring_set
    concrete_class_name = str(excel_data['Concrete Class'][element])
    concrete_class_name = 'src/' + concrete_class_name.replace("/", "_") + ".png"  # converts proper concrete class name to something saveable: C12/15 -> src/C12_15.png
    if element == 0 or excel_data['Concrete Class'][element] != excel_data['Concrete Class'][element - 1]:  # seen multiple times in the script. If the x element is the same in this area as x-1 (but 0) the part gets ommited to save some time
        pyautogui.click(340, 105, duration=time_to_move_mouse)
        pyautogui.click(340, 118, duration=time_to_move_mouse)  # picks the lowest concrete class
        if concrete_class_name != "src/C12_15.png":
            pyautogui.click(340, 105, duration=time_to_move_mouse)
            while True:
                t_end = time.time() + time_to_wait
                while time.time() < t_end:
                    if concrete_class_name == "src/C12_15.png":
                        break  # if the lowest concrete was intended, its kept
                    concrete_class_location = pyautogui.locateCenterOnScreen(concrete_class_name)  # if not it tries to find the concrete class
                    if concrete_class_location is not None:
                        break
                if concrete_class_location is None:
                    not_found_button = pyautogui.confirm(text='Specific concrete class png file (Cxx_xx.png) not found. Please make it visible and let the software click it or close the software',
                                                         title='Button not found', buttons=['Done', 'Exit software'])
                    if not_found_button == 'Exit software':
                        sys.exit(0)
                else:
                    break
            pyautogui.click(concrete_class_location, duration=time_to_move_mouse)  # and if it has found it, it clicks it
    pyautogui.moveRel(xOffset=200, duration=time_to_move_mouse)
    find_and_click('src/Phi.png', offset_x=50, if_grayscale=True)
    pyautogui.click()
    #pyautogui.doubleClick(340, 145, duration=time_to_move_mouse)
    pyautogui.typewrite(str(excel_data['Fi'][element]), interval=time_between_keystrokes)
    pyautogui.press('tab')  # a lot of the soft is written like that. It imputs a value from excel and then it tabs a specific amount of times to a desired new spot.
    pyautogui.typewrite(str(excel_data['Length'][element]), interval=time_between_keystrokes)
    pyautogui.press('tab')

    if element == 0 or excel_data['Type of cross section'][element] != excel_data['Type of cross section'][element - 1]:
        is_rectangle_set = None
        is_circle_set = None
        is_ring_set = None

        while True:
            t_end = time.time() + time_to_wait
            while time.time() < t_end:
                is_rectangle_set = pyautogui.locateCenterOnScreen('src/rectangle.png', region=(0, 0, int(width / 2) - 1, int(height / 2) - 1))
                is_circle_set = pyautogui.locateCenterOnScreen('src/circle.png', region=(0, 0, int(width / 2) - 1, int(height / 2) - 1))
                is_ring_set = pyautogui.locateCenterOnScreen('src/ring.png', region=(0, 0, int(width / 2) - 1, int(height / 2) - 1))  # tries to find which cross-section is currently set up in Frilo

                if is_rectangle_set is not None or is_circle_set is not None or is_ring_set is not None:
                    break
            if is_rectangle_set is None and is_circle_set is None and is_ring_set is None:
                not_found_button = pyautogui.confirm(text='rectangle.png or circle.png or ring.png not found. Please make it visible and let the software click it or close the software', title='Button not found', buttons=['Done', 'Exit software'])
                if not_found_button == 'Exit software':
                    sys.exit(0)
            else:
                break

    if (str(excel_data['Type of cross section'][element]) == 'rectangle' and
            is_rectangle_set is not None):  # based on which cross-section type is currently set up in Frilo and which one the user has picked, it changed simply by using arrows required amount of times
        pyautogui.press('tab')
    elif (str(excel_data['Type of cross section'][element]) == 'rectangle' and
          is_circle_set is not None):
        pyautogui.press('down')
        pyautogui.press('down')
        pyautogui.press('tab')
    elif (str(excel_data['Type of cross section'][element]) == 'rectangle' and
          is_ring_set is not None):
        pyautogui.press('down')
        pyautogui.press('tab')
    elif (str(excel_data['Type of cross section'][element]) == 'circle' and
          is_circle_set is not None):
        pyautogui.press('tab')
    elif (str(excel_data['Type of cross section'][element]) == 'circle' and
          is_ring_set is not None):
        pyautogui.press('down')
        pyautogui.press('down')
        pyautogui.press('tab')
    elif (str(excel_data['Type of cross section'][element]) == 'circle' and
          is_rectangle_set is not None):
        pyautogui.press('down')
        pyautogui.press('tab')
    elif (str(excel_data['Type of cross section'][element]) == 'ring' and
          is_ring_set is not None):
        pyautogui.press('tab')
    elif (str(excel_data['Type of cross section'][element]) == 'ring' and
          is_circle_set is not None):
        pyautogui.press('down')
        pyautogui.press('tab')
    elif (str(excel_data['Type of cross section'][element]) == 'ring' and
          is_rectangle_set is not None):
        pyautogui.press('down')
        pyautogui.press('down')
        pyautogui.press('tab')
    time.sleep(1)  # all of the 'random' sleep commands help avoiding the fact, that it sometimes clicks/pressed too fast

    if str(excel_data['Type of cross section'][element]) == 'rectangle':
        if element == 0 or excel_data['Reinforcement distributed on perimeter'][element] != excel_data['Reinforcement distributed on perimeter'][element - 1]:
            should_it_tab = False
            # checks if the user wants to use reinforcement distributed on perimeter or not. If the setting differ from what is currently set in Frilo, its being changed.
            if str(excel_data['Reinforcement distributed on perimeter'][element]) == 'Yes':
                while True:
                    t_end = time.time() + time_to_wait
                    while time.time() < t_end:
                        perm_clicked_loc = pyautogui.locateCenterOnScreen('src/Perimeter_clicked.png')
                        perm_unclicked_loc = pyautogui.locateCenterOnScreen('src/Perimeter_unclicked.png')
                        if perm_clicked_loc is not None or perm_unclicked_loc is not None:
                            break
                    if perm_clicked_loc is None and perm_unclicked_loc is None:
                        not_found_button = pyautogui.confirm(text='Perimeter_clicked.png or Perimeter_unclicked.png not found. Please make it visible and let the software click it or close the software',
                                                             title='Button not found', buttons=['Done', 'Exit software'])
                        if not_found_button == 'Exit software':
                            sys.exit(0)
                    else:
                        break

                if perm_unclicked_loc is not None:
                    should_it_tab = True
                    pyautogui.click(perm_unclicked_loc, duration=time_to_move_mouse)
            elif str(excel_data['Reinforcement distributed on perimeter'][element]) == 'No':

                while True:
                    t_end = time.time() + time_to_wait
                    while time.time() < t_end:
                        corner_clicked_loc = pyautogui.locateCenterOnScreen('src/Corner_clicked.png')
                        corner_unclicked_loc = pyautogui.locateCenterOnScreen('src/Corner_unclicked.png')
                        if corner_clicked_loc is not None or corner_unclicked_loc is not None:
                            break
                    if corner_clicked_loc is None and corner_unclicked_loc is None:
                        not_found_button = pyautogui.confirm(text='Corner_clicked.png or Corner_unclicked.png not found. Please make it visible and let the software click it or close the software',
                                                             title='Button not found', buttons=['Done', 'Exit software'])
                        if not_found_button == 'Exit software':
                            sys.exit(0)
                    else:
                        break
                if corner_unclicked_loc is not None:
                    should_it_tab = True
                    pyautogui.click(corner_unclicked_loc, duration=time_to_move_mouse)
            # if perimeter reinforcement / corner reinforcement setting had to be changed, it gets done. After it the software has to 'go back' to the proper point simply by tabbing
            if should_it_tab:
                time.sleep(1)
                for tabs in range(7):
                    pyautogui.press('tab', interval=time_between_keystrokes)
    pyautogui.typewrite(str(excel_data['by'][element]), interval=time_between_keystrokes)
    pyautogui.press('tab')
    # depending on which cross-section is chosen, a different 'tabbing' path needs to be taken, as some of the windows are not active in specific configurations.
    if str(excel_data['Type of cross section'][element]) != 'circle':
        if str(excel_data['Type of cross section'][element]) == 'ring':
            if (float(excel_data['by'][element]) - float(excel_data['dz'][element])) / 2 < 8:
                pyautogui.typewrite(str(float(excel_data['by'][element]) - 16), interval=time_between_keystrokes)
                # for ring cross-section the 'flange' needs to be of min. specfic width
            else:
                pyautogui.typewrite(str(excel_data['dz'][element]), interval=time_between_keystrokes)
        else:
            pyautogui.typewrite(str(excel_data['dz'][element]), interval=time_between_keystrokes)
        pyautogui.press('tab')
    if str(excel_data['Type of cross section'][element]) == 'rectangle':
        pyautogui.typewrite(str(excel_data['b1'][element]), interval=time_between_keystrokes)
        pyautogui.press('tab')
    if str(excel_data['Type of cross section'][element]) == 'ring' and float(excel_data['d1'][element]) >= 4:
        pyautogui.typewrite('4', interval=time_between_keystrokes)
    else:
        pyautogui.typewrite(str(excel_data['d1'][element]), interval=time_between_keystrokes)
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.typewrite(str(excel_data['Vg'][element]), interval=time_between_keystrokes)
    pyautogui.press('tab')
    pyautogui.typewrite(str(excel_data['Vq'][element]), interval=time_between_keystrokes)
    pyautogui.press('tab')
    pyautogui.typewrite(str(excel_data['gy'][element]), interval=time_between_keystrokes)
    pyautogui.press('tab')
    pyautogui.typewrite(str(excel_data['qy'][element]), interval=time_between_keystrokes)
    pyautogui.press('tab')
    pyautogui.typewrite(str(excel_data['gz'][element]), interval=time_between_keystrokes)
    pyautogui.press('tab')
    pyautogui.typewrite(str(excel_data['qz'][element]), interval=time_between_keystrokes)
    pyautogui.press('tab')
    pyautogui.typewrite(str(excel_data['Mgzk'][element]), interval=time_between_keystrokes)
    pyautogui.press('tab')
    pyautogui.typewrite(str(excel_data['Mqzk'][element]), interval=time_between_keystrokes)
    pyautogui.press('tab')
    pyautogui.typewrite(str(excel_data['Mgyk'][element]), interval=time_between_keystrokes)
    pyautogui.press('tab')
    pyautogui.typewrite(str(excel_data['Mqyk'][element]), interval=time_between_keystrokes)
    pyautogui.press('tab')
    pyautogui.typewrite(str(excel_data['Mgzf'][element]), interval=time_between_keystrokes)
    pyautogui.press('tab')
    pyautogui.typewrite(str(excel_data['Mqzf'][element]), interval=time_between_keystrokes)
    pyautogui.press('tab')
    pyautogui.typewrite(str(excel_data['Mgyf'][element]), interval=time_between_keystrokes)
    pyautogui.press('tab')
    pyautogui.typewrite(str(excel_data['Mqyf'][element]), interval=time_between_keystrokes)
    pyautogui.press('tab')
    pyautogui.typewrite(str(excel_data['ey'][element]), interval=time_between_keystrokes)
    pyautogui.press('tab')
    pyautogui.typewrite(str(excel_data['ez'][element]), interval=time_between_keystrokes)
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.typewrite(str(excel_data['Bewehrung'][element]), interval=time_between_keystrokes)
    pyautogui.press('tab')

    # simply tabs and enters data until it stops on calculate


def find_and_click(image, offset_x=0, offset_y=0, if_grayscale=False, sleep=0):
    """A function used multiple times in the software. Its' purpose is to find a specific .png file and then click it."""
    def_loc = None  # image location
    time.sleep(sleep)  # sometimes it clicks to fast for Frilo to process the data (roll out windows!) so it's useful to let it sleep for a while
    while True:
        t_end = time.time() + time_to_wait
        while time.time() < t_end:
            def_loc = pyautogui.locateCenterOnScreen(image, grayscale=if_grayscale)
            if def_loc is not None:
                pyautogui.click(def_loc[0] + offset_x, def_loc[1] + offset_y, duration=time_to_move_mouse)
                break
        if def_loc is None:
            not_found_button = pyautogui.confirm(text=(image + ' not found. Please make it visible and let the software click it or close the software'), title='Button not found', buttons=['Done', 'Exit software'])
            if not_found_button == 'Exit software':
                sys.exit(0)
        else:
            break


def set_reinforcement_diameter(which_reinforcement, which_arrows, element):
    """A function that sets up the reinforcement diameter. It works quite simply: open roll out window, pick the lowest diameter. If its not the intended diameter, open roll out window again, find the diameter and click it"""
    pyautogui.click(which_arrows, duration=time_to_move_mouse)
    time.sleep(0.2)
    pyautogui.click(which_arrows[0], which_arrows[1] + 25, duration=time_to_move_mouse)
    time.sleep(0.2)
    pyautogui.press('enter')
    pyautogui.press('enter')
    pyautogui.press('enter')

    if str(excel_data[which_reinforcement][element]) != '6':
        pyautogui.click(which_arrows, duration=time_to_move_mouse)
        time.sleep(0.2)
        find_and_click('src/' + 'd' + str(excel_data[which_reinforcement][element]) + '.png', sleep=1)
        time.sleep(0.2)
        pyautogui.press('enter')
        pyautogui.press('enter')
        pyautogui.press('enter')


def find_blue_squares():
    """Just quickly finds 'blue squares' - parts to click on the sidebar to set up data"""
    global blue_squares
    blue_squares = list(pyautogui.locateAllOnScreen('src/BlueSquare.png', region=(0, 0, 200, height - 1)))  # here its finally done as a list and not as a generator


def def_reinforcement_window(element):
    """Function that set ups the reinforcement window :) """
    pyautogui.doubleClick(blue_squares[2][0], blue_squares[2][1])
    pyautogui.press('enter')  # opens the window, accepts popups

    while True:
        t_end = time.time() + time_to_wait
        while time.time() < t_end:
            def_loc = pyautogui.locateCenterOnScreen('src/cnom.png', grayscale=True)
            if def_loc is not None:
                break
        if def_loc is None:
            not_found_button = pyautogui.confirm(text='cnom.png not found. Please make it visible and let the software click it or close the software', title='Button not found', buttons=['Done', 'Exit software'])
            if not_found_button == 'Exit software':
                sys.exit(0)
        else:
            break

    reinforcement_size_arrows = list(pyautogui.locateAllOnScreen('src/Reinforcement_Size_Arrow.png'))  # finds all of the roll out windows

    if element == 0 or excel_data['Longitudinal reinforcement diameter'][element] != excel_data['Longitudinal reinforcement diameter'][element - 1]:
        set_reinforcement_diameter('Longitudinal reinforcement diameter', (reinforcement_size_arrows[-3][0], reinforcement_size_arrows[-3][1]), element)
        # sets the longitidunal reinforcmeent

    pyautogui.click(reinforcement_size_arrows[-2][0], reinforcement_size_arrows[-2][1], duration=time_to_move_mouse)  # sets up the stirrups
    time.sleep(0.2)
    pyautogui.click(reinforcement_size_arrows[-2][0], reinforcement_size_arrows[-2][1] + 25, duration=time_to_move_mouse)
    time.sleep(0.2)
    pyautogui.press('enter')
    pyautogui.press('enter')
    pyautogui.press('enter')

    if element == 0 or excel_data['Stirrups diameter'][element] != excel_data['Stirrups diameter'][element - 1]:
        if str(excel_data['Stirrups diameter'][element]) != '6':
            pyautogui.click(reinforcement_size_arrows[-2][0], reinforcement_size_arrows[-2][1], duration=time_to_move_mouse)
            time.sleep(0.2)
            find_and_click('src/' + 'd' + str(excel_data['Stirrups diameter'][element]) + '.png', sleep=1)
            time.sleep(0.2)
            pyautogui.press('enter')
            pyautogui.press('enter')
            pyautogui.press('enter')

    if element == 0 or excel_data['Additional reinforcement diameter'][element] != excel_data['Additional reinforcement diameter'][element - 1]:
        set_reinforcement_diameter('Additional reinforcement diameter', (reinforcement_size_arrows[-1][0], reinforcement_size_arrows[-1][1]), element)
        # sets up the additional reinofcement

    find_and_click('src/cnom.png', offset_x=55, if_grayscale=True)
    if element == 0 or str(excel_data['cnom'][element]) != str(excel_data['cnom'][element - 1]):
        pyautogui.typewrite(str(excel_data['cnom'][element]), interval=time_between_keystrokes)
        pyautogui.press('enter')
        pyautogui.press('enter')
    else:
        pyautogui.typewrite(str(float(excel_data['cnom'][element]) + 0.1), interval=time_between_keystrokes)
        pyautogui.press('enter')
        pyautogui.press('enter')
        find_and_click('src/cnom.png', offset_x=55, if_grayscale=True)
        pyautogui.typewrite(str(excel_data['cnom'][element]), interval=time_between_keystrokes)
        pyautogui.press('enter')
        pyautogui.press('enter')
    find_and_click('src/Reinforcement_Arrow.png', offset_x=-50, if_grayscale=True)
    pyautogui.typewrite(str(excel_data['As'][element]), interval=time_between_keystrokes)
    pyautogui.press('enter')
    while True:
        # clicks the red 'f' button and then the first white one. In that matter it 'recalulates' the reinforcement in the cross-section
        t_end = time.time() + time_to_wait
        while time.time() < t_end:
            F_loc = pyautogui.locateCenterOnScreen('src/Reinforcement_F.png')
            F_Yellow_loc = pyautogui.locateCenterOnScreen('src/Reinforcement_F_Yellow.png')
            if F_loc is not None:
                pyautogui.click(F_loc, duration=time_to_move_mouse)
                pyautogui.press('esc')
                break
            if F_Yellow_loc is not None:
                break
        if F_loc is None and F_Yellow_loc is None:
            not_found_button = pyautogui.confirm(text='Reinforcement_F.png or Reinforcement_F_Yellow.png not found. Please make it visible and let the software click it or close the software', title='Button not found', buttons=['Done', 'Exit software'])
            if not_found_button == 'Exit software':
                sys.exit(0)
        else:
            break

    time.sleep(0.5)
    find_and_click('src/Reinforcement_Bar.png', if_grayscale=False)
    pyautogui.press('enter')
    pyautogui.press('enter')

    find_and_click('src/Isolines_Temperature.png', if_grayscale=True)
    pyautogui.press('enter')
    while True:
        t_end = time.time() + time_to_wait
        while time.time() < t_end:
            arrow = pyautogui.locateCenterOnScreen('src/GreenCalculateArrow.png')
            if arrow is not None:
                break
        if arrow is None:
            not_found_button = pyautogui.confirm(text='GreenCalculateArrow.png not found. Please make it visible and let the software click it or close the software', title='Button not found', buttons=['Done', 'Exit software'])
            if not_found_button == 'Exit software':
                sys.exit(0)
        else:
            break
    pyautogui.press('enter')

    find_and_click('src/Reinf_OK.png', if_grayscale=True)
    pyautogui.press('enter')


def compare_screen_changes(screen_before, screen_after):
    """Because saving files lags so much due to slow servers, the function checks, if a new window has already poped"""
    pixels_before_click = [screen_before.getpixel((int(width / a), int(height / a))) for a in range(2, 100)]  # creates a list of pixels from a screenshot before save initiation
    pixels_after_click = [screen_after.getpixel((int(width / a), int(height / a))) for a in range(2, 100)]  # creates a list of pixels from a screenshot after save initiation
    different_pixels_number = 0
    for pixel in range(len(pixels_before_click)):
        if pixels_before_click[pixel] != pixels_after_click[pixel]:
            different_pixels_number += 1
    return different_pixels_number  # calculates number of differences between those lists


def save_file(element):
    """A funtction that saves the calc file"""
    time.sleep(1)
    screenshot_before_save_click = pyautogui.screenshot()  # takes a screenshot before initializing save
    time.sleep(1)
    pyautogui.click(20, 35, duration=time_to_move_mouse)  # click data tab
    pyautogui.click(70, 130, duration=time_to_move_mouse)  # click save under tab
    while True:
        t_end = time.time() + 90
        while time.time() < t_end:
            screenshot_after_save_click = pyautogui.screenshot()  # takes a screenshot after initiazing save
            different_pixels_number = compare_screen_changes(screenshot_before_save_click, screenshot_after_save_click)  # calculates number of differences between the screenshots
            if different_pixels_number > 4:
                break  # if there are more than 5 differences, its assumed that a new (save) window has appered
        if different_pixels_number < 4:
            not_found_button = pyautogui.confirm(text='Save window not shown. Please open save window', title='Save window not shown', buttons=['Done', 'Exit software'])
            if not_found_button == 'Exit software':
                sys.exit(0)
        else:
            break
    time.sleep(15)  # waits 15 sec more just in case - saving on the server is lagging a lot
    save_semicolons = list(pyautogui.locateAllOnScreen('src/Save_Semicolon.png'))
    pyautogui.click(save_semicolons[2][0] + 100, save_semicolons[2][1])
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.typewrite(str(excel_data['Element Name'][element]), interval=time_between_keystrokes)  # write file name
    pyautogui.press('enter')  # confirms
    time.sleep(5)
    pyautogui.press('left')  # makes sure it works if the name exists already
    pyautogui.press('left')  # makes sure it works if the name exists already
    pyautogui.press('enter')  # confirm that position will be saved again
    while True:
        t_end = time.time() + 90
        while time.time() < t_end:
            screenshot_after_save_click = pyautogui.screenshot()  # takes a screenshot after initiazing save
            different_pixels_number = compare_screen_changes(screenshot_before_save_click, screenshot_after_save_click)  # calculates number of differences between the screenshots
            if different_pixels_number < 3:
                break  # if there are more than 5 differences, its assumed that a new (save) window has appered
        if different_pixels_number > 3:
            not_found_button = pyautogui.confirm(text='Save not done. Please close save window', title='Save window not shown', buttons=['Done', 'Exit software'])
            if not_found_button == 'Exit software':
                sys.exit(0)
        else:
            break
    time.sleep(5)


def read_excel_data():
    """Read the calulation data from excel. For the sake of simplicity, the excel needs to be called and saved accordingly"""
    try:
        return(pd.read_excel('Frilo_Data.xlsx', 'Data', index_col=None, na_values=['NA']))
    except:
        pyautogui.alert(text='Please place the proper excel file (Frilo_Data.xlsx) in the software .exe root folder', title='Excel not found.', button='Close the software.')
        sys.exit(0)


def init_settings():
    """Read the setting data from excel. For the sake of simplicity, the excel needs to be called and saved accordingly"""
    global time_between_keystrokes
    global time_to_move_mouse
    global printout_save_path
    global time_to_wait
    time_between_keystrokes = 0  # initial value of 0
    time_to_move_mouse = 0.05  # initial value of 0.05
    try:
        settings = pd.read_excel('Frilo_Data.xlsx', 'Settings', index_col=None, na_values=['NA'])
        # obviously it crashes, if data is not full due to 'NA'.
        time_between_keystrokes = settings['Time between keystrokes [ms]'][0] / 1000
        printout_save_path = str(settings['Printout report save destination'][0])
        time_to_wait = settings['Time to look for a .png [ms]'][0] / 1000
        if settings['Time to move mouse [ms]'][0] / 1000 > 0.05:
            time_to_move_mouse = settings['Time to move mouse [ms]'][0] / 1000
        # tries to read a sheet with settings
    except:
        pyautogui.alert(text='Please place the proper excel file (Frilo_Data.xlsx) in the software .exe root folder', title='Excel not found.', button='Close the software.')
        sys.exit(0)


def perform_calcs(element):
    """A function that starts all of the calcs"""
    if element == 0 or (str(excel_data['Fire calculation method'][element]) != str(excel_data['Fire calculation method'][element - 1]) and
                        str(excel_data['Class of fire resistance'][element]) != str(excel_data['Class of fire resistance'][element - 1])):
        setup_fire_design(element)
    def_values(element)
    def_reinforcement_window(element)
    while True:
        t_end = time.time() + time_to_wait
        while time.time() < t_end:
            calculate_button = pyautogui.locateCenterOnScreen('src/Calculate.png')  # tries to find red from calculate button
            if calculate_button is not None:
                break
        if calculate_button is None:
            not_found_button = pyautogui.confirm(text='Calculate.png not found. Please make it visible and let the software click it or close the software', title='Button not found', buttons=['Done', 'Exit software'])
            if not_found_button == 'Exit software':
                sys.exit(0)
        else:
            break
    pyautogui.click(calculate_button, duration=time_to_move_mouse)
    t_end = time.time() + time_to_wait * 60
    while time.time() < t_end:
        im = pyautogui.screenshot()
        if im.getpixel(calculate_button) != (255, 0, 0):
            break
        else:
            ok_loc = pyautogui.locateCenterOnScreen('src/Calculate_OK.png', grayscale=True)
            if ok_loc is not None:
                pyautogui.click(ok_loc, duration=time_to_move_mouse)
        time.sleep(1)  # repeats sleep for a second as long as the calculate button is red - if its not red it means that the calcs are finished
    time.sleep(2)
    t_end = time.time() + 15
    loc_ok = 'Default'
    while time.time() < t_end:
        loc_ok = pyautogui.locateCenterOnScreen('src/OK.png', grayscale=True)
        if loc_ok is None:
            break
        elif loc_ok != 'Default':
            pyautogui.click(loc_ok, duration=time_to_move_mouse)
            break


def printout_report(element):
    time.sleep(15)
    """A function that prepares the printout report. EPrint printer chosen as its popular and doesn't create a popout pdf as default"""
    pyautogui.hotkey('ctrl', 'p')
    time.sleep(10)
    hwnd = win32gui.GetForegroundWindow()  # gets the active window
    win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)  # maximizes the active window
    if pyautogui.locateCenterOnScreen('src/Print_EPrint_Def.png') is not None:
        find_and_click('src/Print_OK.png', if_grayscale=True)
    else:
        find_and_click('src/Print_Arrow.png', if_grayscale=True)
        time.sleep(1)
        while True:
            t_end = time.time() + time_to_wait
            while time.time() < t_end:
                print_clicked_loc = pyautogui.locateCenterOnScreen('src/Print_EPrint_List_Clicked.png')
                print_unclicked_loc = pyautogui.locateCenterOnScreen('src/Print_EPrint_List.png')
                if print_clicked_loc is not None:
                    break
                elif print_unclicked_loc is not None:
                    pyautogui.click(print_unclicked_loc, duration=time_to_move_mouse)
                    break
            if print_clicked_loc is None and print_unclicked_loc is None:
                not_found_button = pyautogui.confirm(text='Print_EPrint_List_Clicked or Print_EPrint_List.png not found. Please make it visible and let the software click it or close the software', title='Button not found', buttons=['Done', 'Exit software'])
                if not_found_button == 'Exit software':
                    sys.exit(0)
            else:
                break
        find_and_click('src/Print_OK.png', if_grayscale=True)
        pyautogui.click()

    find_and_click('src/Print_Destination_Arrow.png')
    pyautogui.typewrite(printout_save_path, interval=time_between_keystrokes)
    pyautogui.press('enter')
    find_and_click('src/Print_Name.png', offset_y=-5, if_grayscale=True)
    pyautogui.typewrite(str(excel_data['Element Name'][element]), interval=time_between_keystrokes)
    pyautogui.press('enter')
    pyautogui.press('enter')


def setup_fire_design(element):
    """A function that sets up the fire part of the software. Works on similar basis a previous functions"""
    pyautogui.doubleClick(blue_squares[1][0], blue_squares[1][1], duration=time_to_move_mouse)
    time.sleep(1)
    hwnd = win32gui.GetForegroundWindow()  # gets the active window
    win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)  # maximizes the active window
    find_and_click('src/Fire_Unmarked_Circle.png', if_grayscale=True)
    for click in range(2):
        pyautogui.press('up')
    pyautogui.moveRel(-30, 0)
    time.sleep(1)
    if str(excel_data['Fire calculation method'][element]) == 'Method A':
        find_and_click('src/Fire_Unmarked_Circle.png', if_grayscale=True)
    elif str(excel_data['Fire calculation method'][element]) == 'FEM':
        find_and_click('src/Fire_Unmarked_Circle.png', if_grayscale=True)
        pyautogui.press('down')

    if str(excel_data['Class of fire resistance'][element]) != '0':
        while True:
            t_end = time.time() + time_to_wait
            while time.time() < t_end:
                fire_arrow_loc = pyautogui.locateCenterOnScreen('src/Fire_Arrow.png')
                if fire_arrow_loc is not None:
                    pyautogui.click(fire_arrow_loc, duration=time_to_move_mouse)
                    break
            if fire_arrow_loc is None:
                not_found_button = pyautogui.confirm(text='Fire_Arrow.png not found. Please make it visible and let the software click it or close the software', title='Button not found', buttons=['Done', 'Exit software'])
                if not_found_button == 'Exit software':
                    sys.exit(0)
            else:
                break
        time.sleep(1)
        pyautogui.click(fire_arrow_loc[0], fire_arrow_loc[1] + 15, duration=time_to_move_mouse)
        if str(excel_data['Class of fire resistance'][element]) != 'R30':
            pyautogui.click(fire_arrow_loc, duration=time_to_move_mouse)
            find_and_click('src/' + str(excel_data['Class of fire resistance'][element]) + '.png', sleep=1)
    time.sleep(1)
    find_and_click('src/Fire_OK.png', sleep=0.5)


is_rectangle_set = None
is_circle_set = None
is_ring_set = None

check_access()
width, height = maximize_window()
excel_data = read_excel_data()
init_settings()
find_blue_squares()
clean_sidebar()


for element in range(len(excel_data.index)):
    # performs the process for all of the excel data rows
    perform_calcs(element)
    time.sleep(1)
    save_file(element)
    time.sleep(1)
    printout_report(element)
    time.sleep(1)
pyautogui.alert(text='Calculations finished.', title='Calculations finished.', button='Done!')
