# import libraries
import cv2
import os, re, time
import pytesseract
import numpy as np
import ctypes
import keyboard
from PIL import Image
from PIL import ImageGrab
from concurrent.futures import ProcessPoolExecutor, ThreadPoolExecutor
import threading
from difflib import SequenceMatcher

user32 = ctypes.windll.user32

# constants for mouse event
MOUSEEVENTF_MOVE = 0x0001
MOUSEEVENTF_LEFTDOWN = 0x0002
MOUSEEVENTF_LEFTUP = 0x0004

# pyautogui.FAILSAFE = False

# eng to jpn dictionary
eng_to_jpn = {
    'japan': 'nihon',
    'hello': 'konnichiwa',
    'stairs': 'kaidan',
    'cafe': 'kafe',
    'ten': 'juu',
    'busy': 'isogashii',
    'sandwich': 'sandoichi',
    'pizza': 'piza',
    'please': 'kudasai',
    'cake': 'keeki',
    'water': 'mizu',
    'karate': 'karate',
    'clothes': 'fuku',
    'abit': 'sukoshi',
    'eight': 'hachi',
    'ticket': 'kippu',
    'lively': 'nigiyaka',
    'subway': 'chikatetsu',
    'bag': 'kaban',
    'passport': 'pasupooto',
    'platform': 'hoomu',
    'jacket': 'jaketto',
    'small': 'chiisai',
    'thirty': 'han',
    'baseball': 'youkoso',
    'four': 'yon',
    "o'clock": 'ji',
    "o'clock": '',
    'noisy': 'urusai',
    'welcome': 'youkoso',
    'baseball': 'yakyuu',
    'bye': 'jaane',
    'yes': 'hai',
    'hotel': 'hoteru',
    'here': 'koko',
    'nurse': 'kangoshi',
    'six': 'roku',
    'wallet': 'saifu',
    'tv': 'terebi',
    # 'TV': 'terebi',
    # '': 'terebi',
    'news': 'nyuusu',
    'map': 'chizu',
    'curry': 'karee',
    'no': 'iie',
    'no': '',
    'person': 'hito',
    'where': 'doko',
    'airport': 'kuukou',
    'white': 'shiroi',
    'tempura': 'tenpura',
    'far': 'toi', # 'tooi',
    'mom': 'haha',
    'soccer': 'sakkaa',
    'sports': 'supootsu',
    'clean': 'kirei',
    'dad': 'chichi',
    'twelve': 'juuni',
    'hideous': 'dasai',
    'i': 'watashi',
    '': 'watashi',
    'engineer': 'enjinia',
    'bread': 'pan',
    'eleven': 'juuichi',
    'water': 'mizu',
    'thousand': 'sen',
    'howmuch': 'ikura',
    'famous': 'yuumei',
    'jazz': 'jazu',
    'often': 'yoku',
    'very': 'totemo',
    'plays': 'shimasu',
    'hundred': 'hyaku',
    'tennis': 'tenisu',
    'yen': 'en',
    'yen': '',
    'big': 'okii',
    'quiet': 'shizuka',
    'saturday': 'doyoubi',
    'eats': 'tabemasu',
    'coffee': 'koohii',
    'rock': 'rokku',
    'weekend': 'shuumatsu',
    't-shirt': 'tiishatsu',
    'ramen': 'raamen',
    'skirt': 'sukaato',
    'smart': 'atamagaii',
    'movies': 'iga',
    'red': 'akai',
    'coat': 'kooto',
    'myson': 'musuko',
    'old': 'furui',
    'books': 'hon',
    'manga': 'manga',
    'anime': 'anime',
    'yoga': 'yoga',
    'lunch': 'hirugohan',
    'drama': 'dorama',
    'soba': 'soba',
    'transfer': 'norikae',
    'music': 'ongaku',
    'family': 'kazoku',
    'mywife': 'tsuma',
    'dress': 'doresu',
    'five': 'go',
    'there': 'soko',
    'store': 'mise',
    'exit': 'deguchi',
    'taxi': 'takushii',
    'dinner': 'bangohan',
    'new': 'atarashii',
    'restroom': 'otearai',
    'close': 'chikai',
    'sunday': 'nichiyoubi',
    'juice': 'juusu',
    'j-pop': 'poppu',
    'phone': 'denwa',
    'elevator': 'ereebeeta',
    'udon': 'udon',
    'video': 'douga',
    'outlet': 'konsento',
    'judo': 'juudou',
    'nice': 'yasashii',
    'radio': 'rajio',
    'tomorrow': 'ashita',
    'pasta': 'pasuta',
    'waltet': 'saifu',
    'bye': 'jaane',
    'vietnam': 'betonamu',
    'comedy': 'komedi',
    'fantasy': 'fantaji',
    'mystery': 'misuterii',
    'k-pop': 'poppu',
    'action': 'akushon',
    'desk': 'tsukue',
    'spacious': 'hiroi',
    'cramped': 'semai',
    'kitchen': 'kicchin',
    'probably': 'tabun',
    'house': '',
    'fridge': 'reizouko',
    'long': 'nagai',
    'fast': 'hayai',
    'train': 'densha',
    'station': 'eki',
    'hour': 'jikan',
    'bus': 'basu',
    'about': 'gurai',
    'next': 'tsugi',
}

# specify cooordinates of each box to click
# left rows
click_A = [
    {'x': 250, 'y': 500}, # A1
    {'x': 250, 'y': 645}, # A2
    {'x': 250, 'y': 780}, # A3
    {'x': 250, 'y': 920}, # A4
    {'x': 250, 'y': 1060}, # A5
]
# right rows
click_B = [
    {'x': 550, 'y': 510}, # B1
    {'x': 550, 'y': 640}, # B2
    {'x': 550, 'y': 780}, # B3
    {'x': 550, 'y': 920}, # B4
    {'x': 550, 'y': 1060}, # B5
]

# specify the coordinates of the rectangle [top_left_x, top_left_y, bottom_right_x, bottom_right_y]
# left rows
A = [
    [60, 465, 250, 505], #A1
    [60, 600, 250, 645], #A2
    [60, 740, 250, 785], #A3
    [60, 880, 250, 925], #A4
    [60, 1015, 250, 1060], #A5
]
# right rows
B = [
    [360, 450, 540, 480], # 1
    [360, 575, 540, 620], # 2
    [360, 715, 540, 760], # 3
    [360, 850, 540, 898], # 4
    [360, 990, 540, 1030], # 5
]

# continue btn coordinates [top_left_x, top_left_y, bottom_right_x, bottom_right_y]
continue_btn_box = [40, 1220, 560, 1270]
# continue btn click coordinates
continue_btn = {
    'x': 300,
    'y': 1240
}


def click_at(x, y):
    """
    Simulates a mouse click at the given coordinates (x, y).

    Args:
        x (int): The x-coordinate.
        y (int): The y-coordinate.
    """
    # Set cursor position
    user32.SetCursorPos(x, y)
    time.sleep(0.01)
    
    # Simulate mouse left button down and up (click)
    user32.mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    time.sleep(0.02)  # Short delay for click effect
    user32.mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)


def find_similar_values(dictionary, target_value, threshold=0.9):
    """
    Finds and returns a list of items from the given dictionary that have values
    similar to the target value, based on a similarity threshold.

    Args:
        dictionary (dict): The dictionary to search for similar values.
        target_value (str): The target value to compare against.
        threshold (float): The minimum similarity ratio required for a match 
                           (default is 0.9).

    Returns:
        list: A list of tuples, each containing a key, value, and similarity percentage,
              sorted by similarity in descending order.
    """

    matches = []
    
    for key, value in dictionary.items():
        similarity = SequenceMatcher(None, value.lower(), target_value.lower()).ratio()
        if similarity >= threshold:
            matches.append((key, value, round(similarity * 100, 2)))
    return sorted(matches, key=lambda x: x[2], reverse=True)
    

def find_best_match(dictionary, target_value, threshold=0.9):
    """
    Finds and returns the best matching key, value pair from the given dictionary that has a value
    similar to the target value, based on a similarity threshold.

    Args:
        dictionary (dict): The dictionary to search for similar values.
        target_value (str): The target value to compare against.
        threshold (float): The minimum similarity ratio required for a match (default is 0.9).

    Returns:
        tuple or None: A tuple containing the best matching key, value, and similarity percentage,
                       or None if no match was found.
    """
    matches = find_similar_values(dictionary, target_value, threshold)
    return matches[0] if matches else None


# detected values
left_row = {
    0: '',
    1: '',
    2: '',
    3: '',
    4: '',
}
right_row = {
    0: '',
    1: '',
    2: '',
    3: '',
    4: '',
}

# previous matches of rows
previous_matches = {
    0: '',
    1: '',
    2: '',
    3: '',
    4: '',
}
# previous matches trial count
previous_matches_trial_count = {
    0: 0,
    1: 0,
    2: 0,
    3: 0,
    4: 0,
}

# detected continue button text
continue_btn_text = ''

# continue click count to check ongoing round
continue_click_count = 1

def extract_left_row_box(row_idx):
    """
    Extracts the text from the left row box at the given row index by taking a screenshot of the box, saving it, reading it with pytesseract, and returning the extracted text.
    
    Args:
        row_idx (int): The row index of the box to extract.
    
    Returns:
        str: The extracted text from the box.
    """

    # Capture the screenshot of the rectangle
    left_row_screenshot = ImageGrab.grab((A[row_idx][0], A[row_idx][1], A[row_idx][2], A[row_idx][3]))
    
    # convert to grayscale image
    left_row_image=np.array(left_row_screenshot)
    left_row_image_gray=cv2.cvtColor(left_row_image, cv2.COLOR_BGR2GRAY)
     
    # use thresh or blur
    (thresh, left_row_image_gray) = cv2.threshold(left_row_image_gray, 64, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)
    
    # Save the screenshot
    cv2.imwrite(f"A{row_idx+1}.jpg", left_row_image_gray)
	
    left_row_text = pytesseract.image_to_string(Image.open(f"A{row_idx+1}.jpg"), lang='eng')
    # left_row_text = pytesseract.image_to_string(left_row_screenshot, lang='eng')
    
    # show the output images
    # cv2.imshow("Left Output In Grayscale", left_row_image_gray)
    
    # add each word to the left_row dictionary, lowercase, remove newlines and spaces
    left_row[row_idx] = left_row_text.strip().replace('\n', '').replace(' ', '').lower()
    
    return left_row[row_idx]


def extract_right_row_box(row_idx):
    """
    Extracts the text from the right row box at the given row index by taking a screenshot of the box, saving it, reading it with pytesseract, and returning the extracted text.
    
    Args:
        row_idx (int): The row index of the box to extract.
    
    Returns:
        str: The extracted text from the box.
    """

    # Capture the screenshot of the rectangle
    right_row_screenshot = ImageGrab.grab((B[row_idx][0], B[row_idx][1], B[row_idx][2], B[row_idx][3]))
    
    # Save the screenshot
    right_row_screenshot.save(f"B{row_idx+1}.jpg")
    
    right_row_image=np.array(right_row_screenshot)
    right_row_image_gray=cv2.cvtColor(right_row_image, cv2.COLOR_BGR2GRAY)
     
    # use thresh or blur
    (thresh, right_row_image_gray) = cv2.threshold(right_row_image_gray, 64, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)
    
    # save the screenshot
    cv2.imwrite(f"B{row_idx+1}.jpg", right_row_image_gray)
    
    right_row_text = pytesseract.image_to_string(Image.open(f"B{row_idx+1}.jpg"), lang='eng')
    # right_row_text = pytesseract.image_to_string(right_row_screenshot, lang='eng')
    
    # show the output images
    # cv2.imshow("Right Output In Grayscale", right_row_image_gray)
    
    # add each word to the right_row dictionary, lowercase, remove newlines and spaces
    right_row[row_idx] = right_row_text.strip().replace('\n', '').replace(' ', '').lower()
    
    return right_row[row_idx]


def extract_data_per_row(row_idx):
    """
    Extracts data from both the left and right row boxes at the given row index.
    It captures the text from each box, processes it, and stores it in the respective 
    dictionaries: left_row and right_row.

    Args:
        row_idx (int): The row index for which data should be extracted.

    Returns:
        tuple: A tuple containing the extracted text from the left and right row boxes.
    """

    left_row[row_idx] = extract_left_row_box(row_idx)
    right_row[row_idx] = extract_right_row_box(row_idx)

    return left_row[row_idx], right_row[row_idx]


def extract_continue_box():
    """
    Extracts text from the continue button by taking a screenshot of the button area,
    saving it, and using pytesseract to read and return the extracted text.

    Returns:
        str: The extracted and processed text from the continue button.
    """

    # Capture the screenshot of the rectangle
    continue_btn_screenshot = ImageGrab.grab((continue_btn_box[0], continue_btn_box[1], continue_btn_box[2], continue_btn_box[3]))
    
    # Save the screenshot
    continue_btn_screenshot.save("continue_btn.jpg")

    # get the text
    continue_btn_text = pytesseract.image_to_string(continue_btn_screenshot, lang='eng')
    
    # add each word to the right_row dictionary, lowercase, remove newlines and spaces
    continue_btn_text = continue_btn_text.strip().replace('\n', '').replace(' ', '').lower()

    return continue_btn_text


# maintain a list of not found words, to be shown at the end
not_found = []

# match count
matched_count = {
    'count': 0
}

# continue click count, for round continuation
continue_click_count = 1

def click_on_matches():
    """
    Simulates a match by finding the matching words in the left and right row boxes and
    clicking the corresponding left and then right box. It also keeps track of the matched
    words to avoid matching them again.

    Maintains a list of not found words, to be shown at the end.

    It also starts extracting data for matched box in parallel after 3 seconds using
    ThreadPoolExecutor.

    :return: None
    """

    last_clicked_x = -1
    last_clicked_y = -1

    # find the matching words and click the corresponding left and then right box
    for row_idx in range(5):
        if left_row[row_idx] in eng_to_jpn and left_row[row_idx] != 'matched':
            # find index of the matching word in right_row 
            try:
                right_row_match = find_best_match(right_row, eng_to_jpn[left_row[row_idx]], 0.6)
                if right_row_match:
                    match_idx = right_row_match[0]

                    if previous_matches[row_idx] == f'{left_row[row_idx]}-{right_row[match_idx]}' and previous_matches_trial_count[row_idx] < 2:
                        previous_matches_trial_count[row_idx] += 1
                        continue

                    # click the corresponding left and right boxes
                    click_at(click_A[row_idx]['x'], click_A[row_idx]['y'])
                    # print(f'Clicked on A{row_idx+1}')

                    click_at(click_B[match_idx]['x'], click_B[match_idx]['y'])
                    # print(f'Clicked on B{match_idx+1}\n')

                    last_clicked_x = click_B[match_idx]['x']
                    last_clicked_y = click_B[match_idx]['y']

                    print(f"Row {row_idx}: {left_row[row_idx]} is matched with {right_row[match_idx]}")
                                    
                    previous_matches[row_idx] = f'{left_row[row_idx]}-{right_row[match_idx]}'
                    previous_matches_trial_count[row_idx] = 0

                    # set right_row[match_idx] to '' to avoid matching again
                    right_row[match_idx] = 'matched'
                    # also clear the left_row[row_idx] to avoid matching again
                    left_row[row_idx] = 'matched'

                    matched_count['count'] += 1
                else:
                    pass
                    # print(f"NM {row_idx}: {left_row[row_idx]}:{left_row}, {right_row}")
            
            except ValueError as e:
                print(f"Error {e}")
                pass
        else:
            if left_row[row_idx] not in not_found:
                not_found.append(left_row[row_idx])
                print(f"NF {row_idx}: {left_row[row_idx]}: {left_row}, {right_row}")

    # reset click by clicking again on the last clicked box
    if last_clicked_x != -1 and last_clicked_y != -1:
        click_at(last_clicked_x, last_clicked_y)


if __name__ == '__main__':
    print('---------------------Start-----------------------')
    clicked_start = False
    while True:
        if continue_click_count <= 3:
            with ProcessPoolExecutor() as executor:
                results = executor.map(extract_data_per_row, range(5))
                click_on_matches()

            for row_idx, result in enumerate(results):
                left_row[row_idx], right_row[row_idx] = result

            # check if atleast 3 words in left_row and right_row are '' , if so then click on the continue button
            if matched_count['count'] > 5 and len(list(filter(lambda x: x == '', left_row.values()))) >= 2 and len(list(filter(lambda x: x == '', right_row.values()))) >= 2:
                # get continue button text
                continue_btn_text = extract_continue_box()
                
                if SequenceMatcher(None, continue_btn_text.lower(), 'continue').ratio() >= 0.85:
                    # click on the continue button
                    click_at(continue_btn['x'], continue_btn['y'])
                    continue_click_count += 1
                    matched_count['count'] = 0

                    if continue_click_count < 4:
                        print(f'--------------------Round {continue_click_count}-----------------------')

        if continue_click_count > 3:
            if continue_click_count == 4:
                print('--------------------End-----------------------')
                continue_click_count += 1
            
            continue_btn_text = extract_continue_box()
            
            if not clicked_start and SequenceMatcher(None, continue_btn_text.lower(), 'start+225xp').ratio() >= 0.5:
                # click on the continue button
                click_at(continue_btn['x'], continue_btn['y'])
                continue_click_count += 1
                clicked_start = True
            
            if clicked_start:
                if SequenceMatcher(None, continue_btn_text.lower(), 'continue').ratio() >= 0.85:
                    continue_click_count = 1
                    matched_count['count'] = 0
                    clicked_start = False
                    print('---------------------Start-----------------------')
                    time.sleep(2)
                    click_at(continue_btn['x'], continue_btn['y'])

        if keyboard.is_pressed('x'):
            print('Exiting...')
            print("Not Found: ", not_found)
            print("matched_count: ", matched_count['count'])
            break