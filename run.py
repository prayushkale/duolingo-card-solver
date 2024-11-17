import cv2
import os, re, time
import pytesseract
# import pyautogui
import ctypes
import keyboard
from PIL import Image
from PIL import ImageGrab
from concurrent.futures import ProcessPoolExecutor
import threading

user32 = ctypes.windll.user32

# Constants for mouse event
MOUSEEVENTF_MOVE = 0x0001
MOUSEEVENTF_LEFTDOWN = 0x0002
MOUSEEVENTF_LEFTUP = 0x0004

def set_timeout(func, delay_seconds, *args, **kwargs):
    timer = threading.Timer(delay_seconds, func, args=args, kwargs=kwargs)
    timer.daemon = True  # Thread will exit when main program exits
    timer.start()
    return timer

def click_at(x, y):
    # Set cursor position
    user32.SetCursorPos(x, y)
    time.sleep(0.05)
    
    # Simulate mouse left button down and up (click)
    user32.mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    time.sleep(0.05)  # Short delay for click effect
    user32.mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)

def run_click(x, y):
    with ProcessPoolExecutor() as executor:
        executor.submit(click_at, x, y)

# pyautogui.FAILSAFE = False

# Eng to Jpn dictionary
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
    'TV': 'terebi',
    '': 'terebi',
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
    'far': 'tooi',
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
    'big': 'ooki',
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
    'movies': 'eiga',
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
}


# Speciify cooordinates of each box to click
click_A = [
    {'x': 1200, 'y': 540}, # 1
    {'x': 1200, 'y': 640}, # 2
    {'x': 1200, 'y': 730}, # 3
    {'x': 1200, 'y': 820}, # 4
    {'x': 1200, 'y': 910}, # 5
]
click_B = [
    {'x': 1400, 'y': 540}, # 1
    {'x': 1400, 'y': 640}, # 2
    {'x': 1400, 'y': 730}, # 3
    {'x': 1400, 'y': 820}, # 4
    {'x': 1400, 'y': 910}, # 5
]

# Specify the coordinates of the rectangle (left, top, right, bottom)
A = [
    [1095, 520, 1280, 579], # 1
    [1095, 615, 1280, 673], # 2
    [1095, 710, 1280, 767], # 3
    [1095, 805, 1280, 862], # 4
    [1095, 898, 1280, 957], # 5
]
B = [
    [1321, 520, 1511, 549], # 1
    [1321, 614, 1511, 644], # 2
    [1321, 709, 1511, 738], # 3
    [1321, 803, 1511, 832], # 4
    [1321, 898, 1511, 929], # 5
]
from difflib import SequenceMatcher

def find_similar_values(dictionary, target_value, threshold=0.9):
    matches = []
    
    for key, value in dictionary.items():
        similarity = SequenceMatcher(None, value.lower(), target_value.lower()).ratio()
        if similarity >= threshold:
            matches.append((key, value, round(similarity * 100, 2)))
    return sorted(matches, key=lambda x: x[2], reverse=True)

def find_best_match(dictionary, target_value, threshold=0.9):
    matches = find_similar_values(dictionary, target_value, threshold)
    return matches[0] if matches else None

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

previous_matches = {
    0: '',
    1: '',
    2: '',
    3: '',
    4: '',
}

capture_count = 1

def extract_left_row_box(row_idx):
    print(f"Extracting left row box {row_idx+1}")

    # Capture the screenshot of the rectangle
    left_row_screenshot = ImageGrab.grab((A[row_idx][0], A[row_idx][1], A[row_idx][2], A[row_idx][3]))
    
    # Save the screenshot
    left_row_screenshot.save(f"A{row_idx+1}.png")
    
    # We then read the image with text
    left_row_image=cv2.imread(f'A{row_idx+1}.png')
    
    # convert to grayscale image
    left_row_image_gray=cv2.cvtColor(left_row_image, cv2.COLOR_BGR2GRAY)
    
    # use thresh or blur
    cv2.threshold(left_row_image_gray, 0,255,cv2.THRESH_BINARY| cv2.THRESH_OTSU)[1]
    # cv2.medianBlur(gray, 3)
	
    # memory usage with image i.e. adding image to memory
    left_row_filename = f"A{row_idx+1}_temp.jpg"
    cv2.imwrite(left_row_filename, left_row_image_gray)
    left_row_text = pytesseract.image_to_string(Image.open(left_row_filename), lang='eng')
    os.remove(left_row_filename)

    # show the output images
    # cv2.imshow("Left Output In Grayscale", left_row_image_gray)
    
    # add each word to the left_row and right_row dictionary
    if left_row[row_idx] == '' or left_row[row_idx] == 'matched':
        left_row[row_idx] = left_row_text.strip().replace('\n', '').replace(' ', '').lower()
    
    print(f"Extracted Left Row Box {row_idx+1}: {left_row[row_idx]}")
    return left_row[row_idx]


def extract_right_row_box(row_idx):
    print(f"Extracting right row box {row_idx+1}")

    # Capture the screenshot of the rectangle
    right_row_screenshot = ImageGrab.grab((B[row_idx][0], B[row_idx][1], B[row_idx][2], B[row_idx][3]))
    
    # Save the screenshot
    right_row_screenshot.save(f"B{row_idx+1}.png")
    
    # We then read the image with text
    right_row_image=cv2.imread(f'B{row_idx+1}.png')
    
    # convert to grayscale image
    right_row_image_gray=cv2.cvtColor(right_row_image, cv2.COLOR_BGR2GRAY)
    
    # use thresh or blur
    cv2.threshold(right_row_image_gray, 0,255,cv2.THRESH_BINARY| cv2.THRESH_OTSU)[1]
    # cv2.medianBlur(gray, 3)
    
    # memory usage with image i.e. adding image to memory
    right_row_filename = f"B{row_idx+1}_temp.jpg"
    cv2.imwrite(right_row_filename, right_row_image_gray)

    right_row_text = pytesseract.image_to_string(Image.open(right_row_filename), lang='eng')
    os.remove(right_row_filename)

    # show the output images
    # cv2.imshow("Right Output In Grayscale", right_row_image_gray)
    
    # add each word to the left_row and right_row dictionary
    if right_row[row_idx] == '' or right_row[row_idx] == 'matched':
        right_row[row_idx] = right_row_text.strip().replace('\n', '').replace(' ', '').lower()
    
    print(f"Extracted Right Row Box {row_idx+1}: {right_row[row_idx]}")
    return right_row[row_idx]


def extract_data_per_row(row_idx):
    left_row[row_idx] = extract_left_row_box(row_idx)
    right_row[row_idx] = extract_right_row_box(row_idx)

    return left_row[row_idx], right_row[row_idx]

if __name__ == '__main__':
    with ProcessPoolExecutor() as executor:
        results = executor.map(extract_data_per_row, range(5))

    for row_idx, result in enumerate(results):
        left_row[row_idx], right_row[row_idx] = result

    print(left_row)
    print(right_row)

    not_found = []
    while True:
        # break loop on pressing 'x'
        if keyboard.is_pressed('x'):
            print('Not Found: ', not_found)
            break

        last_clicked_x = -1
        last_clicked_y = -1

        print('Left Row: ', left_row)
        print('Right Row: ', right_row)

        # find the matching words and click the corresponding left and then right box
        for row_idx in range(5):
            if left_row[row_idx] in eng_to_jpn and left_row[row_idx] != 'matched':
                # find index of the matching word in right_row 
                try:
                    right_row_match = find_best_match(right_row, eng_to_jpn[left_row[row_idx]], 0.8)
                    if right_row_match:
                        match_idx = right_row_match[0]

                        if previous_matches[row_idx] == f'{left_row[row_idx]}-{right_row[match_idx]}':
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

                        # set right_row[match_idx] to '' to avoid matching again
                        right_row[match_idx] = 'matched'
                        # also clear the left_row[row_idx] to avoid matching again
                        left_row[row_idx] = 'matched'

                        # start extracting data for matched box in parallel after 3 seconds
                        set_timeout(extract_left_row_box, 1, row_idx)
                        set_timeout(extract_right_row_box, 1, match_idx)
                    else:
                        pass
                        # print(f"Row {row_idx}: {left_row[row_idx]} is not found, available matches: {right_row}")
                
                except ValueError:
                    pass
            else:
                if left_row[row_idx] not in not_found:
                    not_found.append(left_row[row_idx])
                    print(f"Row {row_idx}: {left_row[row_idx]} is not matched, available matches: {right_row}")

        # reset click by clicking again on the last clicked box
        if last_clicked_x != -1 and last_clicked_y != -1:
            click_at(last_clicked_x, last_clicked_y)

        # time.sleep(1)
        # print(f'--------------------------------------------')