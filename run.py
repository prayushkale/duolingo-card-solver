import cv2
import os, re, time
import pytesseract
import pyautogui
import keyboard
from PIL import Image
from PIL import ImageGrab
from concurrent.futures import ProcessPoolExecutor

pyautogui.FAILSAFE = False

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
    'news': 'nyuusu',
    'map': 'chizu',
    'curry': 'karee',
    'no': 'iie',
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
}

# Speciify cooordinates of each box to click
# click_A = [
#     {'x': 318, 'y': 383}, # 1
#     {'x': 318, 'y': 473}, # 2
#     {'x': 318, 'y': 563}, # 3
#     {'x': 318, 'y': 653}, # 4
#     {'x': 318, 'y': 742}, # 5
# ]
# click_B = [
#     {'x': 534, 'y': 383}, # 1
#     {'x': 534, 'y': 473}, # 2
#     {'x': 534, 'y': 563}, # 3
#     {'x': 534, 'y': 653}, # 4
#     {'x': 534, 'y': 742}, # 5
# ]

# # Specify the coordinates of the rectangle (left, top, right, bottom)
# A = [
#     [228, 357, 407, 412], # 1
#     [228, 445, 407, 502], # 2
#     [228, 537, 407, 592], # 3
#     [228, 627, 407, 687], # 4
#     [228, 718, 407, 774], # 5
# ]
# B = [
#     [447, 357, 629, 385], # 1
#     [447, 445, 629, 473], # 2
#     [447, 537, 629, 563], # 3
#     [447, 627, 629, 655], # 4
#     [447, 718, 629, 745], # 5
# ]


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
    [1095, 520, 1270, 575], # 1
    [1095, 610, 1270, 665], # 2
    [1095, 700, 1270, 755], # 3
    [1095, 790, 1270, 845], # 4
    [1095, 880, 1270, 935], # 5
]
B = [
    [1310, 520, 1490, 545], # 1
    [1310, 610, 1490, 640], # 2
    [1310, 700, 1490, 730], # 3
    [1310, 790, 1490, 820], # 4
    [1310, 880, 1490, 910], # 5
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
    0: None,
    1: None,
    2: None,
    3: None,
    4: None,
}
right_row = {
    0: None,
    1: None,
    2: None,
    3: None,
    4: None,
}


capture_count = 1

def extract_data_per_row(row_idx):
    # Capture the screenshot of the rectangle
    left_row_screenshot = ImageGrab.grab((A[row_idx][0], A[row_idx][1], A[row_idx][2], A[row_idx][3]))
    right_row_screenshot = ImageGrab.grab((B[row_idx][0], B[row_idx][1], B[row_idx][2], B[row_idx][3]))

    # Save the screenshot
    left_row_screenshot.save(f"A{row_idx+1}.png")
    right_row_screenshot.save(f"B{row_idx+1}.png")

    # We then read the image with text
    left_row_image=cv2.imread(f'A{row_idx+1}.png')
    right_row_image=cv2.imread(f'B{row_idx+1}.png')

    # convert to grayscale image
    left_row_image_gray=cv2.cvtColor(left_row_image, cv2.COLOR_BGR2GRAY)
    right_row_image_gray=cv2.cvtColor(right_row_image, cv2.COLOR_BGR2GRAY)

    # use thresh or blur
    cv2.threshold(left_row_image_gray, 0,255,cv2.THRESH_BINARY| cv2.THRESH_OTSU)[1]
    cv2.threshold(right_row_image_gray, 0,255,cv2.THRESH_BINARY| cv2.THRESH_OTSU)[1]
    # cv2.medianBlur(gray, 3)
	
    # memory usage with image i.e. adding image to memory
    left_row_filename = "{}.jpg".format(os.getpid())
    cv2.imwrite(left_row_filename, left_row_image_gray)
    left_row_text = pytesseract.image_to_string(Image.open(left_row_filename), lang='eng')
    os.remove(left_row_filename)

    right_row_filename = "{}.jpg".format(os.getpid())
    cv2.imwrite(right_row_filename, right_row_image_gray)
    right_row_text = pytesseract.image_to_string(Image.open(right_row_filename), lang='eng')
    # print(right_row_text)
    os.remove(right_row_filename)

    # show the output images
    # cv2.imshow("Left Output In Grayscale", left_row_image_gray)
    # cv2.imshow("Right Output In Grayscale", right_row_image_gray)

    # add each word to the left_row and right_row dictionary
    left_row[row_idx] = left_row_text.strip().replace('\n', '').replace(' ', '').lower()
    right_row[row_idx] = right_row_text.strip().replace('\n', '').replace(' ', '').lower()

    # print(f"Row {row_idx}: {left_row[row_idx]}, {right_row[row_idx]}")

    return left_row[row_idx], right_row[row_idx]

    # wait 3 seconds
    # cv2.waitKey(3000)
    # close the window
    # cv2.destroyAllWindows()

if __name__ == '__main__':

    not_found = []
    while True:
        with ProcessPoolExecutor() as executor:
            results = executor.map(extract_data_per_row, range(5))

        for row_idx, result in enumerate(results):
            left_row[row_idx], right_row[row_idx] = result

        # print(left_row)
        # print(right_row)

        # break loop on pressing 'x'
        if keyboard.is_pressed('x'):
            break

        # find the matching words and click the corresponding left and then right box
        for row_idx in range(5):
            if left_row[row_idx] in eng_to_jpn:
                # find index of the matching word in right_row 
                try:
                    right_row_match = find_best_match(right_row, eng_to_jpn[left_row[row_idx]], 0.5)
                    if right_row_match:
                        match_idx = right_row_match[0]

                        # click the corresponding left and right boxes
                        pyautogui.click(click_A[row_idx]['x'], click_A[row_idx]['y'], button='left', interval=0.8)
                        print(f'Clicked on A{row_idx+1}')

                        time.sleep(0.3)

                        pyautogui.click(click_B[match_idx]['x'], click_B[match_idx]['y'], button='left', interval=0.8)
                        print(f'Clicked on B{match_idx+1}\n')

                        time.sleep(0.3)

                        # print(f"Row {row_idx}: {left_row[row_idx]} is matched with {right_row[match_idx]}")
                                        
                        # set right_row[match_idx] to '' to avoid matching again
                        # right_row[match_idx] = ''
                    else:
                        pass
                        # print(f"Row {row_idx}: {left_row[row_idx]} is not found, available matches: {right_row}")
                
                except ValueError:
                    pass
            else:
                if left_row[row_idx] not in not_found:
                    not_found.append(left_row[row_idx])
                    # print(f"Row {row_idx}: {left_row[row_idx]} is not matched, available matches: {right_row}")

            # next_row_index = (row_idx + 1) % 5
            # pyautogui.click(click_A[next_row_index]['x'], click_A[next_row_index]['y'])
            # print(f'\nClicked on A{next_row_index+1}\n')
        print('--------------------------------------------')