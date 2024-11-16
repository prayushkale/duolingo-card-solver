import cv2
import os, re
import pytesseract
import pyautogui
from PIL import Image
from PIL import ImageGrab
from concurrent.futures import ProcessPoolExecutor

# Eng to Jpn dictionary
eng_to_jpn = {
    'japan': 'nihon',
    'hello': 'konnichiwa',
    'stairs': 'kaidan',
    'cafe': 'kafe',
    'ten': 'juu',
    'busy': 'isogashii',
    'sandwich': 'sandoichi',
}

# Speciify cooordinates of each box to click
click_A = [
    {'x': 366, 'y': 270}, # 1
    {'x': 366, 'y': 390}, # 2
    {'x': 366, 'y': 510}, # 3
    {'x': 366, 'y': 630}, # 4
    {'x': 366, 'y': 750}, # 5
]
click_B = [
    {'x': 593, 'y': 270}, # 1
    {'x': 593, 'y': 390}, # 2
    {'x': 593, 'y': 510}, # 3
    {'x': 593, 'y': 630}, # 4
    {'x': 593, 'y': 750}, # 5
]

# Specify the coordinates of the rectangle (left, top, right, bottom)
A = [
    [281, 234, 451, 303], # 1
    [281, 355, 451, 426], # 2
    [281, 477, 451, 548], # 3
    [281, 599, 451, 670], # 4
    [281, 721, 451, 792], # 5
]
B = [
    [508, 232, 686, 266], # 1
    [508, 355, 686, 389], # 2
    [508, 477, 686, 511], # 3
    [508, 599, 686, 633], # 4
    [508, 721, 686, 755], # 5
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
    os.remove(right_row_filename)

    # show the output images
    # cv2.imshow("Left Output In Grayscale", left_row_image_gray)
    # cv2.imshow("Right Output In Grayscale", right_row_image_gray)

    # add each word to the left_row and right_row dictionary
    left_row[row_idx] = left_row_text.strip().replace('\n', '').replace(' ', '').lower()
    right_row[row_idx] = right_row_text.strip().replace('\n', '').replace(' ', '').lower()

    print(f"Row {row_idx}: {left_row[row_idx]}, {right_row[row_idx]}")

    return left_row[row_idx], right_row[row_idx]

    # wait 3 seconds
    # cv2.waitKey(3000)
    # close the window
    # cv2.destroyAllWindows()

if __name__ == '__main__':

    with ProcessPoolExecutor() as executor:
        results = executor.map(extract_data_per_row, range(5))

    for row_idx, result in enumerate(results):
        left_row[row_idx], right_row[row_idx] = result

    print(left_row)
    print(right_row)

    # find the matching words and click the corresponding left and then right box
    for row_idx in range(5):
        if left_row[row_idx] in eng_to_jpn:
            # find index of the matching word in right_row 
            try:
                right_row_match = find_best_match(right_row, eng_to_jpn[left_row[row_idx]], 0.5)
                if right_row_match:
                    match_idx = right_row_match[0]

                    # click the corresponding left and right boxes
                    pyautogui.click(click_A[row_idx]['x'], click_A[row_idx]['y'])
                    pyautogui.click(click_B[match_idx]['x'], click_B[match_idx]['y'])

                    print(f"Row {row_idx}: {left_row[row_idx]} is matched with {right_row[match_idx]}")
                                    
                    # set right_row[match_idx] to '' to avoid matching again
                    right_row[match_idx] = ''
            
            except ValueError:
                pass


