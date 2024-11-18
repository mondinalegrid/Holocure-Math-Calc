import pyautogui
import tkinter as tk
import keyboard
import json
from sympy import sympify
import cv2
import numpy as np
import os

# Global dictionary to store templates
templates = {}

# config creation
CONFIG_FILE = "config.txt"

# try catch for config
try:
    #with - cleans up file after, open(a,b) opens a in b mode (read,write), store in f
    with open(CONFIG_FILE, "r") as f:
        config = json.load(f)
        x, y, width, height = config["x"], config["y"], config["width"], config["height"]
except FileNotFoundError:
    # stops program from freaking out if it breaks
    x, y, width, height = 200, 100, 600, 150

def save_config():
    #Save the current region configuration to a file.
    config = {"x": x, "y": y, "width": width, "height": height}
    with open(CONFIG_FILE, "w") as f:
        json.dump(config, f)

def create_overlay(x, y, width, height):
    
    #Make box for easier number getting
    #create tkinter gui
    overlay = tk.Tk()
    overlay.overrideredirect(True)  # no borders
    overlay.attributes('-topmost', True)  # stay on top
    overlay.attributes('-alpha', 0.3)  # see through
    overlay.configure(background='red')
    update_overlay_position(overlay, x, y, width, height)
    return overlay

def update_overlay_position(overlay, x, y, width, height):
    #change overlay spot
    overlay.geometry(f"{width}x{height}+{x}+{y}")

def template_matching(bitmap_image):
    # Convert the bitmap image to a numpy array (OpenCV format)
    image = cv2.cvtColor(bitmap_image, cv2.COLOR_BGR2RGB)

    # Convert image to grayscale
    gray_image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    _, image = cv2.threshold(gray_image, 127, 255, cv2.THRESH_BINARY)

    # Path to templates
    path_to_file = os.path.join(os.getcwd(), "templates")

    # Load template images into the global dictionary
    templates.clear()
    for i in range(10):
        templates[str(i)] = cv2.imread(os.path.join(path_to_file, f"{i}.png"), cv2.IMREAD_GRAYSCALE)
    
    # Additional template images
    templates[')'] = cv2.imread(os.path.join(path_to_file, "closedp.png"), cv2.IMREAD_GRAYSCALE)
    templates['/'] = cv2.imread(os.path.join(path_to_file, "divide.png"), cv2.IMREAD_GRAYSCALE)
    templates['*'] = cv2.imread(os.path.join(path_to_file, "multiply.png"), cv2.IMREAD_GRAYSCALE)
    templates['('] = cv2.imread(os.path.join(path_to_file, "openp.png"), cv2.IMREAD_GRAYSCALE)
    templates['+'] = cv2.imread(os.path.join(path_to_file, "plus.png"), cv2.IMREAD_GRAYSCALE)
    templates['?'] = cv2.imread(os.path.join(path_to_file, "question.png"), cv2.IMREAD_GRAYSCALE)
    templates['='] = cv2.imread(os.path.join(path_to_file, "equals.png"), cv2.IMREAD_GRAYSCALE)
    templates['-'] = cv2.imread(os.path.join(path_to_file, "minus.png"), cv2.IMREAD_GRAYSCALE)

    # Perform template matching
    return extract_expression_in_order(gray_image)


def extract_expression_in_order(main_image):
    extracted_text = []
    match_results = []

    # Perform template matching for each template
    for template_key, template in templates.items():
        result = cv2.matchTemplate(main_image, template, cv2.TM_CCOEFF_NORMED)

        threshold = 0.8
        min_distance = 10  # Minimum distance to consider as non-duplicate

        while True:
            min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)

            # Check if the match is above the threshold
            if max_val >= threshold:
                # Check for duplicates based on the proximity
                is_duplicate = any(
                    abs(loc[0] - max_loc[0]) < min_distance and abs(loc[1] - max_loc[1]) < min_distance
                    for _, loc in match_results
                )

                if not is_duplicate:
                    if template_key == "-":
                        matched_width = template.shape[1]
                        matched_height = template.shape[0]
                        aspect_ratio = matched_width / matched_height
                        tolerance_width = matched_width * 0.2
                        tolerance_height = matched_height * 0.2
                        min_aspect_ratio = aspect_ratio * 0.8
                        max_aspect_ratio = aspect_ratio * 1.2

                        if (
                            abs(template.shape[1] - matched_width) < tolerance_width and
                            abs(template.shape[0] - matched_height) < tolerance_height and
                            min_aspect_ratio < aspect_ratio < max_aspect_ratio and
                            not is_near_number_bottom(max_loc, main_image)
                        ):
                            match_results.append((template_key, max_loc))
                    else:
                        match_results.append((template_key, max_loc))

                # Suppress the current match to find the next one
                cv2.rectangle(result, (max_loc[0], max_loc[1]), (max_loc[0] + template.shape[1], max_loc[1] + template.shape[0]), (0, 0, 0), -1)
            else:
                break  # No more matches above the threshold

    # Sort the match results by the x-coordinate (left to right)
    match_results.sort(key=lambda x: x[1][0])

    # Combine adjacent matches based on proximity (group digits together)
    current_group = ""
    last_x = -1

    for match_value, match_loc in match_results:
        match_x = match_loc[0]

        # If the current match is far enough from the last match, treat it as a new group
        if last_x == -1 or match_x > last_x + 20:
            if current_group:
                extracted_text.append(current_group)
            current_group = match_value  # Start a new group
        else:
            current_group += match_value  # Append to the existing group

        last_x = match_x

    if current_group:
        extracted_text.append(current_group)

    # Join the extracted text to form the expression
    return " ".join(extracted_text)


def is_near_number_bottom(location, main_image):
    # Set a tolerance for proximity to the bottom of detected numbers
    proximity_tolerance = 5

    # Path to the number templates
    path_to_file = os.path.join(os.getcwd(), "templates")
    number_templates = [f"{path_to_file}\{i}.png" for i in range(10)]

    for template_path in number_templates:
        template = cv2.imread(template_path, cv2.IMREAD_GRAYSCALE)
        result = cv2.matchTemplate(main_image, template, cv2.TM_CCOEFF_NORMED)

        threshold = 0.8
        while True:
            min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)

            if max_val >= threshold:
                detected_rect = (max_loc[0], max_loc[1], template.shape[1], template.shape[0])
                bottom_y = detected_rect[1] + detected_rect[3]

                if (
                    abs(location[1] - bottom_y) <= proximity_tolerance and
                    detected_rect[0] <= location[0] <= detected_rect[0] + detected_rect[2]
                ):
                    return True  # The location is near the bottom of a number

                # Suppress the current match to find the next one
                cv2.rectangle(result, (max_loc[0], max_loc[1]), (max_loc[0] + template.shape[1], max_loc[1] + template.shape[0]), (0, 0, 0), -1)
            else:
                break  # No more matches above the threshold

    return False


def capture_and_solve_screenshot(x, y, width, height, templates):
    #screenshot, extract the equation, solve the guy
    
    region = (x, y, width, height)
    screenshot = pyautogui.screenshot(region=region)
    screenshot = np.array(screenshot.convert("L")) 

    # match templates
    equation_text = template_matching(screenshot)
    print(f"Extracted Equation (raw): {equation_text}")

    # remove = and ? and replace -2
    equation_text = equation_text.replace("-2", "2").replace("=", "").replace("?", "").strip()
    print(f"Cleaned Equation: {equation_text}")

    try:
        # sympy the equation
        equation = sympify(equation_text)
        result = equation.evalf()
        print(f"Solution: {result}")
    except Exception as e:
        print(f"Could not solve the equation: {e}")

def main():
    global x, y, width, height
    
    overlay = create_overlay(x, y, width, height)

    print("arrow keys to move")
    print("+ / - for width")
    print("[ / ] for height")
    print("F5 to screenshot")
    print("ESC to exit")

    # move around the boy
    def move_up():
        global y
        y -= 5
        update_overlay_position(overlay, x, y, width, height)

    def move_down():
        global y
        y += 5
        update_overlay_position(overlay, x, y, width, height)

    def move_left():
        global x
        x -= 5
        update_overlay_position(overlay, x, y, width, height)

    def move_right():
        global x
        x += 5
        update_overlay_position(overlay, x, y, width, height)

    def increase_width():
        global width
        width += 5
        update_overlay_position(overlay, x, y, width, height)

    def decrease_width():
        global width
        width = max(5, width - 5)  # cannot be below 5
        update_overlay_position(overlay, x, y, width, height)

    def increase_height():
        global height
        height += 5
        update_overlay_position(overlay, x, y, width, height)

    def decrease_height():
        global height
        height = max(5, height - 5)  # same
        update_overlay_position(overlay, x, y, width, height)

    def take_screenshot_and_solve():
        capture_and_solve_screenshot(x, y, width, height, templates)

    def exit_program():
        print("Exiting")
        save_config()
        overlay.destroy()

    # hotkeys
    keyboard.add_hotkey('up', move_up)
    keyboard.add_hotkey('down', move_down)
    keyboard.add_hotkey('left', move_left)
    keyboard.add_hotkey('right', move_right)
    keyboard.add_hotkey('+', increase_width)
    keyboard.add_hotkey('-', decrease_width)
    keyboard.add_hotkey(']', increase_height)
    keyboard.add_hotkey('[', decrease_height)
    keyboard.add_hotkey('f5', take_screenshot_and_solve)
    keyboard.add_hotkey('esc', exit_program)

    # mainloop to keep the boy open
    overlay.mainloop()

    # make sure its actually closed
    keyboard.clear_all_hotkeys()

if __name__ == "__main__":
    main()
