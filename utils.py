import math
from io import BytesIO
import win32clipboard
from PIL import Image
from settings import *

def confirmation_input(ask_str, ask_type):
    if ask_type not in ['Y/n', 'y/N', 'N/y', 'n/Y']:
        ask_type = 'Y/n'
    ask_str = f"{ask_str} [{ask_type}]: "
    while True:
        ask_value = input(ask_str).lower()
        if not ask_value in [''] + TRUTHY + FALSY:
            print("Please provide a valid confirmation!")
            continue
        if ask_type == 'Y/n':
            return ask_value in [''] + TRUTHY
        elif ask_type == 'y/N':
            return ask_value in TRUTHY
        elif ask_type == 'N/y':
            return ask_value in TRUTHY
        elif ask_type == 'n/Y':
            return ask_value in [''] + TRUTHY


def roundoff(val):
    if (float(val) % 1) >= 0.5:
        val = math.ceil(val)
    else:
        val = round(val)
    return val

def chunker(seq, size):
    if size<=0:
        size = len(seq)
    for pos in range(0, len(seq), size):
        yield seq.iloc[pos:pos + size] 


def send_to_clipboard(clip_type, data):
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(clip_type, data)
    win32clipboard.CloseClipboard()

def send_image_to_clipboard(image_dir, image_name):
    image_path = image_dir / image_name
    if image_path.exists():
        image = Image.open(image_path)
        output = BytesIO()
        image.convert("RGB").save(output, "BMP")
        data = output.getvalue()[14:]
        output.close()
        send_to_clipboard(win32clipboard.CF_DIB, data)
        return True
    return False
