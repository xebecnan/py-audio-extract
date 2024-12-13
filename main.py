# coding: utf-8

import os
import moviepy.editor
from pathlib import Path
import numpy as np
from scipy.io import wavfile
import pyautogui
import win32gui
import win32con
import win32api
import win32clipboard
import time
import argparse
import ctypes.wintypes

def getMyVideoDir():
    '参考 https://stackoverflow.com/questions/3858851/how-to-get-windows-special-folders-for-currently-logged-in-user'
    CSIDL_MYVIDEO = 0x000e
    SHGFP_TYPE_CURRENT= 0   # Want current, not default value
    buf= ctypes.create_unicode_buffer(ctypes.wintypes.MAX_PATH)
    ctypes.windll.shell32.SHGetFolderPathW(0, CSIDL_MYVIDEO, 0, SHGFP_TYPE_CURRENT, buf)
    return buf.value

MY_VIDEO_DIR = getMyVideoDir()
PAGI_DIR = Path(__file__).resolve().parent.joinpath('pyautogui_images')
TMP_WAV_PATH = 'tmp.wav'

def convertMp4ToWav(src, dst):
    video = moviepy.editor.VideoFileClip(src)
    video.audio.write_audiofile(dst)

def removeSilentFromData(data):
    i = 0
    step = 160
    scnt = 0
    min_threshold = int(np.min(data) * 0.01)
    max_threshold = int(np.max(data) * 0.01)
    retbuf = np.empty(shape=[0, 2], dtype=np.int16)

    sounds = []
    start_index = None
    end_index = None

    for i in range(0, len(data), step):
        window = data[i:i+step]
        # print(window.shape[1])
        # print('i:', i, i+step)
        tag = np.min(window) <= min_threshold or np.max(window) >= max_threshold
        index = len(sounds)
        sounds.append(window)
        if tag:
            if not start_index:
                start_index = index
            end_index = index

    # 掐头去尾
    for i in range(0, len(sounds)):
        if start_index <= i + 5 and i - 5 <= end_index:
            window = sounds[i]
            retbuf = np.append(retbuf, window, axis=0)

    return retbuf

def removeSilent(src_wav, dst_wav):
    fs, data = wavfile.read(src_wav)
    retbuf = removeSilentFromData(data)
    wavfile.write(dst_wav, fs, retbuf)


def autoAmplifyVolume(src_wav, dst_wav):
    # Read the WAV file
    fs, data = wavfile.read(src_wav)

    # Check if the data is in a format that can be amplified
    if data.dtype != 'int16':
        raise ValueError("The WAV file does not contain 16-bit PCM samples.")

    # Compute the maximum possible amplitude for int16 data
    max_amplitude = 32767

    # Find the maximum absolute value of the audio data
    max_abs_value = np.max(np.abs(data))

    # If the maximum absolute value is already at or exceeds the limit, no amplification is needed
    if max_abs_value >= max_amplitude:
        print("Audio is already at maximum volume or has clipping. No amplification applied.")
        wavfile.write(dst_wav, fs, data)
        return

    # Calculate the amplification factor to normalize the audio to the maximum amplitude
    if max_abs_value == 0:
        print("Audio contains only silence. Cannot amplify.")
        return
    amplification_factor = max_amplitude / max_abs_value

    # Apply the amplification factor
    data = (data.astype(float) * amplification_factor).astype(int)

    # Clip the data to prevent overflow when converting back to int16
    data = np.clip(data, -max_amplitude, max_amplitude)

    # Convert back to int16 and write to the destination WAV file
    data = data.astype(np.int16)
    wavfile.write(dst_wav, fs, data)


def convertWavToMp3(src, dst):
    audio = moviepy.editor.AudioFileClip(src)
    audio.write_audiofile(dst)

def findAnkiWindow():
    data = {}

    def hwndFunc(hwnd, mouse):
        if not win32gui.IsWindow(hwnd):
            return
        if not win32gui.IsWindowEnabled(hwnd):
            return
        if not win32gui.IsWindowVisible(hwnd):
            return
        classname = win32gui.GetClassName(hwnd)
        if classname != 'Qt5QWindowIcon':
            return

        title = win32gui.GetWindowText(hwnd)
        if title.endswith('- Anki'):
            data['HWND'] = hwnd

    win32gui.EnumWindows(hwndFunc, 0)
    return data.get('HWND', None)

def findGoogleTranslateWindow():
    data = {}

    def hwndFunc(hwnd, mouse):
        if not win32gui.IsWindow(hwnd):
            return
        if not win32gui.IsWindowEnabled(hwnd):
            return
        if not win32gui.IsWindowVisible(hwnd):
            return
        classname = win32gui.GetClassName(hwnd)
        if not classname.startswith('Chrome_WidgetWin_'):
            return

        title = win32gui.GetWindowText(hwnd)
        if title.startswith('Google Translate -'):
            data['HWND'] = hwnd

    win32gui.EnumWindows(hwndFunc, 0)
    return data.get('HWND', None)

def findOBSWindow():
    data = {}

    def hwndFunc(hwnd, mouse):
        if not win32gui.IsWindow(hwnd):
            return
        if not win32gui.IsWindowEnabled(hwnd):
            return
        if not win32gui.IsWindowVisible(hwnd):
            return
        classname = win32gui.GetClassName(hwnd)
        if not classname.startswith('Qt'):
            return

        title = win32gui.GetWindowText(hwnd)
        if title.startswith('OBS '):
            data['HWND'] = hwnd

    win32gui.EnumWindows(hwndFunc, 0)
    return data.get('HWND', None)

def findLocationByPng(png_name):
    png = os.path.join(PAGI_DIR, png_name+'.png')
    location = pyautogui.locateCenterOnScreen(png, confidence=0.9)
    return location

def pressSeq(*vks):
    for vk in vks:
        win32api.keybd_event(vk, 0, 0, 0)
    for vk in reversed(vks):
        win32api.keybd_event(vk, 0, win32con.KEYEVENTF_KEYUP, 0)

def pressCtrlAnd(vk):
    pressSeq(0x11, vk) # 0x11: ctrl

def pressCtrlA():
    pressCtrlAnd(0x41) # 0x41: A

def pressCtrlC():
    pressCtrlAnd(0x43) # 0x43: C

def pressCtrlV():
    pressCtrlAnd(0x56) # 0x56: V

def pressCtrlShift0():
    pressSeq(0x11, 0x10, 0x30) # 0x56: V, 0x30: 0 key

def getTextFromClipboard():
    win32clipboard.OpenClipboard(0)
    msg = win32clipboard.GetClipboardData()
    win32clipboard.CloseClipboard()
    return msg

def setTextIntoClipboard(msg):
    # 设置剪贴板文本
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT, msg)
    win32clipboard.CloseClipboard()

def copyTextFromAnki():
    hwnd = findAnkiWindow()
    win32gui.ShowWindow(hwnd, 5)  # 5: SM_SHOW
    win32gui.SetForegroundWindow(hwnd)

    location = findLocationByPng('ANKI_TOP_BAR')
    offset = 80
    pyautogui.click(location.x, location.y + offset)
    pressCtrlA()
    pressCtrlC()
    time.sleep(0.2)

def startRecording():
    hwnd = findOBSWindow()
    assert hwnd, '找不到 OBS 窗口'
    win32gui.ShowWindow(hwnd, 5)  # 5: SM_SHOW
    win32gui.SetForegroundWindow(hwnd)

    location = findLocationByPng('OBS_START_RECORDING')
    assert location
    pyautogui.click(location.x, location.y)
    return location

def stopRecording(location):
    pyautogui.click(location.x, location.y)

def pasteTextIntoGoogleTranslate():
    hwnd = findGoogleTranslateWindow()
    assert hwnd, '找不到 Google Translate 窗口'
    win32gui.ShowWindow(hwnd, 5)  # 5: SM_SHOW
    win32gui.SetForegroundWindow(hwnd)

    location = findLocationByPng('GOOGLE_TRANSLATE_INPUT')
    pyautogui.click(location.x + 30, location.y + 100)
    pressCtrlA()
    pressCtrlV()
    time.sleep(0.2)

def copyTextFromGoogleTranslate():
    hwnd = findGoogleTranslateWindow()
    assert hwnd, '找不到 Google Translate 窗口'
    win32gui.ShowWindow(hwnd, 5)  # 5: SM_SHOW
    win32gui.SetForegroundWindow(hwnd)

    location = findLocationByPng('GOOGLE_TRANSLATE_INPUT')
    pyautogui.click(location.x + 30, location.y + 100)
    pressCtrlA()
    pressCtrlC()
    time.sleep(0.2)

def recordVoiceToMp4():
    hwnd = findGoogleTranslateWindow()
    assert hwnd, '找不到 Google Translate 窗口'
    win32gui.ShowWindow(hwnd, 5)  # 5: SM_SHOW
    win32gui.SetForegroundWindow(hwnd)

    location = findLocationByPng('GOOGLE_TRANSLATE_PLAY')
    assert location

    obs_location = startRecording()
    pyautogui.click(location.x + 20, location.y - 10)
    time.sleep(5)
    stopRecording(obs_location)

def copyTextFromMp4():
    path = [x for x in os.listdir(MY_VIDEO_DIR) if x.lower().endswith('.mp4')][-1]
    print('copyTextFromMp4:', path)
    msg = path[:-4]
    setTextIntoClipboard(msg)

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--msg-in-gt', action='store_true')
    parser.add_argument('--msg-in-mp4', action='store_true')
    parser.add_argument('--msg-in-anki', action='store_true')
    args = parser.parse_args()

    if args.msg_in_mp4:
        copyTextFromMp4()
    elif args.msg_in_gt:
        copyTextFromGoogleTranslate()
    elif args.msg_in_anki:
        copyTextFromAnki()
    else:
        # 默认使用剪贴板中的文本
        pass

    msg = getTextFromClipboard().strip()
    msg_fix = msg.replace(':', '.').replace('?', '.').replace('\r\n', ' ').replace('\n', ' ').replace('\r', '')
    msg_dot = msg_fix.endswith('.') and msg_fix or msg_fix + '.'

    # if not args.msg_in_mp4:
    #     recordVoiceToMp4()

    path = [x for x in os.listdir(MY_VIDEO_DIR) if x.lower().endswith('.mp4')][-1]
    print('path:', path)
    path = os.path.join(MY_VIDEO_DIR, path)
    mp3_path = os.path.join(MY_VIDEO_DIR, f'{msg_dot}mp3')
    print('mp3_path:', mp3_path)

    time.sleep(1)
    print('isfile:', os.path.isfile(path), 'path:', path)
    convertMp4ToWav(path, TMP_WAV_PATH)
    removeSilent(TMP_WAV_PATH, TMP_WAV_PATH)
    autoAmplifyVolume(TMP_WAV_PATH, TMP_WAV_PATH)
    convertWavToMp3(TMP_WAV_PATH, mp3_path)
    os.remove(TMP_WAV_PATH)

if __name__ == '__main__':
    main()
