# -*- coding: utf-8 -*-
"""
Created on Tue Sep 21 2021

@author: Hagai Hamami

Information regarding Microsoft Windows codes could be found online:
    Virtual-key codes:
        https://docs.microsoft.com/en-us/windows/win32/inputdev/virtual-key-codes
    Mouse events:
        https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-mouse_event
"""
# Start
import ctypes
from time import sleep
import subprocess

OpenClipboard = ctypes.windll.user32.OpenClipboard
EmptyClipboard = ctypes.windll.user32.EmptyClipboard
GetClipboardData = ctypes.windll.user32.GetClipboardData
SetClipboardData = ctypes.windll.user32.SetClipboardData
CloseClipboard = ctypes.windll.user32.CloseClipboard
CF_UNICODETEXT = 13

GlobalAlloc = ctypes.windll.kernel32.GlobalAlloc
GlobalLock = ctypes.windll.kernel32.GlobalLock
GlobalUnlock = ctypes.windll.kernel32.GlobalUnlock
GlobalSize = ctypes.windll.kernel32.GlobalSize
GMEM_MOVEABLE = 0x0002
GMEM_ZEROINIT = 0x0040

unicode_type = type(u'')
u32 = ctypes.windll.user32

# Mouse
def pointer():
    # Prints current mouse screen position (x,y coordinates)
    from ctypes import windll, Structure, c_long, byref
    class POINT(Structure):
        _fields_ = [("x", c_long), ("y", c_long)]
    def queryMousePosition():
        pt = POINT()
        windll.user32.GetCursorPos(byref(pt))
        return (pt.x, pt.y)
        #return { "x": pt.x, "y": pt.y}
    pos = queryMousePosition()
    print(pos)
    
# Mouse clickers
def R(x,y):
    # Right click
    u32.SetCursorPos(x, y)
    u32.mouse_event(0x0008, 0, 0, 0,0) # right mouse click down
    u32.mouse_event(0x0010, 0, 0, 0,0) # right mouse click up
    sleep(2)
def L(x,y):
    # Left click
    u32.SetCursorPos(x, y)
    u32.mouse_event(2, 0, 0, 0,0) # left mouse click down
    u32.mouse_event(4, 0, 0, 0,0) # left mouse click up
    sleep(0.7)
def D(x,y):
    # Double (left) click
    u32.SetCursorPos(x, y)
    sleep(1)
    u32.mouse_event(2, 0, 0, 0,0) # left mouse click down
    u32.mouse_event(4, 0, 0, 0,0) # left mouse click up
    u32.mouse_event(2, 0, 0, 0,0) # left mouse click down
    u32.mouse_event(4, 0, 0, 0,0) # left mouse click up
    sleep(2)

# Keyboard - keys

# Keyboard - functionalities

def copy2clip(txt):
    # Copy text from STRING object
    cmd='echo '+txt.strip()+'|clip'
    return subprocess.check_call(cmd, shell=True)

def copy_ctrl_c():
    u32.keybd_event(0x11, 0, 0, 0)  # ctrl down
    u32.keybd_event(0x43, 0, 0, 0)  # c down
    u32.keybd_event(0x11, 0, 2, 0)  # ctrl up
    u32.keybd_event(0x43, 0, 2, 0)  # c up

def paste_ctrl_v():
    u32.keybd_event(0x11, 0, 0, 0)  # ctrl down
    u32.keybd_event(0x56, 0, 0, 0)  # v down
    u32.keybd_event(0x11, 0, 2, 0)  # ctrl up
    u32.keybd_event(0x56, 0, 2, 0)  # v up

def paste_from_string(x):
    # Copy text stored in STRING object x and use keyboard to paste
    copy2clip(x)
    paste_ctrl_v()


def Q():
    # Quit the script gracefully 
    print("""\n### END OF SCRIPT ###""")
    raise SystemExit()


# The STRING object VK_STR holds information regarding microsoft windows virtual-key codes, which could be found here: https://docs.microsoft.com/en-us/windows/win32/inputdev/virtual-key-codes
VK_STR = """VK_LBUTTON
0x01
Left mouse button
VK_RBUTTON
0x02
Right mouse button
VK_CANCEL
0x03
Control-break processing
VK_MBUTTON
0x04
Middle mouse button (three-button mouse)
VK_XBUTTON1
0x05
X1 mouse button
VK_XBUTTON2
0x06
X2 mouse button
-
0x07
Undefined
VK_BACK
0x08
BACKSPACE key
VK_TAB
0x09
TAB key
-
0x0A-0B
Reserved
VK_CLEAR
0x0C
CLEAR key
VK_RETURN
0x0D
ENTER key
-
0x0E-0F
Undefined
VK_SHIFT
0x10
SHIFT key
VK_CONTROL
0x11
CTRL key
VK_MENU
0x12
ALT key
VK_PAUSE
0x13
PAUSE key
VK_CAPITAL
0x14
CAPS LOCK key
VK_KANA
0x15
IME Kana mode
VK_HANGUEL
0x15
IME Hanguel mode (maintained for compatibility; use VK_HANGUL)
VK_HANGUL
0x15
IME Hangul mode
-
0x16
Undefined
VK_JUNJA
0x17
IME Junja mode
VK_FINAL
0x18
IME final mode
VK_HANJA
0x19
IME Hanja mode
VK_KANJI
0x19
IME Kanji mode
-
0x1A
Undefined
VK_ESCAPE
0x1B
ESC key
VK_CONVERT
0x1C
IME convert
VK_NONCONVERT
0x1D
IME nonconvert
VK_ACCEPT
0x1E
IME accept
VK_MODECHANGE
0x1F
IME mode change request
VK_SPACE
0x20
SPACEBAR
VK_PRIOR
0x21
PAGE UP key
VK_NEXT
0x22
PAGE DOWN key
VK_END
0x23
END key
VK_HOME
0x24
HOME key
VK_LEFT
0x25
LEFT ARROW key
VK_UP
0x26
UP ARROW key
VK_RIGHT
0x27
RIGHT ARROW key
VK_DOWN
0x28
DOWN ARROW key
VK_SELECT
0x29
SELECT key
VK_PRINT
0x2A
PRINT key
VK_EXECUTE
0x2B
EXECUTE key
VK_SNAPSHOT
0x2C
PRINT SCREEN key
VK_INSERT
0x2D
INS key
VK_DELETE
0x2E
DEL key
VK_HELP
0x2F
HELP key
VK_LWIN
0x5B
Left Windows key (Natural keyboard) 
VK_RWIN
0x5C
Right Windows key (Natural keyboard)
VK_APPS
0x5D
Applications key (Natural keyboard)
-
0x5E
Reserved
VK_SLEEP
0x5F
Computer Sleep key
VK_NUMPAD0
0x60
Numeric keypad 0 key
VK_NUMPAD1
0x61
Numeric keypad 1 key
VK_NUMPAD2
0x62
Numeric keypad 2 key
VK_NUMPAD3
0x63
Numeric keypad 3 key
VK_NUMPAD4
0x64
Numeric keypad 4 key
VK_NUMPAD5
0x65
Numeric keypad 5 key
VK_NUMPAD6
0x66
Numeric keypad 6 key
VK_NUMPAD7
0x67
Numeric keypad 7 key
VK_NUMPAD8
0x68
Numeric keypad 8 key
VK_NUMPAD9
0x69
Numeric keypad 9 key
VK_MULTIPLY
0x6A
Multiply key
VK_ADD
0x6B
Add key
VK_SEPARATOR
0x6C
Separator key
VK_SUBTRACT
0x6D
Subtract key
VK_DECIMAL
0x6E
Decimal key
VK_DIVIDE
0x6F
Divide key
VK_F1
0x70
F1 key
VK_F2
0x71
F2 key
VK_F3
0x72
F3 key
VK_F4
0x73
F4 key
VK_F5
0x74
F5 key
VK_F6
0x75
F6 key
VK_F7
0x76
F7 key
VK_F8
0x77
F8 key
VK_F9
0x78
F9 key
VK_F10
0x79
F10 key
VK_F11
0x7A
F11 key
VK_F12
0x7B
F12 key
VK_F13
0x7C
F13 key
VK_F14
0x7D
F14 key
VK_F15
0x7E
F15 key
VK_F16
0x7F
F16 key
VK_F17
0x80
F17 key
VK_F18
0x81
F18 key
VK_F19
0x82
F19 key
VK_F20
0x83
F20 key
VK_F21
0x84
F21 key
VK_F22
0x85
F22 key
VK_F23
0x86
F23 key
VK_F24
0x87
F24 key
-
0x88-8F
Unassigned
VK_NUMLOCK
0x90
NUM LOCK key
VK_SCROLL
0x91
SCROLL LOCK key
VK_OEM_specific
0x92-96
OEM specific
-
0x97-9F
Unassigned
VK_LSHIFT
0xA0
Left SHIFT key
VK_RSHIFT
0xA1
Right SHIFT key
VK_LCONTROL
0xA2
Left CONTROL key
VK_RCONTROL
0xA3
Right CONTROL key
VK_LMENU
0xA4
Left MENU key
VK_RMENU
0xA5
Right MENU key
VK_BROWSER_BACK
0xA6
Browser Back key
VK_BROWSER_FORWARD
0xA7
Browser Forward key
VK_BROWSER_REFRESH
0xA8
Browser Refresh key
VK_BROWSER_STOP
0xA9
Browser Stop key
VK_BROWSER_SEARCH
0xAA
Browser Search key 
VK_BROWSER_FAVORITES
0xAB
Browser Favorites key
VK_BROWSER_HOME
0xAC
Browser Start and Home key
VK_VOLUME_MUTE
0xAD
Volume Mute key
VK_VOLUME_DOWN
0xAE
Volume Down key
VK_VOLUME_UP
0xAF
Volume Up key
VK_MEDIA_NEXT_TRACK
0xB0
Next Track key
VK_MEDIA_PREV_TRACK
0xB1
Previous Track key
VK_MEDIA_STOP
0xB2
Stop Media key
VK_MEDIA_PLAY_PAUSE
0xB3
Play/Pause Media key
VK_LAUNCH_MAIL
0xB4
Start Mail key
VK_LAUNCH_MEDIA_SELECT
0xB5
Select Media key
VK_LAUNCH_APP1
0xB6
Start Application 1 key
VK_LAUNCH_APP2
0xB7
Start Application 2 key
-
0xB8-B9
Reserved
VK_OEM_1
0xBA
Used for miscellaneous characters; it can vary by keyboard. For the US standard keyboard, the ';:' key 
VK_OEM_PLUS
0xBB
For any country/region, the '+' key
VK_OEM_COMMA
0xBC
For any country/region, the ',' key
VK_OEM_MINUS
0xBD
For any country/region, the '-' key
VK_OEM_PERIOD
0xBE
For any country/region, the '.' key
VK_OEM_2
0xBF
Used for miscellaneous characters; it can vary by keyboard. For the US standard keyboard, the '/?' key 
VK_OEM_3
0xC0
Used for miscellaneous characters; it can vary by keyboard. For the US standard keyboard, the '`~' key 
-
0xC1-D7
Reserved
-
0xD8-DA
Unassigned
VK_OEM_4
0xDB
Used for miscellaneous characters; it can vary by keyboard. For the US standard keyboard, the '[{' key
VK_OEM_5
0xDC
Used for miscellaneous characters; it can vary by keyboard. For the US standard keyboard, the '\|' key
VK_OEM_6
0xDD
Used for miscellaneous characters; it can vary by keyboard. For the US standard keyboard, the ']}' key
VK_OEM_7
0xDE
Used for miscellaneous characters; it can vary by keyboard. For the US standard keyboard, the 'single-quote/double-quote' key
VK_OEM_8
0xDF
Used for miscellaneous characters; it can vary by keyboard.
-
0xE0
Reserved
VK_OEM_specific
0xE1
OEM specific
VK_OEM_102
0xE2
Either the angle bracket key or the backslash key on the RT 102-key keyboard
VK_OEM_specific
0xE3-E4
OEM specific
VK_PROCESSKEY
0xE5
IME PROCESS key
VK_OEM_specific
0xE6
OEM specific
VK_PACKET
0xE7
Used to pass Unicode characters as if they were keystrokes. The VK_PACKET key is the low word of a 32-bit Virtual Key value used for non-keyboard input methods. For more information, see Remark in KEYBDINPUT, SendInput, WM_KEYDOWN, and WM_KEYUP
-
0xE8
Unassigned
VK_OEM_specific
0xE9-F5
OEM specific
VK_ATTN
0xF6
Attn key
VK_CRSEL
0xF7
CrSel key
VK_EXSEL
0xF8
ExSel key
VK_EREOF
0xF9
Erase EOF key
VK_PLAY
0xFA
Play key
VK_ZOOM
0xFB
Zoom key
VK_NONAME
0xFC
Reserved 
VK_PA1
0xFD
PA1 key
VK_OEM_CLEAR
0xFE
Clear key"""

# Processing VK_STR into a dictionary VK_dic - the key is the name of the keyboard button and its value is the corresponding virtual-key code
VK_LIST = VK_STR.split("\n")
VK_LIST_list =[]
n=0
while n<len(VK_LIST):
   VK_LIST_list.append([VK_LIST[n],VK_LIST[n+1],VK_LIST[n+2]])
   n+=3
VK_dic = {}
for l in VK_LIST_list:
    k = l[2].lower()
    if k[-4:] == " key":
        k = k[:-4]
    try:
        VK_dic[k]=int(l[1],16)
    except:
        VK_dic[k]=l[1]

def vk_sign(sign):
    # This is the main virtual-key function (press & release a single keyboard button)
    # The "sign" passed to this function is the windows virtual-key code
    if type(sign) == str:
        try:
            sign = int(sign,16)
        except:
            pass
    u32.keybd_event (sign, u32.MapVirtualKeyA(sign,0), 1, 0) # key down (pressed)
    u32.keybd_event (sign, u32.MapVirtualKeyA(sign,0), 3, 0) # key up (released)

def virtual_key(key_start,n=1):
    # Uses vk_sign() but more liberal on key name, and can take just the begining (understands "eNt" is the button "Enter")
    # The variable "n" is an optional, in case the key should be pressed more than once.
    key_start=key_start.lower()
    loop = 0
    while n > loop:
        key_start = str(key_start)
        matches = []
        for i in VK_dic:
            if i.startswith(key_start):
                matches.append(i)
        if len(matches)==1:
            MATCH = matches[0]
            sign = VK_dic[MATCH]
            vk_sign(sign)
            sleep(0.1)
        else:
            print("No Match found:",matches)
        loop += 1


def shift_key_down():
    u32.keybd_event(0xA1, 0, 0, 0)
    sleep(0.25)
def shift_key_up():
    u32.keybd_event(0xA1, 0, 2, 0)
    sleep(0.25)

def shift_alt():
    shift_key_down()
    u32.keybd_event(0x12, 0, 0, 0) #ALT is down
    sleep(1)
    shift_key_up()
    u32.keybd_event(0x12, 0, 2, 0) #ALT is up


def type_sign(sig):
    u32.keybd_event(sig, 0, 0, 0)
    u32.keybd_event(sig, 0, 2, 0)
def close_window(): # WARNING  - Mouse positional coordinates varies between monitors
    winkey_action("up arrow")
    sleep(5)
    L(1417, 10)
def shift_type_sign(si):
    sleep(0.5)
    shift_key_down()
    sleep(0.5)
    type_sign(si)
    sleep(0.5)
    shift_key_up()


def winkey_action(sign):
    u32.keybd_event(0x5B, 0, 0, 0)
    try:
        virtual_key(sign)
    except:
        type_sign(sign)
    u32.keybd_event(0x5B, 0, 2, 0)

def KB(keys):
    # Takes a STRING object and types it letter by letter, using virtual-key codes.
    for i in range(len(keys)):
        t = 0
        key = keys[i]
        if key in problematic_characters:
            copy2clip(key)
            paste_ctrl_v()
            continue
        elif (key in heb_unique):
            if (Current_Language()=="English"):
                shift_alt()
            key_eq = heb_dic[key]
            sign = key_dic[key_eq]
        elif (key in eng_unique):
            if (Current_Language()=="Hebrew"):
                shift_alt()
            sign = key_dic[key]
        elif (key in bilang):
            if (Current_Language()=="Hebrew"):
                key = bilang[key]
            sign = key_dic[key]
        elif (key in heb_dic) and (Current_Language()=="Hebrew"):
            key_eq = heb_dic[key]
            sign = key_dic[key_eq]
        elif (key.isupper()) or (key in shifted):
            if key.isalpha() and key.isupper():
                key = key.lower()
            elif key in shifted:
                print(key,shifted[key])
                sleep(1)
                key = shifted[key]
            sign = key_dic[key]
            shift_type_sign(sign)
            t = 1
        else:
            sign = key_dic[key]
        if t == 0:
            try:
                type_sign(sign)
            except:
                sign = key_dic[key]
                type_sign(sign)



def Current_Language():
    # Checking current keyboard language and returns "English" or "Hebrew" or "Language not found"
    user32 = ctypes.WinDLL('user32', use_last_error=True)
    curr_window = user32.GetForegroundWindow()
    thread_id = user32.GetWindowThreadProcessId(curr_window, 0)
    # Made up of 0xAAABBBB, AAA = HKL (handle object) & BBBB = language ID
    klid = user32.GetKeyboardLayout(thread_id)
    # Language ID -> low 10 bits, Sub-language ID -> high 6 bits
    # Extract language ID from KLID
    lid = klid & (2**16 - 1)
    # Convert language ID from decimal to hexadecimal
    lid_hex = hex(lid)
    if klid == 67699721:
        return "English"
    elif klid == 67961869:
        return "Hebrew"
    else:
        return "Language not found:",lid_hex,klid

# Mapping Hebrew keyboard to QWERTY English keyboard
heb_dic = {'פ': 'p', 'ם': 'o', 'ן': 'i', 'ו': 'u', 'ט': 'y', 'א': 't', 'ר': 'r', 'ק': 'e', 'ש': 'a', 'ד': 's', 'ג': 'd', 'כ': 'f', 'ע': 'g', 'י': 'h', 'ח': 'j', 'ל': 'k', 'ך': 'l', 'ף': ';', 'ץ': '.', 'ת': ',', 'צ': 'm', 'מ': 'n', 'נ': 'b', 'ה': 'v', 'ב': 'c', 'ס': 'x', 'ז': 'z',".":"/",",":"'"}
heb_unique = {'פ': 'p', 'ם': 'o', 'ן': 'i', 'ו': 'u', 'ט': 'y', 'א': 't', 'ר': 'r', 'ק': 'e', 'ש': 'a', 'ד': 's', 'ג': 'd', 'כ': 'f', 'ע': 'g', 'י': 'h', 'ח': 'j', 'ל': 'k', 'ך': 'l', 'ף': ';', 'ץ': '.', 'ת': ',', 'צ': 'm', 'מ': 'n', 'נ': 'b', 'ה': 'v', 'ב': 'c', 'ס': 'x', 'ז': 'z'}

key_dic ={"a":0x41,"b":0x42,"c":0x43,"d":0x44,"e":0x45,"f":0x46,"g":0x47,"h":0x48,"i":0x49,"j":0x4A,"k":0x4B,"l":0x4C,"m":0x4D,"n":0x4E,"o":0x4F,"p":0x50,"q":0x51,"r":0x52,"s":0x53,"t":0x54,"u":0x55,"v":0x56,"w":0x57,"x":0x58,"y":0x59,"z":0x5A,"9":0x69,"8":0x68,"7":0x67,"6":0x66,"5":0x65,"4":0x64,"3":0x63,"2":0x62,"1":0x61,"0":0x60,"=":187,"+":0x6B," ":0x20,".":0xBE,",":0xBC,"*":0x6A,"'":0xDE,";":0xBA,"/":0xBF,"[":0xDB,"]":0xDD,"\\":0xDC}
eng_unique ={"a":0x41,"b":0x42,"c":0x43,"d":0x44,"e":0x45,"f":0x46,"g":0x47,"h":0x48,"i":0x49,"j":0x4A,"k":0x4B,"l":0x4C,"m":0x4D,"n":0x4E,"o":0x4F,"p":0x50,"q":0x51,"r":0x52,"s":0x53,"t":0x54,"u":0x55,"v":0x56,"w":0x57,"x":0x58,"y":0x59,"z":0x5A}
x=""
bilang = {",":"'",".":"/","'":"w","/":"q"}
shifted ={'"':"'","?":"/",":":";","!":"1","@":"2","#":"3","%":"5","^":"6","&":"7","*":"8","(":"9",")":"0","_":"-","{":"[","}":"]","|":"\\"}
problematic_characters = ["!","@","#","$","%","^","&","*","(",")"]
