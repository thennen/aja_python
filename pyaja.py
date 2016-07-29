'''
Needed a sane way to do sputter deposition, but AJA program is a disaster.
The only way I can see to interface is by simulating mouse clicks.
Here goes nothing...
'''
#Assuming initial values of controls when program starts
# TODO: Read values of controls to ensure they are being set correctly
# TODO: allow concurrent depositions
# TODO: Should write a log
# TODO: Always check for error window
# TODO: build in txt/email
# TODO: Some kind of indication of progress.  text based is fine
# TODO: Toggle of DC switch changes the state of others.  How to deal with it?

import win32api
import win32con
import win32gui
import win32com.client
from time import sleep

# Control classes that keep track of their state

class Button(object):
    def __init__(self, initval, location):
        initval = bool(initval)
        self.val = initval
        self.default = initval
        # should be a 2-tuple pixel location
        self.loc = location

    def toggle(self):
        click(self.loc, n=1)
        self.val = not self.val


class Numeric(object):
    def __init__(self, initval, location):
        self.val = initval
        self.default = initval
        self.loc = location

    def set(self, value):
        change_value(self.loc, value)
        self.val = value

# Control containers

class Power_Supply(dict):
    def __init__(self, x, y, switched=False, sw_state=False):
        # x, y should be the inner top left corner of the control box
        self['percent'] = Numeric(0.00, (x + 45, y + 130))
        self['ramp'] = Numeric(60, (x + 45, y + 156))
        self['onoff'] = Button(False, (x + 26, y + 213))
        self['shutter'] = Button(False, (x + 82, y + 197))
        if switched:
            self['switch'] = Button(sw_state, (x + 16, y + 158))


class Gas(dict):
    def __init__(self, x, y):
        self['onoff'] = Button(False, (x + 71, y + 12))
        self['stpt'] = Numeric(0.0, (x + 48, y + 69))

# Time in seconds to pause between gui actions
DELAY = 0.000

CONTROLS = {'SYSTEM_CONFIG': Button(False, (54, 252)),
            'PRESSURE_POSITION': Numeric(1000, (765, 205)),
            'DC1': Power_Supply(20, 518),
            'DC2': Power_Supply(144, 518),
            'DC3': Power_Supply(269, 518),
            'DC4': Power_Supply(393, 518),
            'DC5A': Power_Supply(517, 518, switched=True, sw_state=True),
            'DC5B': Power_Supply(641, 518, switched=True),
            'DC5C': Power_Supply(765, 518, switched=True),
            'DC5D': Power_Supply(889, 518, switched=True),
            'HEAT': Gas(171, 157),
            'GAS1': Gas(292, 157),
            'GAS2': Gas(383, 157),
            'GAS3': Gas(473, 157)
            }

# Define connections
# Or get them from a file
CONNECTIONS = {'Ta': 'DC1',
               'Pt': 'DC2',
               'Tb': 'DC3',
               'Si': 'DC5A'}

# TODO: Define sputter rate calibration
# Or get it from a file

# Find AJA window or quit
PHASEII = None


def enum_callback(hwnd, *args):
    global PHASEII
    txt = win32gui.GetWindowText(hwnd)
    if txt == 'AJA INTERNATIONAL PHASE II J COMPUTER CONTROL':
        PHASEII = hwnd
win32gui.EnumWindows(enum_callback, None)
if PHASEII is None:
    raise Exception('AJA PHASE II program not found.  Open it!')

# This is for sending keystrokes
shell = win32com.client.Dispatch('WScript.Shell')


def show_PHASEII():
    # Bring PHASEII program to the foreground
    if not win32gui.GetForegroundWindow == PHASEII:
        shell.SendKeys('%')
        win32gui.SetForegroundWindow(PHASEII)
        sleep(.3)


def click((x, y), n=1):
    # Make sure PHASEII is showing and click a pixel
    show_PHASEII()
    win32api.SetCursorPos((x, y))
    for _ in range(n):
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, x, y, 0, 0)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, x, y, 0, 0)
    sleep(DELAY)


def change_value(control_loc, newvalue):
    # Change the value of a numeric control
    click(control_loc, 2)
    shell.SendKeys(str(newvalue))
    shell.SendKeys('~')
    sleep(DELAY)


def get_value(control):
    # Get the value
    # Double click
    # copy?
    pass


##### High level functions #####

def set_temp(temp):
    pass


def bake(hours=8):
    # Bake chamber
    pass


def gas(which=1, sccm=20):
    # Flow gas
    # TODO: Don't allow if shutter is closed?
    which = int(which)
    assert(which in (1, 2, 3))
    whichgas = 'GAS' + str(which)
    # Set the new sccm
    stpt = CONTROLS[whichgas]['stpt']
    stpt.set(sccm)
    # Decide whether to hit on/off button
    onoff = CONTROLS[whichgas]['onoff']
    state = onoff.val
    if sccm == 0 and state:
        # Press off button if gas is on
        onoff.toggle()
    elif sccm > 0 and not state:
        # Press on button if gas is off
        onoff.toggle()


def shutter(position):
    # Change shutter position
    position = int(position)
    assert(0 <= position <= 1000)
    CONTROLS['PRESSURE_POSITION'].set(position)


def light(material=None, power=10):
    ''' Create plasma '''
    # TODO: check that gas is flowing and shutter is partially closed
    # TODO: check not already lit
    # TODO: take power in percent or watts
    # TODO: option to specify gun or power supply instead of material
    # TODO: verify plasma started!
    ps = CONNECTIONS[material]
    psbox = CONTROLS[ps]
    # Make sure switch is on if it's a switched supply
    if hasattr(psbox, 'switch') and not psbox['switch'].val:
        psbox['switch'].toggle()
    psbox['percent'].set(power)
    psbox['ramp'].set(3)
    psbox['onoff'].toggle()


def unlight(material=None):
    ps = CONNECTIONS[material]
    psbox = CONTROLS[ps]
    if psbox['onoff'].val:
        psbox['onoff'].toggle()


def deposit(material, thickness=None, time=None, power=10):
    ''' Deposit a material, given thickness or time'''
    # Check if lit, if not, light and wait a little
    ps = CONNECTIONS[material]
    psbox = CONTROLS[ps]
    # Sanity check shutter not already open
    if psbox['shutter'].val:
        raise Exception('Shutter already open!')
    state = psbox['onoff'].val
    if state is False:
        light(material, power)
        sleep(2)
    psbox['shutter'].toggle()
    sleep(time)
    psbox['shutter'].toggle()


def codeposit():
    # Sputter more than one thing
    # Give times for each to start?
    pass


def standby():
    # Return all settings to standby state
    def set_default(control):
        if type(control) == Button:
            if not control.val == control.default:
                control.toggle()
        elif type(control) == Numeric:
            # No harm in overwriting default value
            control.set(control.default)

    # Loop through controls and subcontrols
    for c in CONTROLS.values():
        if type(c) in (Power_Supply, Gas):
            for sc in c.values():
                set_default(sc)
        else:
            set_default(c)


def test_deposition():
    # Make deposition
    gas(1, 50)
    shutter(600)
    sleep(10)
    light('Ta', 10)
    light('Si', 10)
    # Make sure they lit!
    sleep(5)
    deposit('Ta', power=10, time=4)
    deposit('Si', power=10, time=3)
    unlight('Ta')
    unlight('Si')
    gas(1, 0)
    shutter(1000)
