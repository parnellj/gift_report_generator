import pyautogui
from pyautogui import *
import os
import glob
import webbrowser
import time
import datetime
from cls_stewardship.cls_stewardship import letter_gen
from Tests import email_sender

print os.path.realpath(__file__)

SHORT_PAUSE = 0.5
LONG_PAUSE = 2.0
LOAD_PAUSE = 5.0
pyautogui.PAUSE = LONG_PAUSE

ROOT_DIR = 'D:'
ICON_DIR = os.path.join('.', 'resources', 'icons')

# These are all local directories specific to my filesystem; change as needed
DOWNLOAD_DIR = os.sep.join([    ROOT_DIR, 'Downloads'])
REPORT_DIR = os.sep.join([      ROOT_DIR, 'Box Sync', 'Misc Projects', '_Gift Reports'])
LETTER_DIR = os.sep.join([      ROOT_DIR, 'Box Sync', 'Office Writing', 'Letters'])

def click_match(img, delay=0):
    """
    Find and click on the center of a specified image, adding an optional delay.
    :param img: The image to click.
    :param delay: Delay in addition to pyautogui.PAUSE
    :return: None
    """
    click(locateCenterOnScreen(img))
    time.sleep(delay)


def get_date_range():
    last_saturday = datetime.datetime.today()
    while last_saturday.isoweekday() != 6: last_saturday = last_saturday - datetime.timedelta(1)
    last_sunday = last_saturday - datetime.timedelta(6)
    return last_sunday, last_saturday


def outlook_email(account, recipients, subject, body, attachment=None, live=False, initial=True):
    if initial:
        os.system('start outlook.exe')
    time.sleep(LOAD_PAUSE)
    hotkey('ctrl', 'n')
    time.sleep(LONG_PAUSE)

    hotkey('alt', 'm')
    if account == 'MM':
        press(['down'] * 2, interval=SHORT_PAUSE)
    elif account == 'COS':
        press(['down'] * 1, interval=SHORT_PAUSE)

    press('enter')
    if isinstance(recipients, str):
        recipients = [recipients]
    for r in recipients:
        typewrite(r + ', ')
    hotkey('alt', 'u')
    typewrite(subject)
    press('tab')

    press(['up'] * 3, interval=SHORT_PAUSE)
    if isinstance(body, str):
        body = [body]
    for b in body:
        if b == 'PASTE':
            hotkey('ctrl', 'v')
        else:
            typewrite(b)

    if attachment is not None:
        press(['alt', 'n', 'a', 'f', 'b'], interval=SHORT_PAUSE)
        typewrite(attachment)
        press('enter')

    if live:
        hotkey('ctrl', 'enter')
        press('space')
        hotkey('alt', 'f4')


def download_report(start=None, end=None):
    pp = make_dates(start, end)

    webbrowser.open_new('http://google.com')
    time.sleep(LONG_PAUSE)
    hotkey('ctrl', 't')
    typewrite('advance.utah.edu')
    press('enter')
    time.sleep(LONG_PAUSE)
    press(['tab'] * 12, interval=SHORT_PAUSE)
    press('enter', interval=LONG_PAUSE)
    press('enter', interval=LONG_PAUSE)
    press(['tab', 'enter'], interval=LONG_PAUSE)
    press(['tab'] * 6, interval=SHORT_PAUSE)
    press('enter')

    time.sleep(LONG_PAUSE)
    hotkey('ctrl', 'f')
    typewrite('stewardship')
    hotkey('escape')
    hotkey('shift', 'tab')
    hotkey('enter')
    typewrite(pp['start'])
    press('tab')
    typewrite(pp['end'])
    press('enter')
    moveTo(x=630, y=230)
    moveTo(x=640, y=240)
    click()

    csv_file = max(glob.glob(os.path.join(DOWNLOAD_DIR, '*')), key=os.path.getctime)
    new_csv_path = os.path.join(REPORT_DIR, 'gift_report_{0}_{1}.csv')
    new_csv_path = new_csv_path.format(pp['start_ymd'], pp['end_ymd'])
    print csv_file
    print new_csv_path
    if not os.path.isfile(new_csv_path):
        os.rename(csv_file, new_csv_path)
    return new_csv_path


def make_pretty_report(csv_path):
    new_report_path = csv_path.replace('csv', 'xlsx')
    print new_report_path
    os.system('start excel.exe "%s"' % csv_path)
    time.sleep(LOAD_PAUSE)
    hotkey('ctrl', 'a')
    press(['alt', 'h', 't', 'enter', 'enter', 'f12'], interval=SHORT_PAUSE)
    hotkey('alt', 't')
    press('e')
    press('e')
    press('e')
    hotkey('alt', 'n')
    typewrite(new_report_path)
    press('enter')
    hotkey('alt', 'f4')

    return new_report_path


def email_pretty_report(rpt_path, recipients=None):
    account = 'COS'
    if not recipients:
        recipients = ['default@mail.com']
    subject = 'College of Science gift report'
    body = 'Hello All,\n\nHere is the report for gifts made to the College of Science last ' \
           'week.  Please let me know if you have any questions.\n\nIf a gift was made to your area during this ' \
           'period, you will be receiving this week\'s thank you letters for review shortly.\n\nThanks!'
    attachment = rpt_path
    outlook_email(account, recipients, subject, body, attachment)
    pass

def make_dates(start, end):
    if start is None or end is None:
        end = datetime.datetime.today()
        start = end - datetime.timedelta(days=6)
    dates = {'start': start.strftime('%m/%d/%Y'),
             'end': end.strftime('%m/%d/%Y'),
             'start_ymd': start.strftime('%Y%m%d'),
             'end_ymd': end.strftime('%Y%m%d'),
             'today': datetime.datetime.today().strftime('%Y-%m-%d')}

    return dates

if __name__ == '__main__':
    dr = (datetime.datetime(2017, 06, 18), datetime.datetime(2017, 07, 01))
    dates = make_dates(*dr)
    csv_path_output = download_report(dates)
    rpt_path = make_pretty_report(dates)
    letter_gen.create_materials(csv_path=csv_path_output, dates=dates)
