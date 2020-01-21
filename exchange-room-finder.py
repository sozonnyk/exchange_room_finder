#!/usr/bin/env python

from tqdm import tqdm
from datetime import datetime
from exchangelib import *
from exchangelib.protocol import NoVerifyHTTPAdapter, BaseProtocol
from exchangelib.items import SEND_TO_ALL_AND_SAVE_COPY
import urllib3
import re
import os
import logging
import yaml
import sys
from getpass import getpass
from calendar import *
from dateutil.relativedelta import *
import signal

DEFAULT_MEETING_DURATION = 30
DEFAULT_INCREMENT_MINUTES = 30

signal.signal(signal.SIGINT, lambda x,y: exit(0))
logging.basicConfig(level=logging.FATAL)
BaseProtocol.HTTP_ADAPTER_CLS = NoVerifyHTTPAdapter
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

config = yaml.safe_load(open('exchange-room-finder.yml', 'r'))
user = config['primary_email']
password = getpass(prompt="Password:")
server = config['server']
rooms_prefixes = config['rooms_prefixes']


def unset_proxy():
    for var in [v for v in os.environ.keys() if re.search(r"proxy", v, flags=re.IGNORECASE)]:
        del os.environ[var]

def matching_room(name, include_informal=False, include_standing=False):
    m = re.match(r".*\((\d+)\)",name)

    if not m:
        return False

    size = int(m.group(1))
    tags = re.findall(r"\((\D+?)\)",name)

    if size < 3:
        return False

    if 'Informal' in tags and not include_informal:
        return False

    if ('Stand-Up' in tags or 'Standing Only' in tags) and not include_standing:
        return False

    return True

def round_minutes(dt, direction, resolution):
    new_minute = (dt.minute // resolution + (1 if direction == 'up' else 0)) * resolution
    return dt + datetime.timedelta(minutes=new_minute - dt.minute)

def no_overlap(A_start, A_end, B_start, B_end):
    latest_start = max(A_start, B_start)
    earliest_end = min(A_end, B_end)
    return latest_start >= earliest_end

def ask(question, default, to_lower = True):

    if default:
        default_str = f" [{default}]"
    else:
        default_str = ""

    default = str(default)
    response = input(f"{question}{default_str}: ") or default
    if to_lower:
        response = response.lower()
    return response

def colorize(string, color_code):
    return "\033[{}m{}\033[0m".format(color_code, string)


def red(string):
    return colorize(string, 31)


def green(string):
    return colorize(string, 32)


def yellow(string):
    return colorize(string, 33)


def magenta(string):
    return colorize(string, 35)


def white(string):
    return colorize(string, 97)

class NiceCalendar(TextCalendar):
    current_month = 0
    def prmonth(self, theyear, themonth, w=0, l=0):
        """
        Print a month's calendar.
        """
        self.current_month = themonth
        print(self.formatmonth(theyear, themonth, w, l), end='')

    def formatday(self, day, weekday, width):
        """
        Returns a formatted day.
        """
        if day == 0:
            s = ''
        else:
            s = '%2i' % day             # right-align single-digit days
        if weekday in [5,6]:
            s = red(s)
        if day == datetime.now().day and self.current_month == datetime.now().month:
            s = green(s)
        return s.center(width)

#unset_proxy()
print("Logging in to Exchange server.")
credentials = Credentials(user, password)
config = Configuration(server=server, credentials=credentials)
a = Account(primary_smtp_address=user, config=config, autodiscover=False, access_type=DELEGATE)

informal = (ask("Include informal rooms ?", "N") == "y")

# Load all rooms
all_rooms = []
for room_prefix in tqdm(iterable=rooms_prefixes, desc="Loading rooms for each floor", ncols=100, unit="floor"):
    for mailbox, contact in a.protocol.resolve_names([room_prefix], return_full_contact_data=True):
        if matching_room(contact.display_name, include_standing=True, include_informal=informal):
            all_rooms.append((mailbox.email_address, contact.display_name))
print()

tz = EWSTimeZone.localzone()
today = EWSDateTime.now()

date_str = ask(f"Meeting today {white(datetime.strftime(today, '%d/%m/%Y'))} ?", "Y")
if date_str != "y":
    print()
    NiceCalendar().prmonth(today.year, today.month)
    print()
    same_month = ask(f"Meeting this month ?", "Y")
    month = today.month
    if same_month != "y":
        next_month = datetime.today() + relativedelta(months=+1)
        print()
        NiceCalendar().prmonth(next_month.year, next_month.month)
        print()
        month = next_month.month

    day = ask(f"Meeting day ?", today.day)
    today = today.replace(day=int(day), month=int(month))

start_of_business = tz.localize(today.replace(hour=6, minute=0, second=0, microsecond=0))
end_of_business = tz.localize(today.replace(hour=21, minute=0, second=0, microsecond=0))

print(f"Looking for rooms on {white(datetime.strftime(today, '%d/%m/%Y'))}")

rooms_data = []
for room_email, room_name in tqdm(iterable=all_rooms, desc="Loading rooms availability", ncols=100, unit="room"):
    busy_data = []
    for free_busy_info in a.protocol.get_free_busy_info(accounts=[(Account(primary_smtp_address=room_email,
                                                                     config=config, autodiscover=False,
                                                                     access_type=DELEGATE), 'Required', False)],
                                                  start=start_of_business,
                                                  end=end_of_business):
        for event in (free_busy_info.calendar_events or []):
            busy_data.append((event.start.astimezone(tz), event.end.astimezone(tz)))

    rooms_data.append({"email": room_email, "name": room_name, "busy": busy_data})

print()

def generate_ews_time(hour, minute):
    return tz.localize(EWSDateTime(today.year, month=today.month, day=today.day, hour=hour, minute=minute, second=0, microsecond=0))

def get_time(prefix, default=None):

    if default:
        default_str = f" [{default}]"
    else:
        default_str = ""

    valid_date = False
    result = {}
    while not valid_date:
        try:
            time_str = input(f"{prefix} HH24[:MM]-HH24[:MM] {default_str or ''} ? ") or default
            m = re.match(r"(\d+):*(\d*)\s*-\s*(\d+):*(\d*)", time_str)
            result['start'] = generate_ews_time(int(m.group(1)), int(m.group(2) or 0))
            result['end'] = generate_ews_time(int(m.group(3)), int(m.group(4) or 0))
            valid_date = True
        except Exception as e:
            print(f"Bad date {time_str}")
            #print(e)

    return result

room_found = False
while not room_found:
    free_rooms = []

    preference = get_time("Preferred time", "9:30-16:30")
    duration = int(ask(f"Meeting duration in minutes ?", DEFAULT_MEETING_DURATION))
    for room in rooms_data:
        room['available'] = []
        slot_start = None
        slot_end = None

        while True:
            slot_start = slot_start or preference["start"]
            slot_end = tz.localize(EWSDateTime.fromtimestamp(slot_start.timestamp() + duration * 60))

            if slot_end.timestamp() > preference["end"].timestamp():
                break

            if all(no_overlap(slot_start, slot_end, event[0], event[1]) for event in room["busy"]):
                room['available'].append((slot_start,slot_end))

            slot_start = tz.localize(EWSDateTime.fromtimestamp(slot_start.timestamp() + DEFAULT_INCREMENT_MINUTES * 60))

        if any(room['available']):
            free_rooms.append(room)
            room_found = True

    if not room_found:
        print(yellow(f"No free rooms in this time range. Please specify different time."))

def format_range(slot_start, slot_end):
    return f"{datetime.strftime(slot_start, '%H:%M')}-{datetime.strftime(slot_end, '%H:%M')}"

def slots_string(room, num):
    result = ""
    avail_num = len(room['available'])
    if avail_num > 0:
        slots = []
        for i in range(0, min(num, avail_num)):
            slot_start, slot_end = room['available'][i]
            slots.append(format_range(slot_start, slot_end))

        result = ", ".join(slots)
        result = f" [{result}]"
        if avail_num > num:
            result += f" and {avail_num-num} more"

    return result

print()
print("Slots available:")

for idx, room in enumerate(free_rooms):
    print(f"{idx}: {room['name']}{slots_string(room, 2)}")

room_selected = None
while not room_selected:
    try:
        room_idx_str = input("Choose a room ? ")
        room_selected = free_rooms[int(room_idx_str)]
    except Exception as e:
        print(f"Bad index {room_idx_str}")
        #print(e)

slot_selected = None

available = len(room_selected['available'])
if available == 1:
    start, end = room_selected['available'][0]

if available > 1:
    print(f"Choose a time slot for {white(room_selected['name'])}:")
    for idx, slot in enumerate(room_selected['available']):
        print(f"{idx}: {format_range(*slot)}")
    while not slot_selected:
        try:
            slot_idx_str = input("Choose a slot ? ")
            slot_selected = room_selected['available'][int(slot_idx_str)]
            start, end = slot_selected
        except Exception as e:
            print(f"Bad index {slot_idx_str}")
            print(e)

subj = ask("Meeting subject ?", "New Appointment", to_lower = False)

# create a meeting request and send it out
item = CalendarItem(
    account=a,
    folder=a.calendar,
    start=start,
    end=end,
    subject=subj,
    body="",
    required_attendees=[room_selected["email"]]
)

item.save(send_meeting_invitations=SEND_TO_ALL_AND_SAVE_COPY)

print(green("All done"))

