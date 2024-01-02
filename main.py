import csv
import sys
from icalendar import Calendar, Event
import datetime


def create_event(date, time):
    start_time = datetime.datetime.strptime(f'{date} {time}', '%d/%m/%Y %H:%M')
    end_time = start_time + datetime.timedelta(minutes=15)

    event = {
        'summary': 'English Lesson',
        'location': '',
        'start': {
            'dateTime': start_time,
        },
        'end': {
            'dateTime': end_time,
        },
        'reminders': {
            'useDefault': True
        },
    }

    return event


def save_event_to_cal(event, cal):
    e = Event()
    e.add('summary', event.get('summary'))
    e.add('location', event.get('location'))
    e.add('description', event.get('description'))
    e.add('dtstart', event.get('start').get('dateTime'))
    e.add('dtend', event.get('end').get('dateTime'))

    cal.add_component(e)


def main(file_path, save_path):
    cal = Calendar()
    cal.add('prodid', '-//My Calendar Events//mxm.dk//')
    cal.add('version', '2.0')

    with open(file_path, 'r', encoding='utf-8') as file:
        reader = csv.reader(file)
        next(reader)  # Skip the header row
        for row in reader:
            date = row[1]
            time = row[2]

            event = create_event(date, time)
            save_event_to_cal(event, cal)

    with open(save_path, 'wb') as f:
        f.write(cal.to_ical())


if __name__ == '__main__':
    if len(sys.argv) != 3:
        print('Usage: python main.py csv_file out_path.ics')
        print('You can then import the meetings from the ics file to outlook using "Add Calendar"')
        sys.exit(1)

    file_path = sys.argv[1]
    save_path = sys.argv[2]
    main(file_path, save_path)
