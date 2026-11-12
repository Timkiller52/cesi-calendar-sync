#!/usr/bin/env python3
"""Convert CESI's publicly published Outlook calendar into a subscribed ICS feed."""
import datetime as dt
import json
import os
import time
from pathlib import Path
import urllib.error
import urllib.parse
import urllib.request

BASE = 'https://outlook.cloud.microsoft/owa/published/814669fb1f2f4113ba73d9119a06ce8b@cesi.fr/a4bfb5780b964c958e1cc95e51a50d87437230687839495538'
FOLDER = ('AAMkADgxNDY2OWZiLTFmMmYtNDExMy1iYTczLWQ5MTE5YTA2Y2U4Yg'
          'AuAAAAAAC/pgTeeh+RSrN0AXJ0B9GkAQAJiFqvS9MHTahlJT3vKLgiAAjVNYIPAAA=')
OUTPUT = Path('public/calendar.ics')


def escape(s):
    return str(s or '').replace('\\', '\\\\').replace('\n', '\\n').replace('\r', '').replace(';', '\\;').replace(',', '\\,')


def fold(line):
    # RFC 5545 content lines are limited to 75 octets, not characters.
    chunks = []
    current = ''
    for c in line:
        if len((current + c).encode('utf-8')) > (74 if chunks else 75):
            chunks.append(current)
            current = ' ' + c
        else:
            current += c
    chunks.append(current)
    return '\r\n'.join(chunks)


def parse_time(value):
    if not value:
        raise ValueError('Missing event time')
    value = value.replace('Z', '+00:00')
    t = dt.datetime.fromisoformat(value)
    if t.tzinfo is None:
        from zoneinfo import ZoneInfo
        t = t.replace(tzinfo=ZoneInfo('Europe/Paris'))
    return t.astimezone(dt.timezone.utc).strftime('%Y%m%dT%H%M%SZ')


def fetch(start, end):
    payload = {
        '__type': 'FindItemJsonRequest:#Exchange',
        'Header': {'__type': 'JsonRequestHeaders:#Exchange',
                   'RequestServerVersion': 'V2018_01_08',
                   'TimeZoneContext': {'__type': 'TimeZoneContext:#Exchange',
                                       'TimeZoneDefinition': {'__type': 'TimeZoneDefinitionType:#Exchange', 'Id': 'Romance Standard Time'}}},
        'Body': {'__type': 'FindItemRequest:#Exchange',
                 'ParentFolderIds': [{'__type': 'FolderId:#Exchange', 'Id': FOLDER}],
                 'ItemShape': {'__type': 'ItemResponseShape:#Exchange', 'BaseShape': 'IdOnly'},
                 'Traversal': 'Shallow',
                 'Paging': {'__type': 'CalendarPageView:#Exchange',
                            'StartDate': start.strftime('%Y-%m-%dT00:00:00.000'),
                            'EndDate': end.strftime('%Y-%m-%dT23:59:59.999')}}}
    url = BASE + '/service.svc?action=FindItem&app=PublishedCalendar&n=2'
    headers = {'Accept': '*/*', 'Action': 'FindItem', 'Content-Type': 'application/json; charset=utf-8',
               'x-req-source': 'PublishedCalendar',
               'x-owa-urlpostdata': json.dumps(payload, separators=(',', ':'))}
    req = urllib.request.Request(url, data=b'', headers=headers, method='POST')
    for attempt in range(4):
        try:
            with urllib.request.urlopen(req, timeout=45) as r:
                data = json.load(r)
            break
        except urllib.error.HTTPError:
            raise
        except (urllib.error.URLError, ConnectionError, TimeoutError):
            if attempt == 3:
                raise
            delay = 5 * (2 ** attempt)
            print(f'Calendar connection failed; retrying in {delay}s', flush=True)
            time.sleep(delay)
    messages = data['Body']['ResponseMessages']['Items']
    if not messages:
        raise RuntimeError('No Exchange response messages')
    result = messages[0]
    if result.get('ResponseCode') != 'NoError' or result.get('ResponseClass') != 'Success':
        raise RuntimeError('Exchange error: ' + repr(result.get('ResponseCode')))
    folder = result['RootFolder']
    if folder.get('IncludesLastItemInRange') is False:
        raise RuntimeError('Results truncated; refusing to publish incomplete calendar')
    items = folder.get('Items', [])
    if isinstance(items, dict):
        items = items.get('Items', [])
    if not isinstance(items, list):
        raise RuntimeError('Unexpected calendar items structure')
    return items


def event_id(item):
    # Prefer stable Exchange ID; recurrence occurrences may have distinct IDs.
    raw = item.get('UID') or item.get('ICalUid') or (item.get('ItemId') or {}).get('Id')
    if not raw:
        raise ValueError('Event without stable ID')
    import hashlib
    return hashlib.sha256(str(raw).encode()).hexdigest()[:32] + '@cesi-calendar-sync'


def main():
    today = dt.datetime.now(dt.timezone.utc).date()
    # Rolling window: previous 30 days to 365 days ahead, split into 35-day windows.
    start = today - dt.timedelta(days=30)
    last = today + dt.timedelta(days=365)
    all_events = {}
    cursor = start
    while cursor <= last:
        end = min(cursor + dt.timedelta(days=34), last)
        events = fetch(cursor, end)
        print(f'{cursor} through {end}: {len(events)} events')
        for event in events:
            if not isinstance(event, dict):
                continue
            uid = event_id(event)
            # Guard against conflicting duplicates in adjacent date windows.
            all_events[uid] = event
        cursor = end + dt.timedelta(days=1)
    if not all_events:
        raise RuntimeError('Zero events returned; refusing to overwrite existing feed')
    stamp = dt.datetime.now(dt.timezone.utc).strftime('%Y%m%dT%H%M%SZ')
    lines = ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//CESI Calendar Sync//EN',
             'CALSCALE:GREGORIAN', 'METHOD:PUBLISH', 'X-WR-CALNAME:CESI Timetable',
             'X-WR-TIMEZONE:Europe/Paris', 'REFRESH-INTERVAL;VALUE=DURATION:PT30M']
    valid = 0
    for uid, event in sorted(all_events.items()):
        try:
            begin, finish = parse_time(event['Start']), parse_time(event['End'])
        except (ValueError, KeyError) as exc:
            raise RuntimeError('Invalid event time, refusing partial feed') from exc
        lines += ['BEGIN:VEVENT', 'UID:' + uid, 'DTSTAMP:' + stamp,
                  'DTSTART:' + begin, 'DTEND:' + finish,
                  'SUMMARY:' + escape(event.get('Subject') or 'Untitled class')]
        location = event.get('Location')
        if isinstance(location, dict):
            location = location.get('DisplayName') or location.get('Name')
        if location:
            lines.append('LOCATION:' + escape(location))
        if event.get('IsCancelled'):
            lines.append('STATUS:CANCELLED')
        lines.append('END:VEVENT')
        valid += 1
    lines.append('END:VCALENDAR')
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    temp = OUTPUT.with_suffix('.tmp')
    temp.write_bytes(('\r\n'.join(fold(x) for x in lines) + '\r\n').encode('utf-8'))
    os.replace(temp, OUTPUT)
    print(f'Wrote {valid} events to {OUTPUT}')

if __name__ == '__main__':
    main()
