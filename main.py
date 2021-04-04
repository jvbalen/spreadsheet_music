"""Spreadsheet Music: Google Sheets as a MIDI controller.

(c) Jan Van Balen, 2021
https://jvbalen.github.io

This simple app allows some form of live coding in a spreadsheet by continuously
parsing a Google Sheet and sending out the results as local MIDI events.

Requirements:
- macOS
- gspread_asyncio, simplecoremidi
- a Google sheets API service account, see:
  https://gspread.readthedocs.io/en/latest/oauth2.html#for-bots-using-service-account
- a Google Sheets spreadsheet with read access for service account
- a DAW with MIDI-in support

## How to use
- run this file. See argument help for details
- in your DAW, look for a new MIDI port with the same name as your spreadsheet
  and make sure you receive MIDI from it
- make a header row in the Google sheets spreadsheet and include a column named `pitch`
  and any of the other fields supported by the Note class
- create notes by adding rows with MIDI-compatible values for each of the fields
  e.g. pitch and velocity should be between 0 and 127

The async logic in this module is loosely based on the queue example in this tutorial:
    https://realpython.com/async-io-python/
"""
import sys
import random
import logging
import asyncio
from time import time
from functools import partial
from dataclasses import dataclass
from argparse import ArgumentParser, RawDescriptionHelpFormatter

from google.oauth2.service_account import Credentials 
from gspread_asyncio import AsyncioGspreadClientManager
from simplecoremidi import MIDISource

NOTE_ON = 144
NOTE_OFF = 128


@dataclass(order=True)
class Note:
    pitch: int
    channel: int = 1
    loop: float = 1.0
    onset: float = 0.0
    duration: float = 0.1
    velocity: float = 64.0
    probability: float = 1.0


async def main(
        sheet_name: str,
        client_manager: AsyncioGspreadClientManager,
        midi_out: MIDISource,
        send_interval: float = 0.002,
    ):
    """Main loop consisting of two asynchronous tasks:
    - a `receiver` that periodically checks the GSheet and puts notes on a queue
    - a `sender` that checks the queue and sends NOTE_ON an NOTE_OFF MIDI events

    Args:
    - sheet_name (str): name of Google Sheet to be parsed
    - client_manager (AsyncioGspreadClientManager): client manager with creds_fn
    - midi_out (MIDISource): where to send MIDI events
    - send_interval (float): minimum interval for sending MIDI messages, in seconds
    """
    start_time = time()
    queue = asyncio.PriorityQueue()
    receiver = asyncio.create_task(
        receive(sheet_name, client_manager, queue, start_time=start_time)
    )
    sender = asyncio.create_task(
        send(queue, midi_out, interval=send_interval, start_time=start_time)
    )
    await asyncio.gather(receiver)  # implicitly awaits sender as well


async def receive(
        sheet_name: str,
        client_manager: AsyncioGspreadClientManager,
        queue: asyncio.PriorityQueue,
        start_time: float = 0.0,
    ) -> None:
    """Periodically read Google Sheet, parse rows to Note and add to queue.

    Args:
    - sheet_name (str): name of Google Sheet to be parsed
    - client_manager (AsyncioGspreadClientManager): client manager with appropriate creds_fn
    - queue (asyncio.PriorityQueue): priority queue on which to place Note events
    - start_time (float): time() at the start of playback, optional
    """
    client = await client_manager.authorize()
    sheet = await client.open(sheet_name)
    worksheet = await sheet.get_worksheet(0)
    logging.info(f'Spreadsheet: https://docs.google.com/spreadsheets/d/{sheet.id}')
    logging.info(f'Listening...')
    while True:
        records = await worksheet.get_all_records()
        queue = await clear_onsets(queue)

        parse_start = time()
        for record in records:
            t = time() - start_time
            try:
                note = note_from_dict(record)
                t_event = t // note.loop * note.loop + note.onset % note.loop
                while t_event < t:
                    t_event += note.loop
                await queue.put((t_event, NOTE_ON, note))
                logging.debug(f'Note added @ {t_event + note.loop:.3f}: {note}')
            except (ValueError, TypeError) as e:
                logging.debug(f'Error while parsing row {record}: {e}')
        logging.info(f'Sheet parsed in {(time() - parse_start) * 1000:.3f} ms')


async def send(
        queue: asyncio.PriorityQueue,
        midi_out: MIDISource,
        interval: float = 0.002,
        start_time: float = 0.0,
    ) -> None:
    """Send out notes from queue as MIDI events

    Not a strict consumer, function also produces new queue items.
    On every NOTE_ON:
    - a new NOTE_OFF event is added to queue at t + note.duration
    - a new NOTE_ON event is added at t + note.loop

    Args:
    - queue (asyncio.PriorityQueue): priority queue containing Note events
    - midi_out (MIDISource): where to send MIDI events
    - interval (float): minimum interval for sending MIDI messages, in seconds
    - start_time (float): time() at the start of playback, optional
    """
    while True:
        t = time() - start_time
        t_event, message, note = await queue.get()
        if t > t_event:
            midi_event = (message + note.channel - 1, note.pitch, note.velocity)
            if message == NOTE_OFF:
                midi_out.send(midi_event)
            else:
                if random.random() < note.probability:
                    midi_out.send(midi_event)
                    logging.debug(f'MIDI out: {midi_event}')
                    await queue.put((t_event + note.duration, NOTE_OFF, note))
                await queue.put((t_event + note.loop, NOTE_ON, note))
                logging.debug(f'Note added @ {t_event + note.loop:.3f}: {note}')
        else:
            await queue.put((t_event, message, note))
            await asyncio.sleep(interval)


async def clear_onsets(queue: asyncio.PriorityQueue) -> asyncio.PriorityQueue:
    """Remove all NOTE_ON events from a queue.

    TODO: should this be synchronous? Doing this asynchronously might end
    up removing items that were not yet present when function was called?
    
    Args:
    - queue (asyncio.PriorityQueue)
    """
    offsets = []
    while not queue.empty():
        t_event, message, note = await queue.get()
        if message == NOTE_OFF:
            offsets.append((t_event, message, note))

    # restore the offsets
    for event in offsets:
        await queue.put(event)

    return queue


def note_from_dict(record: dict) -> Note:
    """Parse dictionary representing a Google sheets row, as a Note object
    Mostly wraps Note(**record), but also drops '' and enforces types, sanity checks.

    Args:
    - record (dict): spreadsheet row as a dict, with header as keys

    Returns:
    - Note: spreadsheet note parsed as a Note object
    """
    types = Note.__annotations__
    record = {k: types[k](v) for k, v in record.items() if k in types and v != ''}
    note = Note(**record)

    # sanity checks go here
    if note.loop == 0.0:
        raise ValueError('Loop length is zero')

    return note


def get_credentials(secrets_file: str) -> Credentials:
    """Callback function that fetches credentials off disk.

    gspread_asyncio needs this to re-authenticate when credentials expire. To obtain a
    service account JSON file, follow these steps:
        https://gspread.readthedocs.io/en/latest/oauth2.html#for-bots-using-service-account

    Args:
    - secrets_file (str): path to json file contain the servive account secrets
    """
    creds = Credentials.from_service_account_file(secrets_file)
    creds = creds.with_scopes([
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ])
    return creds


if __name__ == "__main__":
    
    parser = ArgumentParser(description=__doc__, formatter_class=RawDescriptionHelpFormatter)
    parser.add_argument("--sheet-name", "-n", help="name of Google Sheet to be parsed")
    parser.add_argument("--secrets-file", "-f", default="client_secret.json", help="secrets file")
    parser.add_argument("--receive", "-r", type=float, default=2.0, help="receive interval in seconds")
    parser.add_argument("--send", "-s", type=float, default=0.002, help="send interval in seconds")
    parser.add_argument("--debug", "-d", action="store_true", help="set logging level to DEBUG")
    args = parser.parse_args()

    log_level = logging.DEBUG if args.debug else logging.INFO
    log_format, date_format = '%(asctime)s - %(message)s', '%H:%M:%S'
    logging.basicConfig(stream=sys.stdout, level=log_level, format=log_format, datefmt=date_format)

    creds_fn = partial(get_credentials, secrets_file=args.secrets_file)
    client_manager = AsyncioGspreadClientManager(creds_fn, gspread_delay=args.receive)
    midi_out = MIDISource(args.sheet_name)
    asyncio.run(
        main(args.sheet_name, client_manager, midi_out, send_interval=args.send),
        debug=args.debug,
    )
