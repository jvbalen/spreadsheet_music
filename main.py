"""Spreadsheet Music: using Google Sheets as a MIDI controller
"""
import sys
import random
import logging
import asyncio
from dataclasses import dataclass
from argparse import ArgumentParser

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
    start_time = time()
    queue = asyncio.PriorityQueue()
    receiver = asyncio.create_task(
        receive(client_manager, sheet_name, queue, start_time=start_time)
    )
    sender = asyncio.create_task(
        send(queue, midi_out, interval=send_interval, start_time=start_time)
    )
    await asyncio.gather(receiver)


async def receive(
        client_manager: AsyncioGspreadClientManager,
        sheet_name: str,
        queue: asyncio.PriorityQueue,
        start_time: float = 0.0,
    ) -> None:

    client = await client_manager.authorize()
    sheet = await client.open(sheet_name)
    worksheet = await sheet.get_worksheet(0)
    while True:
        records = await worksheet.get_all_records()
        queue = await clear_onsets(queue)
        for record in records:
            t = time() - start_time
            try:
                note = note_from_dict(record)
                t_event = t // note.loop * note.loop + note.onset
                while t_event < t:
                    t_event += note.loop
                await queue.put((t_event, NOTE_ON, note))
                logging.debug(f'Note added @ {t_event + note.loop:.3f}: {note}')
            except (ValueError, TypeError) as e:
                logging.warning(f'Error while parsing row {record}: {e}')


async def send(
        queue: asyncio.PriorityQueue,
        midi_out: MIDISource,
        interval: float = 0.002,
        start_time: float = 0.0,
    ) -> None:

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


async def clear_onsets(queue) -> asyncio.PriorityQueue:

    # remove everything but save the offsets
    offsets = []
    while not queue.empty():
        t_event, message, note = await queue.get()
        if message == NOTE_OFF:
            offsets.append((t_event, message, note))

    # restore the offsets
    for t_event, message, note in offsets:
        await queue.put((t_event, message, note))

    return queue


def note_from_dict(d: dict) -> Note:

    types = Note.__annotations__
    d = {k: types[k](v) for k, v in d.items() if k in types and v != ''}
    if not len(d):
        raise ValueError('Empty row')
    note = Note(**d)
    if note.loop == 0.0:
        raise ValueError('Loop length is zero')

    return note


def get_credentials() -> Credentials:
    """Callback function that fetches credentials off disk.

    gspread_asyncio needs this to re-authenticate when credentials expire. To obtain a
    service account JSON file, follow these steps:
        https://gspread.readthedocs.io/en/latest/oauth2.html#for-bots-using-service-account
    """
    creds = Credentials.from_service_account_file("client_secret.json")
    scoped = creds.with_scopes([
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ])
    return scoped


if __name__ == "__main__":
    
    parser = ArgumentParser("Spreadsheet Music. Use Google Sheets as a MIDI controller.")
    parser.add_argument("--sheet-name", "-n", type=str, help="sheet name")
    parser.add_argument("--receive", "-r", type=float, default=4.0, help="receive interval in seconds")
    parser.add_argument("--send", "-s", type=float, default=0.002, help="send interval in seconds")
    parser.add_argument("--debug", "-d", action="store_true", help="set logging level to DEBUG")
    args = parser.parse_args()

    logging_level = logging.DEBUG if args.debug else logging.INFO
    logging_format, date_format = '%(asctime)s - %(message)s', '%H:%M:%S'
    logging.basicConfig(stream=sys.stdout, level=logging_level, format=logging_format, datefmt=date_format)

    client_manager = AsyncioGspreadClientManager(get_credentials, gspread_delay=args.receive)
    midi_out = MIDISource(args.sheet_name)
    asyncio.run(
        main(args.sheet_name, client_manager, midi_out, send_interval=args.send),
        debug=args.debug,
    )
