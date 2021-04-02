"""Spreadsheet Music: using Google Sheets as a MIDI controller
"""
import sys
import random
import logging
from typing import Iterable
from time import time, sleep
from dataclasses import dataclass
from argparse import ArgumentParser

import gspread
from oauth2client.service_account import ServiceAccountCredentials
from simplecoremidi import MIDISource

NOTE_ON = 144
NOTE_OFF = 128


@dataclass
class Note:
    pitch: int
    channel: int = 0
    loop: float = 1.0
    onset: float = 0.0
    duration: float = 0.1
    velocity: float = 64.0
    probability: float = 1.0


class Tables(object):

    def __init__(self,
                 sheet_name: str,
                 receive_interval: float = 60.0,
                 send_interval: float = 0.001,
                 secrets_file: str = "client_secret.json"):
        """Parse a spreadsheet into a sequence of Note objects and send them out as MIDI events.
        All notes are played back on loop, however each note can have its own loop length.

        Args:
        - sheet_name (str): name of spreadsheet in Google Sheet
            NOTE: requires a service account with access to this sheet
        - receive interval (float): how often to fetch and parse the spreadsheet, in seconds
            NOTE: this should probably be > 1 to avoid rate limiting issues (def: 1 second)
        - send interval (float): how often to send MIDI events from queue, in seconds
            NOTE: this should probably be < 10 ms to maintain rhyhtmic expressivity (def: 1 ms)
        - secrets_file (str): path to secrets JSON file for service account
        """
        self.receive_interval = receive_interval
        self.send_interval = send_interval
        self.sheet = get_sheet(sheet_name, secrets_file)
        self.midi_out = MIDISource(sheet_name)

    def run(self) -> None:
        """Main loop.

        Alternates between receiving notes every receive_interval seconds,
        and sending out MIDI events every `send_interval` seconds
        """
        notes = []
        start_time = time()
        last_received = start_time - self.receive_interval
        last_sent = start_time - self.send_interval
        while True:
            frame_time = time()
            if frame_time - last_received > self.receive_interval:
                notes = self.receive_notes()
                last_received = frame_time
            self.send_midi(notes, frame_time - start_time, last_sent - start_time)
            last_sent = frame_time
            sleep(self.send_interval)

    def receive_notes(self) -> Iterable[Note]:
        """Receive notes from spreadsheet

        Returns:
        - Iterable[Note]: iterable of Note objects
        """
        notes = []
        logging.debug('Parsing...')
        sheet = self.sheet.sheet1.get_all_values()
        for row in sheet[1:]:
            logging.debug(f'Row: {row}')
            try:
                note = note_from_dict(dict(zip(sheet[0], row)))
                logging.info(f'Adding {note}')
                notes.append(note)
            except (ValueError, TypeError) as e:
                logging.warning(f'Error caught during parsing: {e}')
                logging.warning(f'Note could not be parsed: {dict(zip(sheet[0], row))}')

        return notes

    def send_midi(self, notes, t_now, t_prev) -> None:
        """Send out any MIDI onsets or offsets that were due between t_prev and t_now.
        
        NOTE: the loop logic is handled by first converting all times to a kind of
        phase `phi` representing the notes' onset and offset as a number in (0, 1).

        Args:
        - notes (Iterable[Note]): iterable of Note objects
        - t_now (float): time of current iteration, since start time 
        - t_prev (float): time of previous iteration, since start time 
        """
        for note in notes:
            phi_now = t_now % note.loop
            phi_prev = t_prev % note.loop
            if (time_to_play(note.onset % note.loop, phi_now, phi_prev) and
                random.random() < note.probability):
                event = NOTE_ON + note.channel, note.pitch, note.velocity
                self.midi_out.send(event)
                logging.debug(f'NOTE_ON for note {note}')
            if time_to_play((note.onset + note.duration) % note.loop, phi_now, phi_prev):
                event = NOTE_OFF + note.channel, note.pitch, 0
                self.midi_out.send(event)
                logging.debug(f'NOTE_OFF for note {note}')


def time_to_play(phi_note, phi_now, phi_prev):
    """Return true if
    - phi_prev < phi_note <= phi_now, or
    - equivalent edge cases where phi_prev > phi_now because loops
    """
    return (phi_prev < phi_note <= phi_now or
            phi_note <= phi_now < phi_prev or
            phi_now < phi_prev <= phi_note)


def note_from_dict(d):

    types = Note.__annotations__
    d = {k: types[k](v) for k, v in d.items() if k in types and v != ''}
    if not len(d):
        raise ValueError('Empty row')
    note = Note(**d)
    if note.loop == 0.0:
        raise ValueError('Loop length is zero')

    return note


def get_sheet(sheet_name: str, secrets_file: str):

    scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_name(secrets_file, scope)
    client = gspread.authorize(credentials)

    return client.open(sheet_name)


if __name__ == '__main__':
    
    parser = ArgumentParser("Tables. Use Google Sheets as a MIDI controller.")
    parser.add_argument("--sheet-name", "-n", type=str, help="sheet name")
    parser.add_argument("--receive", "-r", type=float, default=60.0, help="receive interval in seconds")
    parser.add_argument("--send", "-s", type=float, default=0.001, help="send interval in seconds")
    parser.add_argument("--debug", "-d", action="store_true", help="set logging level to DEBUG")
    args = parser.parse_args()

    logging_level = logging.DEBUG if args.debug else logging.INFO
    logging_format, date_format = '%(asctime)s - %(message)s', '%H:%M:%S'
    logging.basicConfig(stream=sys.stdout, level=logging_level, format=logging_format, datefmt=date_format)

    tables = Tables(args.sheet_name, receive_interval=args.receive, send_interval=args.send)
    tables.run()
