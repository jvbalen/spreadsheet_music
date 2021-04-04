# Spreadsheet Music

Use Google Sheets as a MIDI controller.

This python module provides a backend for live coding in a spreadsheet, by continuously
parsing a Google Sheet and sending out the results as local MIDI events.

### Requirements

- macOS
- Python 3.7+
- a DAW with MIDI-in support

### Installation:

The two main dependencies can be installed with pip:
```
pip install gspread_asyncio simplecoremidi
```

You will also need to create a Google sheets API service account, see [here](https://gspread.readthedocs.io/en/latest/oauth2.html#for-bots-using-service-account). Save its credentials as `client_secret.json` in this directory.

### Getting started

How to use:
- create a Google Sheets spreadsheet with read access for your service account
- run `main.py -n $SHEET_NAME` (see `python main.py -h` for more options)
- in your DAW, look for a new MIDI port with the same name as your spreadsheet and make sure to receive MIDI from it
- add the column label `pitch` to the first row of your spreadhseet, and any of the other ones listed below
- create notes by adding rows with MIDI-compatible values for each supported column
  e.g. pitch and velocity should be between 0 and 127

Supported spreadsheet columns (see also the included `Note` dataclass):
- `pitch`: MIDI note number in 0-127
- `velocity`: MIDI note velocity in 0-127
- `channel` MIDI channel, 1-based
- `onset`: note onset time relative to loop start, in seconds
- `duration`: note duration in seconds
- `loop`: loop note every `loop` seconds
- `probability`: probability of playing note in a given loop

### References

The `asyncio` logic in this module is loosely based on the queue example in [this tutorial](https://realpython.com/async-io-python/).

### License

Copyright © 2021, [Jan Van Balen](https://jvbalen.github.io). Released under the MIT License.
