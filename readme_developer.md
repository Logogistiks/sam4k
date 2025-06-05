# SAM4K Developer Info

This section is intended for developers setting up / maintaining the project, and is therefore written in english like the code itself.

## 📋Table of contents

- #### [ℹ️Overview](#ℹ%EF%B8%8Foverview-1)
- #### [📦Installation & Setup](#installation--setup)
- #### [📫Communication with SAM4000](#communication-with-sam4000-1)
- #### [🛑Exit codes](#exit-codes-1)
- #### [📚Documentation](#documentation-1)

## ℹ️Overview

This project is made with focus on Windows. \
Theoretically, everything is coded cross-platform, but its not tested outside of windows.

### Basic functionality:
1. transmitted bytedata is parsed as a [Transmission](#transmission) object
1. shot data of transmission is collected in short-term [Memory](#memoryhandler)
1. if short-term memory holds enough shots for a full series, its moved into long-term memory

There is a long-term memory for every entered person.

### Saving:
On quitting, each persons data is saved in its own section in the same output `.xlsx` file. \
The filenames contains the current date and time. \
They are located in the main directory, sorted in folders by **year** and **month**. \
The file is opened automatically on quitting.

## 📦Installation & Setup

1. Clone the repository:
    ```
    git clone https://github.com/Logogistiks/sam4k.git
    ```
1. Install dependencies:
    ```
    pip install -r requirements.txt
    ```
1. Project is ready. Run `SAM_Auswertung.py` to get started.

## 📫Communication with SAM4000

The communication with the device is based on the RS 232 protocol, as written in [the manual](https://github.com/Logogistiks/sam4k/blob/main/SAM4000_Anleitung.pdf) p. 31.

### Communication codes used:

| code    | hex  | sender   | meaning
| ------- | ---- | -------- | -------
| STX     | 0x02 | SAM      | result is ready!
| ENQ     | 0x05 | PC       | is result ready?
| ACK     | 0x06 | PC       | data received correctly
| CR      | 0x0D | SAM      | separator in data block
| NAK     | 0x15 | SAM / PC | result not ready / data not received correctly
| ETB     | 0x17 | SAM      | separates data and checksum
| _EXIT_  | 0xB0 | PC       | log out, SAM goes inactive
| _BAR_   | 0xB1 | PC       | sign in, barcode is used
| _NOBAR_ | 0xB2 | PC       | sign in, barcode is ignored

### Communication process:

* PC sends `BAR` or `NOBAR`
* Loop:
    * PC sends `ENQ`
    * If new strip-data is not available
        * SAM sends `NAK`
    * If new strip-data is available:
        * SAM sends `STX`
        * SAM sends data
        * SAM sends `ETB`
        * SAM sends checksum
        * SAM sends end-byte "$"
        * If data correct: PC sends `ACK`
        * If data not correct: PC sends `NAK`, SAM repeats **from data til "$"**

### Structure of data block:

* 8 bytes: barcode (ASCII 0-9)
* 1 byte: `CR`
* 8 bytes: manual code (ASCII 0-9)
* 1 byte: `CR`
* 2 bytes: target type (ASCII "LG" / "LP" / "KK" / "ZS")
* 1 byte: `CR`
* 2 bytes: number of targets (ASCII "XX")
* 1 byte: `CR`
* 3 bytes: div-divfactor (ASCII "X.X")
* 1 byte: `CR`
* 2 bytes: number of shots (ASCII "XX")
* 1 byte: `CR`
* Repeats "number of targets" times:
    * 4 bytes: ringvalue (ASCII "XX.X")
    * 1 byte: `CR`
    * 6 bytes: div (ASCII "XXXX.X")
    * 1 byte: `CR`
    * 5 bytes: X-distance (ASCII "±XXXX")
    * 1 byte: `CR`
    * 5 bytes: Y-distance (ASCII "±XXXX")
    * 1 byte: `CR`

Missing values are transmitted as "?" (number of figures is preserved). \
All number values are padded with leading zeroes. \
Divisor, X- and Y-distance are in 1/100 mm from target center.

## 🛑Exit codes

All occuring errors should be catched and printed to console.  \
If that happens, the program terminates gracefully with a specific exit code:

| code | cause | fix
| ---- | ----- | ---
| 0    | no data to save on quitting | ---
| 2    | cant import external library | `pip install -r requirements.txt`
| 10   | configuration error with `SHOTS_PER_SERIES` | edit `SHOTS_PER_SERIES` at top of file
| 20   | configured serial port not found | check serial ports, set another
| 30   | received empty response from serial | check that device is on, check cable connection
| 99   | non accounted error occured | ...

## 📚Documentation

List of selected objects and what they do.

### Constants

* #### `PORT`

    The serial port that the device is connected to.

* #### `SHOTS_PER_SERIES`

    How many shots should be saved in a series (one row in the excel file).

* #### `PATTERN_HEADER`

    Which color / gradient the table-head is colored.

* #### `PATTERN_MARK1`

    Which color / gradient the manually corrected shots are colored.

* #### `PATTERN_MARK2`

    Which color / gradient the missing shots are colored in.

* #### `LOG_TRANSMISSIONS`

    Whether to log the raw bytes received from the device.

* #### `CHSUM_RETRY`

    How many times to retry fetching the transmission on wrong checksum.

* #### `COM_CODES`

    List of communication codes used by this project.

### Classes

* #### `SHOT()`

    Dataclass to store one shot.

* #### `Transmission()`

    Handles a byte transmission from the SAM device

    ##### Methods

    * `.create_empty()`

        Staticmethod that returns an empty transmission object with attributes of respective type.

    * `.example()`

        Staticmethod that returns an example transmission object filled with dummy data.

    * `.from_bytes()`

        Staticmethod that returns a transmission object with data parsed from bytes.

    * `.receive()`

        Staticmethod that returns a transmission object with data read from a serial object.

* #### `MemoryHandler()`

    Dynamic memory for shot-data, manages multiple people.

    ##### Methods

    * `.update_person()`

        Classmethod that sets up a new person in memory.

    * `.update_memory()`

        Classmethod that manages adding a new transmission to memory.

    * `.finalize()`

        Classmethod that prepares the memory to be saved by deleting empty people.

### Functions

* `log()`

    Logs string or byte data to a file in the log directory.

* `clear()`

    Clears the console.

* `nowtime()`

    Returns the current date and time, optionally human readable.

* `checksum_xor()`

    Calculates the xor checksum for given bytes, as done by the SAM device.

* `record_keypresses()`

    Listens for keypresses in given timeframe and returns them as list.

* `open_file()`

    Opens a given filepath with the systems default program.

* `set_cell()`

    Sets a cell in an excel worksheet with the given parameters.

* `draw_header()`

    Draws the header on an excel worksheet.

* `draw_wireframe()`

    Draws the wireframe on an excel worksheet.

* `fill_wireframe()`

    Fills the wireframe with given data.

* `save_data()`

    Saves the given data to an excel file and returns the filepath.

---

### [Go to top](#sam4k-developer-info)
