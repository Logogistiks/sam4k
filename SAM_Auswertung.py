from __future__ import annotations
import os
import csv
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from time import sleep
from math import trunc
from datetime import datetime
from serial import Serial, PARITY_NONE, STOPBITS_ONE, EIGHTBITS
import re

#todo
# abfrage wieviel schuss pro scheibe soll
# wenn von maschine mehr schüsse als schüsseSOLL übermittelt, dann nur erste schüsseSOLL nehmen (kommt eigentlich nicht vor wegen manueller eingabe an maschine, aber nur fals halt einfach die letzten verwerfen)
# wenn von maschine weniger schüsse als schüsseSOLL übermittelt, mit 0 auffüllen
# nach jeder übermittlung prüfen ob min 10 schüsse (1 serie) (konstante definieren) da sind, wenn ja in result speichern
# -> rest wieder zurückhalten

#todo bytes aufteilen in objekte

# ringwert == 0 und teilerwert == ? => fehlschuss
# ringwert > 0 und teilerwert == ? => manuelle korrektur
# ringwert > 0 und teilerwert > 0 => normaler schuss

# evtl teiler und x/y-abstand weglassen

# communication codes
CODE_CTRLS = [
    CODE_STX := b"\x02",
    CODE_ENQ := b"\x05",
    CODE_ACK := b"\x06",
    CODE_CR := b"\x0D",
    CODE_NAK := b"\x15",
    CODE_ETB := b"\x17",
    CODE_EXIT := b"\xB0",
    CODE_BAR := b"\xB1",
    CODE_NOBAR := b"\xB2"
]

class Transmission:
    "This class implements datastorage of a typical transmission by the SAM4000 device, which is sent via a serial connection."
    def __init__(self, barcode: str=None, manual_code: str=None, target_type: str=None, target_num: int=None, div: float=None, shots_num: int=None, shots: list[dict[str, float | int]]=None) -> None:
        """Initializes a Transmission object with the given parameters, allthough `Transmission.from_bytes` should be used."""
        self.barcode: str = barcode
        self.manual_code: str = manual_code
        self.target_type: str = target_type
        self.target_num: int = target_num
        self.div: float = div
        self.shots_num: int = shots_num
        self.shots: list[dict[str, float | int]] = shots

    def _valid_barcode(self, bc: str) -> bool:
        """Checks if a barcode is of valid form"""
        return len(bc) == 8 and bc.isdecimal()

    def _valid_manual_code(self, mc: str) -> bool:
        """Checks if a manual code is of valid form"""
        return len(mc) == 8 and mc.isdecimal()

    def _valid_target_type(self, tt: str) -> bool:
        """Checks if a target type is of valid form"""
        return len(tt) == 2 and tt.isalpha()

    def _valid_target_num(self, tn: str) -> bool:
        """Checks if a target number is of valid form"""
        return len(tn) == 2 and tn.isdecimal()

    def _valid_div(self, div: str) -> bool:
        """Checks if a division factor is of valid form"""
        return len(div) == 3 and bool(re.fullmatch(r"\d\.\d", div))

    def _valid_shot_number(self, sn: str) -> bool:
        """Checks if a shot number is of valid form"""
        return len(sn) == 2 and sn.isdecimal()

    def from_bytes(self, byt: bytes) -> Transmission:
        """Parses the given bytes into a Transmission object. \\
        Returns the Transmission object itself to allow fluent style chaining."""
        bc, mc, tt, tn, div, sn, *s = [part.decode("unicode-escape") for part in byt.split(CODE_CR)]
        if len(s) % 4 != 0: # s is a list of strings, each 4 strings represent a shot
            raise ValueError("bytes are of invalid form, shot data does not make sense (not a multiple of 4)")
        # technically the ? check is not necessary, but is left for clarity
        if not "?" in bc and self._valid_barcode(bc):
            self.barcode = bc
        if not "?" in mc and self._valid_manual_code(mc):
            self.manual_code = mc
        if not "?" in tt and self._valid_target_type(tt):
            self.target_type = tt
        if not "?" in tn and self._valid_target_num(tn):
            self.target_num = int(tn)
        if not "?" in div and self._valid_div(div):
            self.div = float(div)
        if not "?" in sn and self._valid_shot_number(sn):
            self.shots_num = int(sn)
        self.shots = []
        for i in range(0, len(s), 4):
            self.shots.append({
                "ring": float(s[i]) if not "?" in s[i] else None,
                "div": float(s[i+1]) if not "?" in s[i+1] else None,
                "x": int(s[i+2]) if not "?" in s[i+2] else None,
                "y": int(s[i+3]) if not "?" in s[i+3] else None
            })
        return self
#todo implement dynamic templating for saving

def clear():
    """Clears the console"""
    os.system(("cls" if os.name == "nt" else "clear"))

def nowtime():
    """Returns the current time in YYYY_MM_DD-HH_MM_SS format"""
    return datetime.now().strftime("%Y_%m_%d-%H_%M_%S")

def checksum_xor(byt: bytes) -> int:
    """Calculates the XOR checksum of the given bytes. \\
    Works by XORing all the bytes together."""
    chsum = 0
    for b in byt:
        chsum ^= b
    return chsum

def modal(options: list[tuple[str, str]], prompt: str=">>> ", retry: bool=True) -> str:
    """Prints a modal dialog and returns the selected option. \\
    `options` should be passed as a list of tuples with the first element being the display text and the second element being the string the user has to enter to choose that option, this is case INsensitive
    Example Use: \\
    `modal([("Option 1", "1"), ("Option 2", "2")], prompt="Select an option: ")` \\
    Returns `None` if the user enters an invalid option and `retry` is set to False"""
    ans = None
    while True:
        clear()
        for text, code in options:
            print(text)
        ans = input(prompt if prompt.endswith(" ") else prompt + " ")
        if ans in [code.lower() for _, code in options] or not retry:
            return ans

def saveData(lst: list[str], mode: str) -> str:
    """Saves the data to an Excel file and returns the filename"""
    wb = Workbook()
    ws = wb.active
    values = [0]*len(lst)
    for row, line in enumerate(lst, start=1):
        print(f"Line: {repr(line)}")
        for col, v2 in enumerate(csv.reader([line], delimiter=";", quotechar='"').__next__(), start=1):
            if v2 in CODE_CTRLS:
                continue
            if row != 1 and col >= 7 and col % 4 == 3:
                if "?" in v2 or not v2:
                    v2 = "00.0"
                match mode:
                    case "1": # with decimal
                        ws.cell(row, col, v2).fill = pattern2
                        values[row-1] += float(v2)
                    case "2": # truncate
                        ws.cell(row, col, str(trunc(float(v2)))).fill = pattern2
                        values[row-1] += int(trunc(float(v2)))
                    case "3": # with decimal, but truncate final score
                        ws.cell(row, col, v2).fill = pattern2
                        values[row-1] += int(trunc(float(v2)))
            else:
                ws.cell(row, col, v2) # implicitly write header
    ws.insert_cols(idx=7)
    ws.cell(1, 7, "Gesamt").fill = pattern1
    for row, val in enumerate(values[1:], start=2): # ignore the header value which is 0
        ws.cell(row, 7, val).fill = pattern1
    ws.cell(row+1, 7, sum(values[1:])).fill = pattern3
    fname = f"output_{nowtime()}.xlsx"
    wb.save(fname)
    return fname

def fileOpen(fname: str):
    """Opens the file with the default program"""
    if os.name == "nt":
        os.startfile(fname, "open")
    elif os.name == "posix":
        try:
            os.system(f"xdg-open {fname}")
        except:
            print("Could not open the file")
    else:
        print("Could not open the file")

def main():
    HEADER = '"Barcode";"Manueller Code";"Scheibentyp";"Anzahl Scheiben";"Teiler-Teilerfaktor";"Anzahl Einschüsse"'
    global pattern1, pattern2, pattern3
    pattern1 = PatternFill(start_color="00c2c2c2", end_color="00c2c2c2", fill_type="solid") # Grey
    pattern2 = PatternFill(start_color="00abcdef", end_color="00abcdef", fill_type="solid") # Blue
    pattern3 = PatternFill(start_color="00ff0000", end_color="00ff0000", fill_type="solid") # Red

    PORT = {"nt": "COM3", "posix": "/dev/ttyUSB0"}[os.name]
    result = []
    #todo abfrage ideal anzahl schüsse pro streifen
    mode = modal([("1) with decimal", "1"), ("2) truncate", "2"), ("3) with decimal, but truncate final score", "3")], prompt="[1/2/3] >>> ")
    with Serial(port=PORT, baudrate=9600, timeout=1, parity=PARITY_NONE, stopbits=STOPBITS_ONE, bytesize=EIGHTBITS, xonxoff=False, rtscts=False) as ser:
        try:
            ser.write(CODE_NOBAR)
            print("start")
            count = 0
            while True:
                ser.write(CODE_ENQ)
                response = ser.read(1)
                if response == CODE_NAK: # no result
                    sleep(0.5)
                    continue
                if response == CODE_STX: # transmission start
                    response = ser.read_until(b"\x24")[:-1] # read until dollar sign exclusively
                    data, checksum = response.split(CODE_ETB)
                    calc_checksum = checksum_xor(CODE_STX + data + CODE_ETB)
                    if calc_checksum != checksum:
                        print(f"ERROR: checksum doesnt match!")
                        print(f"transmitted checksum: {checksum}")
                        print(f"calculated checksum : {calc_checksum}")
                        raise NotImplementedError #todo implement sending NAK and rereceiving data
                    trans = Transmission().from_bytes(data)

                    #response = response.replace(CODE_ETB, bytes())
                    # enclose data stream by double quotes and replace CR byte with ";"
                    #data = b"\x22" + response.replace(CODE_CR, b"\x22\x3B\x22")[:-2] # remove 1x double quote and 1x semicolon at end of line
                    #result.append(data.decode("unicode-escape"))
                    #_ = ser.read_until(b"\x24") # read the rest, unimportant
                    ser.write(CODE_ACK) # com cycle finished
                count += 1
                print(f"transmission [{count}] finished, insert more or press Ctrl + c (Strg + c) to stop")
                sleep(0.5)
        except KeyboardInterrupt:
            try:
                print("KeyboardInterrupt")
                ser.write(CODE_EXIT) # set device inactive
                print("ser being closed 3")
                ser.close()
                fname = saveData([HEADER]+result, mode)
                fileOpen(fname)
            except Exception as e:
                print(f"Error occured during saving: {e}")
                print("ser being closed 1")
                ser.close()
        except Exception as e:
            print(f"Error occured during runtime: {e}")
            print("ser being closed 2")
            ser.close()

if __name__ == "__main__":
    main()
