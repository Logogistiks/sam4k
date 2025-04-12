from __future__ import annotations

#built in modules
import os
import re
from copy import deepcopy
from datetime import datetime
from math import trunc
from time import sleep
from dataclasses import dataclass

#external modules
try:
    import openpyxl
    import openpyxl.cell
    import openpyxl.styles
    from serial import EIGHTBITS, PARITY_NONE, STOPBITS_ONE, Serial
    import beaupy
    from colorama import init, Fore
except ImportError as e:
    print(f"Error: {e}. Please install the required modules using 'pip install -r requirements.txt'")
    raise SystemExit

init(convert=True)

COM_CODES = [
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
"""Codes for communication with the SAM4000 device"""

SERIES_SHOTS_NUM = 10 # should be 1, 2, 5, or a multiple of 10
"""How many shots should be saved in a series"""

@dataclass
class Shot:
    """Represents a shot in the transmission"""
    ring: float | None = None
    div: float | None = None
    x: int | None = None
    y: int | None = None

    def __str__(self) -> str:
        """Returns a human readable representation of the Shot object"""
        return f"Shot(ring={self.ring}, div={self.div}, x={self.x}, y={self.y})"

class Transmission:
    "This class implements handling of a typical transmission by the SAM4000 device, which is received in bytes via serial connection."

    def __init__(self, barcode: str=None, manual_code: str=None, target_type: str=None, target_num: int=None, div: float=None, shots_num: int=None, shots: list[Shot]=None) -> None:
        """Initializes a Transmission object with the given parameters, allthough usage of `Transmission.create_empty` or `Transmission.from_bytes` is recommended."""
        self.barcode: str = barcode
        self.manual_code: str = manual_code
        self.target_type: str = target_type
        self.target_num: int = target_num
        self.div: float = div
        self.shots_num: int = shots_num
        self.shots: list[Shot] = shots

    @staticmethod
    def create_empty() -> Transmission:
        """returns an empty Transmission object with attributes of respective type instead of `None`"""
        return Transmission(barcode="", manual_code="", target_type="", target_num=0, div=0.0, shots_num=0, shots=[])

    def __str__(self) -> str:
        """Returns a human readable representation of the Transmission object"""
        res = ""
        res += f"Transmission(\n"
        res += f"    barcode={self.barcode},\n"
        res += f"    manual_code={self.manual_code},\n"
        res += f"    target_type={'None' if self.target_type is None else f'{self.target_type}'},\n"
        res += f"    target_num={self.target_num},\n"
        res += f"    div={self.div},\n"
        res += f"    shots_num={self.shots_num},\n"
        res += f"    shots=[\n"
        for i, shot in enumerate(self.shots):
            res += f"        {shot}{',' if i != len(self.shots)-1 else ''}\n"
        res += f"    ]\n"
        res += f")"
        return res

    @staticmethod
    def _valid_barcode(bc: str) -> bool:
        """Checks if a barcode string is of valid form"""
        return bool(re.fullmatch(r"[0-9]{8}", bc))

    @staticmethod
    def _valid_manual_code(mc: str) -> bool:
        """Checks if a manual code string is of valid form"""
        return bool(re.fullmatch(r"[0-9]{8}", mc))

    @staticmethod
    def _valid_target_type(tt: str) -> bool:
        """Checks if a target type string is of valid form"""
        return tt in ("LG", "LP", "KK", "ZS", "LS")

    @staticmethod
    def _valid_target_num(tn: str) -> bool:
        """Checks if a target number string is of valid form"""
        return bool(re.fullmatch(r"[0-9]{2}", tn))

    @staticmethod
    def _valid_div(div: str) -> bool:
        """Checks if a division factor string is of valid form"""
        return bool(re.fullmatch(r"[0-9]\.[0-9]", div))

    @staticmethod
    def _valid_shot_number(sn: str) -> bool:
        """Checks if a shot number string is of valid form"""
        return bool(re.fullmatch(r"[0-9]{2}", sn))

    @staticmethod
    def from_bytes(byt: bytes, log: bool=False) -> Transmission:
        """Parses the given bytes into a Transmission object and returns it."""
        if log:
            print(byt)
        trans = Transmission.create_empty()
        bc, mc, tt, tn, div, sn, *s = [part.decode("unicode-escape") for part in byt.split(CODE_CR)] # remove last empty string
        s = [item for item in s if item] # remove empty strings
        if log:
            for item in [bc, mc, tt, tn, div, sn, s]:
                print(item)
        if len(s) % 4 != 0: # s is a list of strings, each 4 strings represent a shot
            raise ValueError("bytes are of invalid form, shot data does not make sense (not a multiple of 4)")
        # technically the ? check is not necessary, but is left for clarity
        if not "?" in bc and Transmission._valid_barcode(bc):
            trans.barcode = bc
        if not "?" in mc and Transmission._valid_manual_code(mc):
            trans.manual_code = mc
        if not "?" in tt and Transmission._valid_target_type(tt):
            trans.target_type = tt
        if not "?" in tn and Transmission._valid_target_num(tn):
            trans.target_num = int(tn)
        if not "?" in div and Transmission._valid_div(div):
            trans.div = float(div)
        if not "?" in sn and Transmission._valid_shot_number(sn):
            trans.shots_num = int(sn)
        trans.shots = []
        for i in range(0, len(s), 4):
            trans.shots.append(Shot(
                ring=float(s[i]) if not "?" in s[i] else None,
                div=float(s[i+1]) if not "?" in s[i+1] else None,
                x=int(s[i+2]) if not "?" in s[i+2] else None,
                y=int(s[i+3]) if not "?" in s[i+3] else None
            ))
        #*Note: ↓ maybe useful later for distinguishing between cases: ↓
        #   ring is 0 and div is ? => missed shot
        #   ring > 0 und div is ? => manually corrected shot
        #   rind > 0 und Div > 0 => normal shot
        return trans

    def get_valid_shot_num(self) -> int:
        """Returns the number of valid shots in the transmission"""
        return sum(1 for shot in self.shots if shot.ring is not None)

    def get_invalid_shot_num(self) -> int:
        """Returns the number of invalid shots in the transmission"""
        return sum(1 for shot in self.shots if shot.ring is None)

    def get_manual_corrected_num(self) -> int:
        """Returns the number of shots that were manually corrected"""
        return sum(1 for shot in self.shots if shot.ring is not None and shot.div is None)

    def get_valid_shots(self, fill: int=None) -> list[Shot]:
        """Returns a list of valid shots in the transmission. \\
        If `fill` is given, pads the list with empty shots to the given length."""
        valid_shots = [shot for shot in self.shots if shot.ring is not None]
        if fill is not None and len(valid_shots) < fill:
            valid_shots += [Shot(0.0, None, None, None) for _ in range(fill - len(valid_shots))]
        return valid_shots

def clear() -> None:
    """Clears the console"""
    os.system(("cls" if os.name == "nt" else "clear"))

def nowtime() -> str:
    """Returns the current time in YYYY_MM_DD-HH_MM_SS format"""
    return datetime.now().strftime("%Y_%m_%d-%H_%M_%S")

def checksum_xor(byt: bytes) -> int:
    """Calculates the XOR checksum of the given bytes. \\
    Works by XORing all the bytes together."""
    chsum = 0
    for b in byt:
        chsum ^= b
    return chsum

def open_file(fname: str) -> None:
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

def save_data(shot_data: list[list[Shot]], mode: int, name_: str="") -> str:
    """Saves the data to an Excel file and returns the filename"""
    pattern_header = openpyxl.styles.PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid") # light blue
    pattern_mark1 = openpyxl.styles.PatternFill(start_color="FFF176", end_color="FFF176", fill_type="solid") # light yellow
    pattern_mark2 = openpyxl.styles.PatternFill(start_color="F08080", end_color="F08080", fill_type="solid") # light coral

    def set_cell(ws, row: int, col: int, value=None, fill=None, b_left: bool=False, b_right: bool=False, b_top: bool=False, b_bottom: bool=False, center_h: bool=False, center_v: bool=False) -> None:
        """Sets the value of a cell and applies the given fill and border settings"""
        cell: openpyxl.cell.Cell = ws.cell(row, col)
        if value is not None:
            cell.value = value
        if fill is not None:
            cell.fill = fill
        if any((b_left, b_right, b_top, b_bottom)):
            cell.border = openpyxl.styles.Border(
                left=openpyxl.styles.Side(style="thin") if b_left else None,
                right=openpyxl.styles.Side(style="thin") if b_right else None,
                top=openpyxl.styles.Side(style="thin") if b_top else None,
                bottom=openpyxl.styles.Side(style="thin") if b_bottom else None)
        if any((center_h, center_v)):
            cell.alignment = openpyxl.styles.Alignment(horizontal="center" if center_h else None, vertical="center" if center_v else None)

    wb = openpyxl.Workbook()
    ws = wb.active

    # write wireframe
    set_cell(ws, 2, 2, "Schuss", pattern_header, b_left=True, b_right=True, b_top=True, b_bottom=True) # just text
    for i in range(len(shot_data)):
        set_cell(ws, 3+i, 2, "Ringwert", pattern_header, b_left=True, b_right=True) # just text
        shot_range = f"C{3+i}:{chr(ord('C') + SERIES_SHOTS_NUM - 1)}{3+i}"
        if mode == 3:
            set_cell(ws, 3+i, 3+SERIES_SHOTS_NUM, f"=SUMPRODUCT(TRUNC({shot_range}))", b_left=True, b_right=True) # total sum
        else:
            set_cell(ws, 3+i, 3+SERIES_SHOTS_NUM, f"=SUM({shot_range})", b_left=True, b_right=True) # total sum
    set_cell(ws, 3+len(shot_data), 3+SERIES_SHOTS_NUM, f"=SUM({chr(ord('C') + SERIES_SHOTS_NUM)}3:{chr(ord('C') + SERIES_SHOTS_NUM)}{3+len(shot_data)-1})", b_left=True, b_right=True, b_top=True, b_bottom=True) # total total sum
    for i in range(SERIES_SHOTS_NUM):
        set_cell(ws, 2, 3+i, i+1, pattern_header, b_top=True, b_bottom=True, center_h=True) # just text
        set_cell(ws, 3+len(shot_data), 3+i, b_top=True) # just border
    set_cell(ws, 3+len(shot_data), 2, b_top=True) # just border
    set_cell(ws, 2, 3+SERIES_SHOTS_NUM, "Gesamt", pattern_header, b_left=True, b_right=True, b_top=True, b_bottom=True) # just text
    ws.merge_cells(start_row=3+len(shot_data)+1, start_column=2, end_row=3+len(shot_data)+1, end_column=3)
    set_cell(ws, 3+len(shot_data)+1, 2, "manuell korrigiert", pattern_mark1, center_h=True) # just text
    ws.merge_cells(start_row=3+len(shot_data)+1, start_column=4, end_row=3+len(shot_data)+1, end_column=5)
    set_cell(ws, 3+len(shot_data)+1, 4, "Fehlschuss", pattern_mark2, center_h=True) # just text

    # write data
    set_cell(ws, 1, 1, name_)
    for row, series in enumerate(shot_data):
        for col, shot in enumerate(series):
            value = trunc(shot.ring) if mode == 2 else shot.ring
            if shot.ring > 0 and shot.div is None: # manually corrected
                fill = pattern_mark1
            elif shot.ring == 0 and shot.div is None: # missed shot
                fill = pattern_mark2
            else: # normal shot
                fill = None
            set_cell(ws, 3+row, 3+col, value, fill, center_h=True)

    dir_year = datetime.now().strftime("%Y")
    if not os.path.exists(dir_year):
        os.mkdir(dir_year)
    dir_month = os.path.join(dir_year, datetime.now().strftime("%m"))
    if not os.path.exists(dir_month):
        os.mkdir(dir_month)
    fname = os.path.join(dir_month, f"output_{nowtime()}.xlsx")
    wb.save(fname)
    return str(fname)

def main(log: bool=False) -> None:
    if SERIES_SHOTS_NUM not in (1, 2, 5, 10) and SERIES_SHOTS_NUM % 10 != 0:
        raise ValueError("The number of shots in a series (SERIES_SHOTS_NUM) must be 1, 2, 5, or a multiple of 10")
    
    PORT = {"nt": "COM3", "posix": "/dev/ttyUSB0"}[os.name]
    if not os.path.exists(PORT): # checks if path is valid serial port before checking cwd
        print(f"Gerät nicht an Port {PORT} gefunden, bitte Kabelverbindung prüfen und Gerätemanager checken (IT rufen) -> ende")
        input("Drücke Enter zum Schließen")
        raise SystemExit(1)
    else:
        print(f"Anschluss {PORT} gefunden\n")

    name_ = beaupy.prompt("Name des Schützen eintippen:") # prompt text is cleared after execution
    print(f"Name des Schützen eintippen:\n> {Fore.LIGHTCYAN_EX}{name_}{Fore.RESET}\n")

    print("Schussanzahl pro Streifen mit Pfeiltasten auswählen und mit Enter bestätigen:")
    shots_per_target = beaupy.select([1, 2, 5, 10], cursor=">", cursor_style="bright_yellow", cursor_index=3)
    print(f"> {Fore.LIGHTCYAN_EX}{shots_per_target}{Fore.RESET}\n")

    modes = [
    "1) mit Teiler",
    "2) ohne Teiler",
    "3) Einzelergebnisse mit Teiler anzeigen, aber ohne Teiler summieren"
    ]
    print("Speicher-Modus mit Pfeiltasten auswählen und mit Enter bestätigen:")
    mode = int(beaupy.select(modes , cursor=">", cursor_style="bright_yellow", return_index=True)) + 1
    print(f"> {Fore.LIGHTCYAN_EX}{modes[mode-1]}{Fore.RESET}\n")

    with Serial(port=PORT, baudrate=9600, timeout=1, parity=PARITY_NONE, stopbits=STOPBITS_ONE, bytesize=EIGHTBITS, xonxoff=False, rtscts=False) as ser:
        try:
            ser.write(CODE_NOBAR)
            print("Gerät gefunden -> start")
            memory: list[Shot] = []
            result: list[list[Shot]] = []
            count = 0
            while True:
                ser.write(CODE_ENQ)
                response = ser.read(1)
                if response == b"":
                    continue
                if response == CODE_NAK: # no result
                    sleep(0.5)
                    continue
                if response == CODE_STX: # transmission start
                    retries = 0
                    while True:
                        response = ser.read_until(b"\x24")[:-1] # read until dollar sign exclusively
                        if log:
                            if not os.path.exists("log"):
                                os.mkdir("log")
                            with open(os.path.join("log", f"log-{nowtime()}.bin"), "wb") as f:
                                f.write(response)
                        data, checksum = response.split(CODE_ETB)
                        calc_checksum = checksum_xor(CODE_STX + data + CODE_ETB)
                        if calc_checksum != ord(checksum):
                            #print(f"ERROR: checksum doesnt match!")
                            #print(f"transmitted checksum: {ord(checksum)}")
                            #print(f"calculated checksum : {calc_checksum}")
                            if retries <= 3:
                                ser.write(CODE_NAK)
                                retries += 1
                                continue
                            else:
                                ser.write(CODE_ACK)
                                print("Fehler: Übertragung fehlerhaft, bitte Kabel auf Wackelkontakt o.ä. prüfen und Serie neu erfassen")
                                raise KeyboardInterrupt("'hack' to jump into catch block")
                        else:
                            break
                    trans = Transmission.from_bytes(data, log=False)

                    # extract valid data from transmission
                    if trans.get_valid_shot_num() > shots_per_target:
                        memory += trans.get_valid_shots()[:shots_per_target]
                    elif trans.get_valid_shot_num() < shots_per_target:
                        memory += trans.get_valid_shots(fill=shots_per_target)
                    else:
                        memory += trans.get_valid_shots()

                    # handle current transmission
                    if len(memory) > SERIES_SHOTS_NUM: # case should theoretically never happen
                        result.append(memory[:SERIES_SHOTS_NUM]) # discard the rest
                        memory.clear()
                    elif len(memory) < SERIES_SHOTS_NUM:
                        count += 1
                        print(f"Scheibe [{count}] übertragen, weitere einlegen oder Strg + c drücken, um Ergebnisse anzuzeigen")
                        continue
                    else:
                        result.append(deepcopy(memory))
                        memory.clear()

                    ser.write(CODE_ACK) # com cycle finished
                    #* Note:
                    # it is guaranteed that SERIES_SHOTS_NUM is a multiple of shots_per_target
                    # if num valid shots is more than shots per target, discard the rest
                    # if num valid shots is less than shots per target, fill with 0
                    # if this is equal to series num, save shots to result
                    # if this is less than series num, save shots to memory
                count += 1
                print(f"Scheibe [{count}] übertragen, weitere einlegen oder Strg + c drücken, um Ergebnisse anzuzeigen")
                sleep(0.5)
        except KeyboardInterrupt: # This isn't an error, but the intended way to exit the program
            try:
                print("KeyboardInterrupt")
                ser.write(CODE_EXIT) # set device inactive
                fname = save_data(result, mode, name_)
                open_file(fname)
            except Exception as e:
                print(f"Error occured during saving: {e}")
        except Exception as e:
            print(f"Error occured during runtime: {e}")

if __name__ == "__main__":
    main(log=False)
