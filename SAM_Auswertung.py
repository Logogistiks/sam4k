from __future__ import annotations

import os
import re
from datetime import datetime
from math import trunc
from time import sleep

import openpyxl
import openpyxl.cell
import openpyxl.styles
from serial import EIGHTBITS, PARITY_NONE, STOPBITS_ONE, Serial

# communication codes
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

SERIES_SHOTS_NUM = 10 # should be 1, 2, 5, or a multiple of 10

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

    def __str__(self) -> str:
        """Returns a string representation of the Transmission object"""
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

    def _valid_barcode(self, bc: str) -> bool:
        """Checks if a barcode is of valid form"""
        return bool(re.fullmatch(r"[0-9]{8}", bc))

    def _valid_manual_code(self, mc: str) -> bool:
        """Checks if a manual code is of valid form"""
        return bool(re.fullmatch(r"[0-9]{8}", mc))

    def _valid_target_type(self, tt: str) -> bool:
        """Checks if a target type is of valid form"""
        return tt in ("LG", "LP", "KK", "ZS", "LS")

    def _valid_target_num(self, tn: str) -> bool:
        """Checks if a target number is of valid form"""
        return bool(re.fullmatch(r"[0-9]{2}", tn))

    def _valid_div(self, div: str) -> bool:
        """Checks if a division factor is of valid form"""
        return bool(re.fullmatch(r"[0-9]\.[0-9]", div))

    def _valid_shot_number(self, sn: str) -> bool:
        """Checks if a shot number is of valid form"""
        return bool(re.fullmatch(r"[0-9]{2}", sn))

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
        # maybe useful for later distinguishingg between cases:
        #   ring is 0 and div is ? => missed shot
        #   ring > 0 und div is ? => manually corrected shot
        #   rind > 0 und Div > 0 => normal shot
        return self

    def get_valid_shot_num(self) -> int:
        """Returns the number of valid shots in the transmission"""
        return sum(1 for shot in self.shots if shot["ring"] is not None)

    def get_invalid_shot_num(self) -> int:
        """Returns the number of invalid shots in the transmission"""
        return sum(1 for shot in self.shots if shot["ring"] is None)

    def get_manual_corrected_num(self) -> int:
        """Returns the number of shots that were manually corrected"""
        return sum(1 for shot in self.shots if shot["ring"] is not None and shot["div"] is None)

    def get_valid_shots(self) -> list[dict[str, float | int]]:
        """Returns a list of valid shots in the transmission"""
        return [shot for shot in self.shots if shot["ring"] is not None]

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

def save_data(shot_data: list[list[dict[str, float | int]]], mode: str) -> str:
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
    set_cell(ws, 2, 2, "Schuss", pattern_header, b_left=True, b_right=True, b_top=True, b_bottom=True) # text
    for i in range(len(shot_data)):
        set_cell(ws, 3+i, 2, "Ringwert", pattern_header, b_left=True, b_right=True) # text
        shot_range = f"C{3+i}:{chr(ord('C') + SERIES_SHOTS_NUM - 1)}{3+i}"
        if mode == 3:
            set_cell(ws, 3+i, 3+SERIES_SHOTS_NUM, f"=SUMPRODUCT(TRUNC({shot_range}))", b_left=True, b_right=True) # total sum
        else:
            set_cell(ws, 3+i, 3+SERIES_SHOTS_NUM, f"=SUM({shot_range})", b_left=True, b_right=True) # total sum
    set_cell(ws, 3+len(shot_data), 3+SERIES_SHOTS_NUM, f"=SUM({chr(ord('C') + SERIES_SHOTS_NUM)}3:{chr(ord('C') + SERIES_SHOTS_NUM)}{3+len(shot_data)-1})", b_left=True, b_right=True, b_top=True, b_bottom=True) # total total sum
    for i in range(SERIES_SHOTS_NUM):
        set_cell(ws, 2, 3+i, i+1, pattern_header, b_top=True, b_bottom=True, center_h=True) # text
        set_cell(ws, 3+len(shot_data), 3+i, b_top=True) # just border
    set_cell(ws, 3+len(shot_data), 2, b_top=True) # just border
    set_cell(ws, 2, 3+SERIES_SHOTS_NUM, "Gesamt", pattern_header, b_left=True, b_right=True, b_top=True, b_bottom=True) # text
    ws.merge_cells(start_row=3+len(shot_data)+1, start_column=2, end_row=3+len(shot_data)+1, end_column=3)
    set_cell(ws, 3+len(shot_data)+1, 2, "manuell korrigiert", pattern_mark1, center_h=True) # text
    ws.merge_cells(start_row=3+len(shot_data)+1, start_column=4, end_row=3+len(shot_data)+1, end_column=5)
    set_cell(ws, 3+len(shot_data)+1, 4, "Fehlschuss", pattern_mark2, center_h=True) # text

    # write data
    for r, series in enumerate(shot_data):
        for c, shot in enumerate(series):
            value = trunc(shot["ring"]) if mode == 2 else shot["ring"]
            if shot["ring"] > 0 and shot["div"] is None: # manually corrected
                fill = pattern_mark1
            elif shot["ring"] == 0 and shot["div"] is None: # missed shot
                fill = pattern_mark2
            else: # normal shot
                fill = None
            set_cell(ws, 3+r, 3+c, value, fill, center_h=True)

    fname = f"output_{nowtime()}.xlsx"
    wb.save(fname)
    return fname

def main() -> None:
    if SERIES_SHOTS_NUM not in (1, 2, 5, 10) and SERIES_SHOTS_NUM % 10 != 0:
        raise ValueError("The number of shots in a series (SERIES_SHOTS_NUM) must be 1, 2, 5, or a multiple of 10")
    global pattern1, pattern2, pattern3
    pattern1 = openpyxl.styles.PatternFill(start_color="00c2c2c2", end_color="00c2c2c2", fill_type="solid") # Grey
    pattern2 = openpyxl.styles.PatternFill(start_color="00abcdef", end_color="00abcdef", fill_type="solid") # Blue
    pattern3 = openpyxl.styles.PatternFill(start_color="00ff0000", end_color="00ff0000", fill_type="solid") # Red

    print("Please enter the supposed number of shots per target:")
    shots_per_target = int(modal([("[1]", "1"), ("[2]", "2"), ("[5]", "5"), ("[10]", "10")], prompt="[1/2/5/10] >>> "))

    print("Please select the mode of operation:")
    mode = modal([("1) with decimal", "1"), ("2) truncate", "2"), ("3) with decimal, but truncate final score", "3")], prompt="[1/2/3] >>> ")

    PORT = {"nt": "COM3", "posix": "/dev/ttyUSB0"}[os.name]
    with Serial(port=PORT, baudrate=9600, timeout=1, parity=PARITY_NONE, stopbits=STOPBITS_ONE, bytesize=EIGHTBITS, xonxoff=False, rtscts=False) as ser:
        try:
            ser.write(CODE_NOBAR)
            print("start")
            result: list[list[dict[str, float | int]]] = []
            memory: list[dict[str, float | int]] = []
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
                        #? needs data on what gets sent after sending NAK
                    trans = Transmission().from_bytes(data)

                    # extract valid data from transmission
                    if trans.get_valid_shot_num() > shots_per_target:
                        shotlist = trans.get_valid_shots()[:shots_per_target]
                    elif trans.get_valid_shot_num() < shots_per_target:
                        shotlist = trans.get_valid_shots() + [{"ring": 0.0, "div": None, "x": None, "y": None} for _ in range(shots_per_target)]
                    else:
                        shotlist = trans.get_valid_shots()

                    # handle current transmission
                    memory += shotlist
                    if len(memory) > SERIES_SHOTS_NUM: # case should never happen
                        result.append(memory[:SERIES_SHOTS_NUM]) # discard the rest
                        memory.clear()
                    elif len(memory) < SERIES_SHOTS_NUM:
                        continue
                    else:
                        result.append(memory)
                        memory.clear()

                    ser.write(CODE_ACK) # com cycle finished
                    #ser.write(CODE_NAK) #todo test what data gets resend
                    #print(ser.read_until(b"\x24"))

                    #* Note:
                    # it is guaranteed that SERIES_SHOTS_NUM is a multiple of shots_per_target
                    # if num valid shots is more than shots per target, discard the rest
                    # if num valid shots is less than shots per target, fill with 0
                    # if this is equal to series num, save shots to result
                    # if this is less than series num, save shots to memory
                count += 1
                print(f"transmission [{count}] finished, insert more or press Ctrl + c (Strg + c) to stop")
                sleep(0.5)
        except KeyboardInterrupt:
            try:
                print("KeyboardInterrupt")
                ser.write(CODE_EXIT) # set device inactive
                fname = save_data(result, mode)
                open_file(fname)
            except Exception as e:
                print(f"Error occured during saving: {e}")
        except Exception as e:
            print(f"Error occured during runtime: {e}")

if __name__ == "__main__":
    main()