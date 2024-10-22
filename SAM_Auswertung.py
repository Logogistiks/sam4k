import os
import csv
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from time import sleep
from math import trunc
from datetime import datetime
from serial import Serial, PARITY_NONE, STOPBITS_ONE, EIGHTBITS

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

HEADER = '"Barcode";"Manueller Code";"Scheibentyp";"Anzahl Scheiben";"Teiler-Teilerfaktor";"Anzahl Einschüsse"'

pattern1 = PatternFill(start_color="00c2c2c2", end_color="00c2c2c2", fill_type="solid") # Grey
pattern2 = PatternFill(start_color="00abcdef", end_color="00abcdef", fill_type="solid") # Blue
pattern3 = PatternFill(start_color="00ff0000", end_color="00ff0000", fill_type="solid") # Red

def clear():
    """Clears the console"""
    os.system(("cls" if os.name == "nt" else "clear"))

def nowtime():
    """Returns the current time in YYYY_MM_DD-HH_MM_SS format"""
    return datetime.now().strftime("%Y_%m_%d-%H_%M_%S")

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
    result = []
    while True:
        clear()
        print("1) with decimal")
        print("2) truncate")
        print("3) with decimal, but truncate final score")
        mode = input("[1/2/3] >>> ")
        if mode in ["1", "2", "3"]:
            break
    with Serial(port="COM3", baudrate=9600, timeout=1, parity=PARITY_NONE, stopbits=STOPBITS_ONE, bytesize=EIGHTBITS, xonxoff=False, rtscts=False) as ser:
        try:
            ser.write(CODE_NOBAR)
            print("start")
            while True:
                ser.write(CODE_ENQ)
                response = ser.read(1)
                if response == CODE_NAK: # no result
                    sleep(0.5)
                    continue
                if response == CODE_STX: # transmission start
                    response = ser.read_until(CODE_ETB) # read the data part
                    response = response.replace(CODE_ETB, bytes())
                    # enclose data stream by double quotes and replace CR byte with ";"
                    data = b"\x22" + response.replace(CODE_CR, b"\x22\x3B\x22")[:-2] # remove 1x double quote and 1x semicolon at end of line
                    result.append(data.decode("unicode-escape"))
                    _ = ser.read_until(b"\x24") # read the rest, unimportant
                    ser.write(CODE_ACK) # com cycle finished
                print("transmission finished, insert more or press Ctrl + c (Strg + c) to stop")
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
