from time import sleep
from serial import Serial, PARITY_NONE, STOPBITS_ONE, EIGHTBITS

CODE_STX = b"\x02"
CODE_ENQ = b"\x05"
CODE_ACK = b"\x06"
CODE_CR = b"\x0D"
CODE_NAK = b"\x15"
CODE_ETB = b"\x17"
CODE_EXIT = b"\xB0"
CODE_BAR = b"\xB1"
CODE_NOBAR = b"\xB2"

bytemap = {CODE_CR: b"\x22\x3B\x22", # CR -> ";"
           b"\x2E": b"\x2C"} # . -> ,

with Serial(port="/dev/ttyUSB0", baudrate=9600, timeout=1, parity=PARITY_NONE, stopbits=STOPBITS_ONE, bytesize=EIGHTBITS, xonxoff=False, rtscts=False) as ser:
    try:
        ser.write(CODE_NOBAR)
        print("start")
        while True:
            ser.write(CODE_ENQ)
            resp = ser.read(1)
            if resp == CODE_NAK: # no result
                sleep(0.5)
                continue
            if resp == CODE_STX: # transmission start
                data = b"\x22" # "
                while True:
                    byte = ser.read(1)
                    if byte == CODE_ETB: # end of data
                        break
                    data += bytemap.get(byte, byte)
                data += b"\x22" # "
                print(data)
            ser.write(CODE_ACK) # com cycle finished
            sleep(0.5)
    except KeyboardInterrupt:
        print("KeyboardInterrupt")
    finally:
        ser.write(CODE_EXIT) # set device inactive
        print("ende")