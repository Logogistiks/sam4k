"""Code by ChatGPT"""

import serial # Installiere die 'serial'-Bibliothek mit 'pip install pyserial'
import time
import keyboard  # Installiere die 'keyboard'-Bibliothek mit 'pip install keyboard'


print("Start in 3...")
time.sleep(1)
print("Start in 2...")
time.sleep(1)
print("Start in 1...")
time.sleep(1)

# Konfiguration der seriellen Schnittstelle
# ser = serial.Serial('COM1', 9600, timeout=1)  # 'COM1' durch den richtigen Port ersetzen
ser = serial.Serial(port='/dev/ttyUSB0', baudrate=9600, timeout=1, parity=serial.PARITY_NONE, stopbits=serial.STOPBITS_ONE, bytesize=serial.EIGHTBITS, xonxoff=False, rtscts=False)

ser.write(b'\xB2')  # Melde PC an ohne Barcodelesen (B1 wäre mit Barcode)

try:
    while True:
        # Sende ENQ
        # print("ENQ")
        ser.write(b'\x05')  # ASCII-Code für ENQ

        # Warte auf STX
        response = ser.read(1)
        if response == b'\x02':  # ASCII-Code für STX
            #print("STX empfangen")
            # Empfange die Daten
            data = b'\x22'  # Initialisiere den Datenpuffer
            while True:

                byte = ser.read(1)
                if byte == b'\x17':  # ETB-Zeichen
                    #print("ETB empfangen")
                    break
                # data += byte
                if byte == b'\x0D':  # Enter-Zeichen
                    data += b'\x22\x3B\x22'
                elif byte == b'\x2E':
                    data += b'\x2C'
                else:
                    data +=byte
            # Wenn Daten vorhanden sind, dekodiere und ausgeben
            #data += b'\x22\x0D\x0A'
            data += b'\x22'

            if data:
                line_text = data.decode('UTF-8')
                print(f"{line_text}")
                # Schreibe die empfangenen Daten in den Tastaturpuffer
                # keyboard.write(line_text)

            # Sende ACK als Bestätigung
            #print("ACK")
            ser.write(b'\x06')  # ASCII-Code für ACK

        time.sleep(0.5)  # Warte eine Sekunde vor dem nächsten Durchlauf

except KeyboardInterrupt:
    print("Programm wurde beendet.")

finally:
    ser.write(b'\xB0')  # Melde PC ab
    print("Ende")

    ser.close()  # Schließe die serielle Schnittstelle, wenn das Programm beendet wird
