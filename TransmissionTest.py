"""This file tests the Transmission class with the temp data"""

from SAM_Auswertung import CODE_CR, Transmission

with open("temp.txt", "r") as f:
    text = f.readlines()[-1][1:-1].replace("\n", "").replace("\r", "")
byt = text.encode("utf-8").replace(b"\x22\x3B\x22", CODE_CR)

print(byt)

trans = Transmission().from_bytes(byt)

print(trans.__dict__)