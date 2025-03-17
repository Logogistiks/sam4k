"""This file is for testing stuff"""
from time import sleep
import os

from SAM_Auswertung import CODE_CR, Transmission, save_data, open_file, checksum_xor, CODE_ETB, CODE_STX

with open("temp2.txt", "r") as f:
    #print(repr(f.readlines()))
    lines = f.readlines()
    text1 = lines[-1].replace("\n", "").replace("\r", "")
    text2 = lines[-2].replace("\n", "").replace("\r", "")
byt1 = text1.encode("utf-8").replace(b"\x22\x3B\x22", CODE_CR).replace(b"\x22", b"")
byt2 = text2.encode("utf-8").replace(b"\x22\x3B\x22", CODE_CR).replace(b"\x22", b"")

print(byt1)

trans1 = Transmission().from_bytes(byt1)
#trans2 = Transmission().from_bytes(byt2)

print(trans1)

print(f"valid: {trans1.get_valid_shot_num()}")
#print(f"invalid: {trans.get_invalid_shot_num()}")
#print(f"manual: {trans.get_manual_corrected_num()}")
print(f"valid+fill: {trans1.get_valid_shots(fill=20)}")

"""open_file(save_data([trans1.shots, trans2.shots], mode=1))
while True:
    sleep(5)"""

"""with open("log\log.bin", "ab") as f:
    f.write(byt1)"""

"""logfiles = [os.path.join("log", filename) for filename in os.listdir("log") if filename.endswith(".bin")]
newest_logfile = min(logfiles, key=os.path.getctime)
with open(newest_logfile, "rb") as f:
    content = f.read()
    data, checksum = content.split(CODE_ETB)
    print(data, ord(checksum))
    print(checksum_xor(CODE_STX + data + CODE_ETB))
    #last_char = content[-1]
    #print(last_char)
    #print(hex(last_char))
    #print(checksum_xor(last_char))
    #print(ord(last_char))
    #print(hex(ord(last_char)))"""