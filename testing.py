"""This file is for testing stuff"""
from time import sleep
import os

from SAM_Auswertung import CODE_CR, Transmission, save_data, open_file

"""with open("temp2.txt", "r") as f:
    #print(repr(f.readlines()))
    lines = f.readlines()
    text1 = lines[-1].replace("\n", "").replace("\r", "")
    text2 = lines[-2].replace("\n", "").replace("\r", "")
byt1 = text1.encode("utf-8").replace(b"\x22\x3B\x22", CODE_CR).replace(b"\x22", b"")
byt2 = text2.encode("utf-8").replace(b"\x22\x3B\x22", CODE_CR).replace(b"\x22", b"")"""

#print(byt)

#trans1 = Transmission().from_bytes(byt1)
#trans2 = Transmission().from_bytes(byt2)

#print(trans)

#print(f"valid: {trans.get_valid_shot_num()}")
#print(f"invalid: {trans.get_invalid_shot_num()}")
#print(f"manual: {trans.get_manual_corrected_num()}")

"""open_file(save_data([trans1.shots, trans2.shots], mode=1))
while True:
    sleep(5)"""

"""with open("log\log.bin", "ab") as f:
    f.write(byt1)"""

logfiles = [os.path.join("log", filename) for filename in os.listdir("log") if filename.endswith(".bin")]
newest_logfile = max(logfiles, key=os.path.getctime)
with open("log\\log.bin", "rb") as f:
    content = f.read()
    print(content)
    print(CODE_CR in content)