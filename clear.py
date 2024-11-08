import os
from time import sleep
import sys

if __name__ == "__main__":
    if input("Clear xlsx output files? [y/n] >>> ").lower() != "y":
        print("cancelled")
        sleep(1)
        sys.exit()

    for file in os.listdir():
        if file.endswith(".xlsx"):
            os.remove(file)
            #print(f"Removed {file}")
    print("Cleared output files!")

    if input("Also clear bin logfiles? [y/n] >>> ").lower() == "y":
        for file in os.listdir("log"):
            if file.endswith(".bin"):
                os.remove(os.path.join("log", file))
                #print(f"Removed {file}")