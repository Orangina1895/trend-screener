import os
import pandas as pd

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

FILE1 = os.path.join(BASE_DIR, "testfile_1.xlsx")
FILE2 = os.path.join(BASE_DIR, "testfile_2.xlsx")

df = pd.DataFrame({
    "a": [1, 2, 3],
    "b": [4, 5, 6],
})

df.to_excel(FILE1, index=False)
df.to_excel(FILE2, index=False)

print("=== DEBUG ===")
print("Arbeitsverzeichnis:", BASE_DIR)
print("Dateien nach dem Schreiben:", os.listdir(BASE_DIR))
print("=== DEBUG ENDE ===")
