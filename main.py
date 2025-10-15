from collections import Counter
from itertools import combinations
import xlrd
import pandas as pd

workbook = xlrd.open_workbook('Euro Millions-archivio-estrazioni-2024.xls', ignore_workbook_corruption=True)
df = pd.read_excel(workbook)   
# Leggi il file Excel (modifica il nome del file se necessario)
# df = pd.read_excel('Euro Millions-archivio-estrazioni-2024.xls', engine='xlrd', ignore_workbook_corruption=True)   # Assicurati che il file sia nella stessa cartella

# Supponiamo che le colonne dei numeri siano: 'N1', 'N2', 'N3', 'N4', 'N5'
numeri_cols = ['N1', 'N2', 'N3', 'N4', 'N5']

# Estrai tutti i numeri
tutti_numeri = df[numeri_cols].values.flatten()

# 1. Frequenza dei singoli numeri
frequenza_numeri = Counter(tutti_numeri)
print("Frequenza numeri:")
for num, freq in frequenza_numeri.most_common():
    print(f"{num}: {freq} volte")

# 2. Frequenza degli ambi (coppie)
ambi = Counter()
for _, row in df[numeri_cols].iterrows():
    coppia = tuple(sorted(row))
    for ambo in combinations(coppia, 2):
        ambi[ambo] += 1

print(f"\nAmbi più frequenti (esempio top 5):")
for ambo, freq in ambi.most_common(5):
    print(f"{ambo}: {freq} volte")

# 3. Frequenza dei terni
terni = Counter()
for _, row in df[numeri_cols].iterrows():
    terna = tuple(sorted(row))
    for terno in combinations(terna, 3):
        terni[terno] += 1

print(f"\nTerni più frequenti (esempio top 5):")
for terno, freq in terni.most_common(5):
    print(f"{terno}: {freq} volte")

# 4. Previsione della prossima estrazione (esempio: 5 numeri più frequenti)
numeri_previsti = [num for num, _ in frequenza_numeri.most_common(5)]
print(f"\nPrevisione basata su frequenza: {numeri_previsti}")   