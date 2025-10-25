from collections import Counter
from itertools import combinations
import xlrd
import pandas as pd
import random

workbook = xlrd.open_workbook('Euro Millions-archivio-estrazioni-2020-2024.xls', ignore_workbook_corruption=True)
df = pd.read_excel(workbook)   

# Supponiamo che le colonne dei numeri siano: 'N1', 'N2', 'N3', 'N4', 'N5', 'STAR1', 'STAR2'
numeri_cols = ['N1', 'N2', 'N3', 'N4', 'N5']  # Colonne dei numeri principali
numeri_star = ['STAR1', 'STAR2']  # Colonne delle stelle

# Estrai tutti i numeri
tutti_numeri = df[numeri_cols].values.flatten()

# Estrai tutti le stelle
tutti_stars = df[numeri_star].values.flatten()

# Frequenza dei singoli numeri
frequenza_numeri = Counter(tutti_numeri)
print("Frequenza numeri:")
for num, freq in frequenza_numeri.most_common():
    print(f"{num}: {freq} volte")

# Frequenza delle stelle
frequenza_stars = Counter(tutti_stars)
print("Frequenza stelle:")
for num, freq in frequenza_stars.most_common():
    print(f"{num}: {freq} volte")
    
# Frequenza degli ambi (coppie)
ambi = Counter()
for _, row in df[numeri_cols].iterrows():
    coppia = tuple(sorted(row))
    for ambo in combinations(coppia, 2):
        ambi[ambo] += 1

print(f"\nAmbi più frequenti (esempio top 5):")
for ambo, freq in ambi.most_common(5):
    print(f"{ambo}: {freq} volte")

# Frequenza dei terni
terni = Counter()
for _, row in df[numeri_cols].iterrows():
    terna = tuple(sorted(row))
    for terno in combinations(terna, 3):
        terni[terno] += 1

print(f"\nTerni più frequenti (esempio top 5):")
for terno, freq in terni.most_common(5):
    print(f"{terno}: {freq} volte")

# Previsione della prossima estrazione numeri (esempio: 5 numeri più frequenti)
numeri_previsti = [int(num) for num, _ in frequenza_numeri.most_common(15)]   
print(f"\nPrevisione basata su frequenza numeri: {numeri_previsti}")   

# Previsione della prossima estrazione stelle (esempio: 2 stelle più frequenti)
stelle_previsti = [int(num) for num, _ in frequenza_stars.most_common(6)]   
print(f"\nPrevisione basata su frequenza stelle: {stelle_previsti}")

# Combinazione casuale
for i in range(1, 2036):
    comninazione_casuale = random.sample(range(1, 51), 5) + random.sample(range(1, 13), 2)
print(f"\nCombinazione casuale (5 numeri + 2 stelle): {comninazione_casuale}")
