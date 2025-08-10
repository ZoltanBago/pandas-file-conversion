#!/usr/bin/env python
# coding: utf-8

# # Python Pandas Projekt | Fájlok konvertálása
# ---
# A Python Pandas könyvtár segítségével Excel és Calc fájlokat olvasunk be és konvertálunk át csv formátumba.
# 
# ![](logistics.jpg)
# 
# *Kép forrása: pixabay*

# ## 1. A szükséges könyvtárak telepítése
# Három Python könyvtárra van szükségünk:
# 
# - openpyxl
# - pyexcel-ods3
# - odfpy
# 
# Az `openpyxl` az Excel .xlsx fájlokhoz szükséges, a `pyexcel-ods3` és az `odfpy` pedig a Calc .ods formátumát kezeli.
# 
# Mindhárom könyvtárat a pip segítségével lehet telepíteni:

# In[ ]:


pip install openpyxl


# In[ ]:


pip install pyexcel-ods3


# In[ ]:


pip install odfpy


# ## 2. A csv fájl beolvasása az adatkeretbe  

# In[ ]:


# A Pandas könyvtár importálása alias formában
import pandas as pd

# A logi_data nevű változó létrehozása
logi_data = pd.read_csv("logistics_data.csv")

# A teljes adatkeret megjelenítése
logi_data


# A csv fájl oszlopai:
# 
# - **OrderID** - Egyedi azonosító a megrendeléshez
# - **CustomerName** - Az ügyfél neve
# - **ShipmentDate** - A szállítás dátuma
# - **DestinationCountry** - A megrendelő ország neve
# - **ProductCategory** - A termék kategóriája
# - **WeightKG** - A csomag súlya kilogrammban
# - **ShippingCostUSD** - A szállítási költség USD-ben
# - **Status** - A szállítás állapota

# ## 3. Konvertálás XLSX formátumba
# Az `.xlsx` a Microsoft Office Excel által támogatott fájlformátum. 

# In[ ]:


logi_data.to_excel("logi_data.xlsx")


# Ha megnyítjuk a `logi_data.xlsx` fájlt, akkor azt láthatjuk, hogy az `A1` cella üres. 
# 
# Ha azt szeretnénk, hogy a táblázatunk rendezettebb formában jelenjen meg, akkor az `openpyxl` és az `index=False` segítségével ezt megoldhatjuk.

# In[ ]:


# Az openpyxl importálása
import openpyxl


# In[ ]:


# Újra végezzük el a fenti müveletet 
logi_data.to_excel("logi_data.xlsx", index=False)


# Ezzel egy rendezettebb táblázatot kapunk, ahol az A1 cella már nem üres és jobban hasonlít az Excel tábla kinézetére.

# ## 4. Konvertálás ODS formátumba
# Az `.ods` a Libre Office Calc által használt formátum. 
# 
# Az ods esetében szükségünk lesz a `pyexcel-ods3` és az `odfpy` könyvtárra.  

# In[ ]:


import pyexcel_ods3
logi_data.to_excel("logi_data.ods", index=False, engine="odf")


# ## 5. Az XLSX fájl beolvasása

# In[ ]:


logi_xlsx = pd.read_excel("logi_data.xlsx")

logi_xlsx


# ## 6. Az ods fájl beolvasása

# In[ ]:


logi_ods = pd.read_excel("logi_data.ods")

logi_ods


# ## 7. Az XLSX fájl átalakítása CSV formátumba

# In[ ]:


logi_xlsx.to_csv("logi_xlsx.csv", index=False)


# ## 8. Az ODS fájl átalakítása CSV formátumba

# In[ ]:


logi_ods.to_csv("logi_ods.csv", index=False)


# ---
