# ZiherPlus
..czyli automatyczny import danych z Excela do ZiHeRa. \
Bez ręcznego przepisywania cyferek i wypełniania formularzy. \
(Wciąż jest wolny, bo ZiHeR jest wolny - na to nic nie poradzę)

## Wymagania i zależności
- Python3.6 lub nowszy
- Selenium i openpyxl `> pip install -r requirements.txt`
- Firefox, inne przeglądarki nie są automatycznie wspierane (poniewż tylko Firefox od razu dostarcza niezbędnych sterowników, inne wymagają ręcznej instalacji)

## Użycie

Wszystko opisane jest w quickstart.py, 
kolejność kolumn można dopasować w excel_specific.py

W skrócie:
```python
zp = ZiherPlus().Firefox()     # utworzenie i inicjalizacja sterownika
zp.load(plik.xlsx, arkusz1)    # ładowanie danych do programu
zp.login(email, hasło, okrąg)  # logowanie do ZiHeRa
zp.send(k. bankowa, 15, 100)   # przesłanie danych od 15 do 100 rekordu
# dla pewności każdy rekord trzeba ręcznie zatwierdzić
# (można też opcję zatwierdzania wyłączyć)
zp.logout().quit()             # wylogowanie i zamknięcie sterownika
```

> **Uwaga!** Jeżeli skrypt nie wykrywa Firefoxa od razu to należy podać podać ścieżkę do pliku `firefox.exe` \
> `zp = ZiherPlus().Firefox("C:\\Program Files\\Mozilla Firefox\\firefox.exe")`

> **Uwaga!** Skrypt będzie działał jedynie tak długo jak nie zmieni się strona ZiHeRa. 
