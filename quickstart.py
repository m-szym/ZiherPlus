# ZiherPlus quickstart
# Author: Marek Szymański

# Wymagania:
# - Python3.6 lub nowszy
#   - pakiety Selenium i openpyxl (i ich zależności)

from ziher_plus import ZiherPlus, ZiherPlusSafeMode

# Żeby rzeczywiście importować dane do ZiHeRa użyj klasy ZiherPlus 
# (użycie jest dokładnie takie samo, ale pełna wersja posiada jeszcze dodatkowy tryb 
# pozwalający na importowanie danych bez ręcznego zatwierdzania każdego rekordu)
# ZiherPlusSafeMode to bezpieczna wersja demonstracyjna - pozbawiona możliwości wysyłania danych
# dzięki temu możesz sprawdzić poprawność importowania danych bez konsekwencji 

# Podstawowa konfiguracja zakłada, że arkusz w Excelu zorganizowany jest jak poniżej (dłuuga linia):
# nr rekordu | data | nr dokumentu | opis |                     wpływy                |                                            wydatki
#            |      |              |      | składki | statutowe | dotacje | pozostałe | wyposażenie | zużycie materiałów i energii | usługi obce | podatki i opłaty | wynagrodzenia | ubezpieczenia i inne świadczenia | ubezpieczenia majątkowe | podróże służbowe | wyżywienie | nagrody | bilety wstępu | noclegi | transport | inne
# !!! kolumna z datą musi być ustawiona jako 'datatime' w Excelu albo być w formacie YYYY-MM-DD !!!
# Możesz dopasować konfigurajcę do swojego arkusza zgodnie z instrukcjami w pliku 'excel_specific.py'

zp = ZiherPlusSafeMode.Firefox()                        # utwórz i skonfiguruj sterownik ZiherPlus, korzystający z Firefoxa
zp.login(email='jan.kowal@zhr.pl', password='hasło',    # zaloguj się ZiHeRa 
         region='pomorze')                              # (w wersji właściwej dla Twojego okręgu)
zp.load(filename='dane.xlsx', sheetname='lorem-ipsum')  # załaduj dane z arkusza 'lorem-ipsum' pliku 'dane.x' do programu 
                                                        # (jeszcze nie zostaną wysłane, jedynie wczytane przez ZiherPlus)
zp.send(logbook="bankowa", min_row=5, max_row=15)       # importuj dane od wiersza 5 do wiersza 15 do księgi bankowej
                                                        # po każdym rekordzie możesz zatwierdzić lub odrzucić jego import
zp.logout()                                             # wyloguj się z ZiHeRa
zp.quit()                                               # zamknij sterownik ZiherPlus i użytą przeglądarkę