import openpyxl
import time
import pyperclip


def znajdz_dane_w_pliku_excel(plik, imie, nazwisko):
    try:
        workbook = openpyxl.load_workbook(plik)
    except FileNotFoundError:
        print(f"Plik {plik} nie istnieje.")
        return
    except Exception as e:
        print(f"Wystąpił błąd podczas otwierania pliku: {str(e)}")
        return

    arkusz = workbook.active

    znaleziono = False

    for wiersz in arkusz.iter_rows(values_only=True):
        if wiersz[1] == imie and wiersz[2] == nazwisko:
            print("Znaleziono pasujące dane:")
            print(wiersz)
            wynik=str(wiersz)
            znaleziono = True
            pyperclip.copy(wynik)
            break

    if not znaleziono:
        print(f"Nie znaleziono danych dla: {imie} {nazwisko}")

    workbook.close()


if __name__ == "__main__":
    nazwa_pliku = "D:\kierowcy.xlsx"
    imie = input("Podaj imię: ").capitalize()
    nazwisko = input("Podaj nazwisko: ").capitalize()
    znajdz_dane_w_pliku_excel(nazwa_pliku, imie, nazwisko)

time.sleep(10)
