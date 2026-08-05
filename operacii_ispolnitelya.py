import os

import pandas as pd

from tablica import Tablica


def vybrat_nomer(kolichestvo, tekst, mozhno_nazad=True):
    vybor = input(tekst)
    while not (
        mozhno_nazad and vybor == "-1"
    ) and (not vybor.isdigit() or not 1 <= int(vybor) <= kolichestvo):
        vybor = input("Нет такого номера. Введите еще раз: ")
    if mozhno_nazad and vybor == "-1":
        return -1
    return int(vybor)


def sohranit_excel(tablica):
    imya_bez_rasshireniya = os.path.splitext(tablica.imya)[0]
    tablica.imya = imya_bez_rasshireniya + ".xlsx"
    tablica.df.to_excel(tablica.imya, index=False)


def sobrat_operacii_ispolnitelya(tablicy, papka_goda=".", god=""):
    istochniki = [
        tablica for tablica in tablicy if "Исполнитель" in tablica.df.columns
    ]

    if not istochniki:
        print("В текущей сессии нет таблиц с колонкой 'Исполнитель'")
        return None, False

    print("\nВыберите таблицу с операциями:")
    for nomer, tablica in enumerate(istochniki, start=1):
        print(nomer, "-", tablica.imya)
    print("-1 - вернуться в главное меню")

    nomer_istochnika = vybrat_nomer(
        len(istochniki), "Введите номер таблицы: "
    )
    if nomer_istochnika == -1:
        return None, False
    istochnik = istochniki[nomer_istochnika - 1]

    lyudi = {
        1: "Алексей",
        2: "Владимир",
        3: "Дмитрий",
        4: "Андрей",
    }

    print("\nИсполнители:")
    for nomer, chelovek in lyudi.items():
        print(nomer, "-", chelovek)
    print("-1 - вернуться в главное меню")

    nomer_cheloveka = vybrat_nomer(4, "Введите номер исполнителя: ")
    if nomer_cheloveka == -1:
        return None, False
    chelovek = lyudi[nomer_cheloveka]

    operacii = istochnik.df[
        istochnik.df["Исполнитель"].astype(str).str.strip().str.lower()
        == chelovek.lower()
    ].copy()

    imya_rezultata = os.path.join(papka_goda, f"{chelovek}_{god}.xlsx")

    if os.path.exists(imya_rezultata):
        print("\nНайден файл:", imya_rezultata)
        print("1 - дописать новые операции в конец")
        print("2 - ничего не делать")
        print("-1 - вернуться в главное меню")
        deystvie = vybrat_nomer(2, "Введите номер пункта: ")

        if deystvie in [-1, 2]:
            print("Сохранение отменено")
            return None, False

        try:
            staraya_tablica = pd.read_excel(imya_rezultata)
        except Exception as oshibka:
            print("Ошибка чтения файла:", oshibka)
            return None, False
    else:
        print("\nФайл", imya_rezultata, "пока не существует")
        print("1 - создать новую таблицу")
        print("2 - считать старую таблицу и дополнить ее")
        print("-1 - вернуться в главное меню")
        deystvie = vybrat_nomer(2, "Введите номер пункта: ")

        if deystvie == -1:
            return None, False

        if deystvie == 1:
            staraya_tablica = pd.DataFrame()
        else:
            imya_starogo = input("Введите название старого файла (-1 — назад): ")
            if imya_starogo.strip() == "-1":
                return None, False
            try:
                staraya_tablica = pd.read_excel(imya_starogo)
            except Exception as oshibka:
                print("Ошибка чтения файла:", oshibka)
                return None, False

    novyi_df = pd.concat(
        [staraya_tablica, operacii], ignore_index=True
    )

    # Если итоговый файл уже был загружен в этой сессии, обновляем его DataFrame.
    rezultat = None
    for tablica in tablicy:
        if os.path.abspath(tablica.imya) == os.path.abspath(imya_rezultata):
            rezultat = tablica
            rezultat.df = novyi_df
            break

    if rezultat is None:
        rezultat = Tablica(imya_rezultata, novyi_df)
        dobavit_v_sessiyu = True
    else:
        dobavit_v_sessiyu = False

    sohranit_excel(rezultat)
    print("Готово. Операции добавлены в файл:", rezultat.imya)
    return rezultat, dobavit_v_sessiyu
