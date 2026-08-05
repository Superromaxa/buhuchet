import os

import pandas as pd

from tablica import Tablica


def sostavit_premii(tablicy, papka_rezultatov="."):
    tablicy_s_ispolnitelyami = [
        tablica
        for tablica in tablicy
        if "_с_исполнителями" in tablica.imya
    ]

    if not tablicy_s_ispolnitelyami:
        print("В текущей сессии нет файлов с исполнителями")
        return None

    print("\nФайлы с исполнителями:")
    for nomer, tablica in enumerate(tablicy_s_ispolnitelyami, start=1):
        print(nomer, "-", tablica.imya)

    print("Можно выбрать несколько файлов, например: 1, 2, 4")
    print("-1 - вернуться в главное меню")
    vybor = input("Введите номера файлов: ").strip()

    if vybor == "-1":
        return None

    while True:
        try:
            nomera = [int(nomer.strip()) for nomer in vybor.split(",")]
            nomera = list(dict.fromkeys(nomera))

            if nomera and all(
                1 <= nomer <= len(tablicy_s_ispolnitelyami)
                for nomer in nomera
            ):
                break
        except ValueError:
            pass

        vybor = input("Нет таких номеров. Введите еще раз: ").strip()
        if vybor == "-1":
            return None

    operacii_prishlo = []

    for nomer in nomera:
        tablica = tablicy_s_ispolnitelyami[nomer - 1]

        if "Тип операции" not in tablica.df.columns:
            continue

        stroki = tablica.df[
            tablica.df["Тип операции"].astype(str).str.strip().str.lower()
            == "пришло"
        ].copy()
        operacii_prishlo.append(stroki)

    if operacii_prishlo:
        rezultat = pd.concat(operacii_prishlo, ignore_index=True)
    else:
        rezultat = pd.DataFrame()

    imya_rezultata = os.path.join(papka_rezultatov, "премии_пришло.xlsx")
    rezultat.to_excel(imya_rezultata, index=False)

    print("\nГотово. Результат сохранен в файле:", imya_rezultata)
    return Tablica(imya_rezultata, rezultat)
