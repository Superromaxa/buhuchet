import os

import pandas as pd

from tablica import Tablica


def vybrat_nomer(kolichestvo, tekst):
    vybor = input(tekst).strip()
    while vybor != "-1" and (
        not vybor.isdigit() or not 1 <= int(vybor) <= kolichestvo
    ):
        vybor = input("Нет такого номера. Введите еще раз: ").strip()

    if vybor == "-1":
        return -1
    return int(vybor)


def pereimenovat_fayl(tablicy):
    if not tablicy:
        print("В текущей сессии пока нет загруженных таблиц")
        return

    print("\nВыберите файл для переименования:")
    for nomer, tablica in enumerate(tablicy, start=1):
        print(nomer, "-", tablica.imya)
    print("-1 - вернуться назад")

    nomer = vybrat_nomer(len(tablicy), "Введите номер файла: ")
    if nomer == -1:
        return

    tablica = tablicy[nomer - 1]
    staroe_imya = tablica.imya
    novoe_imya = input("Введите новое имя файла (-1 — назад): ").strip()
    if novoe_imya == "-1":
        return
    if not novoe_imya:
        print("Имя файла не может быть пустым")
        return
    novoe_imya = os.path.basename(novoe_imya)

    staroe_rasshirenie = os.path.splitext(staroe_imya)[1]
    if not os.path.splitext(novoe_imya)[1]:
        novoe_imya += staroe_rasshirenie or ".xlsx"

    novyi_put = os.path.join(os.path.dirname(staroe_imya), novoe_imya)
    if os.path.exists(novyi_put) and os.path.abspath(novyi_put) != os.path.abspath(staroe_imya):
        print("Файл с таким именем уже существует")
        return

    if os.path.exists(staroe_imya):
        os.rename(staroe_imya, novyi_put)

    # Обновляем имя у всех DataFrame, которые ссылались на этот файл.
    for drugaya_tablica in tablicy:
        if os.path.abspath(drugaya_tablica.imya) == os.path.abspath(staroe_imya):
            drugaya_tablica.imya = novyi_put

    print("Файл переименован:", novyi_put)


def dozagruzit_tablicy(tablicy, papka_mesaca, papka_goda):
    zagruzhennye = {os.path.abspath(tablica.imya) for tablica in tablicy}
    dostupnie = []

    for papka in [papka_mesaca, papka_goda]:
        if not os.path.isdir(papka):
            continue

        for imya in sorted(os.listdir(papka)):
            polnoe_imya = os.path.join(papka, imya)
            if not os.path.isfile(polnoe_imya):
                continue
            if os.path.splitext(imya)[1].lower() not in [".xlsx", ".xls", ".csv"]:
                continue
            if os.path.abspath(polnoe_imya) not in zagruzhennye:
                dostupnie.append(polnoe_imya)

    if not dostupnie:
        print("Все доступные таблицы уже загружены")
        return

    print("\nТаблицы, которые можно дозагрузить:")
    for nomer, imya in enumerate(dostupnie, start=1):
        print(nomer, "-", imya)
    print("-1 - вернуться назад")

    vybor = input("Введите номера через запятую: ").strip()
    if vybor == "-1":
        return

    while True:
        try:
            nomera = [int(nomer.strip()) for nomer in vybor.split(",")]
            nomera = list(dict.fromkeys(nomera))
            if nomera and all(1 <= nomer <= len(dostupnie) for nomer in nomera):
                break
        except ValueError:
            pass

        vybor = input("Нет таких номеров. Введите еще раз: ").strip()
        if vybor == "-1":
            return

    for nomer in nomera:
        imya = dostupnie[nomer - 1]
        try:
            if imya.lower().endswith(".csv"):
                df = pd.read_csv(imya, sep=";", encoding="utf-8-sig")
            else:
                df = pd.read_excel(imya)
        except Exception as oshibka:
            print("Ошибка чтения файла", imya + ":", oshibka)
            continue
        tablicy.append(Tablica(imya, df))
        print("Загружено:", imya)


def technicheskie_voprosy(tablicy, papka_mesaca, papka_goda):
    while True:
        print("\nТехнические вопросы")
        print("1 - переименовать файл")
        print("2 - дозагрузить таблицы")
        print("-1 - вернуться в главное меню")

        punkt = input("Введите номер пункта: ").strip()

        if punkt == "1":
            pereimenovat_fayl(tablicy)
        elif punkt == "2":
            dozagruzit_tablicy(tablicy, papka_mesaca, papka_goda)
        elif punkt == "-1":
            return
        else:
            print("Такого пункта пока нет")
