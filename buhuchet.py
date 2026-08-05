import os
import sys

import pandas as pd

from obrabotka_altegra import dobavit_tip_operacii
from dobavit_ispolnitelya import dobavit_ispolnitelya
from sostavlenie_premiy import sostavit_premii
from operacii_ispolnitelya import sobrat_operacii_ispolnitelya
from tablica import Tablica
from technicheskie_voprosy import technicheskie_voprosy


def poluchit_rabochuyu_papku():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def vybrat_tablicu(tablicy):
    print("\nВыберите источник данных:")
    for nomer, tablica in enumerate(tablicy, start=1):
        print(nomer, "-", tablica.imya)

    nomer_novogo = len(tablicy) + 1
    print(nomer_novogo, "- взять данные из нового файла")
    print("-1 - вернуться в главное меню")

    vybor = input("Введите номер: ")
    while vybor != "-1" and (
        not vybor.isdigit() or not 1 <= int(vybor) <= nomer_novogo
    ):
        vybor = input("Нет такого номера. Введите еще раз: ")

    if vybor == "-1":
        return "назад"
    if int(vybor) == nomer_novogo:
        return None
    return tablicy[int(vybor) - 1]


def sohranit_tablicy(tablicy):
    for tablica in tablicy:
        papka = os.path.dirname(tablica.imya)
        if papka:
            os.makedirs(papka, exist_ok=True)

        rasshirenie = os.path.splitext(tablica.imya)[1].lower()
        if rasshirenie == ".csv":
            tablica.df.to_csv(
                tablica.imya, index=False, sep=";", encoding="utf-8-sig"
            )
        else:
            if rasshirenie == ".xls":
                tablica.imya = os.path.splitext(tablica.imya)[0] + ".xlsx"
            tablica.df.to_excel(tablica.imya, index=False)
        print("Сохранено:", tablica.imya)


def zagruzit_tablicy(papka_mesaca, papka_goda):
    dostupnie_fayly = []

    for papka in [papka_mesaca, papka_goda]:
        if not os.path.isdir(papka):
            continue

        for imya in sorted(os.listdir(papka)):
            polnoe_imya = os.path.join(papka, imya)
            if not os.path.isfile(polnoe_imya):
                continue
            if os.path.splitext(imya)[1].lower() in [".xlsx", ".xls", ".csv"]:
                dostupnie_fayly.append(polnoe_imya)

    if not dostupnie_fayly:
        print("В папках года и месяца пока нет сохраненных таблиц")
        return []

    print("\nДоступные сохраненные таблицы:")
    for nomer, imya in enumerate(dostupnie_fayly, start=1):
        print(nomer, "-", imya)

    vybor = input(
        "Введите номера таблиц через запятую или нажмите Enter, чтобы не загружать: "
    ).strip()
    if not vybor:
        return []

    while True:
        try:
            nomera = [int(nomer.strip()) for nomer in vybor.split(",")]
            nomera = list(dict.fromkeys(nomera))
            if nomera and all(1 <= nomer <= len(dostupnie_fayly) for nomer in nomera):
                break
        except ValueError:
            pass
        vybor = input("Нет таких номеров. Введите еще раз: ").strip()

    tablicy = []
    for nomer in nomera:
        imya = dostupnie_fayly[nomer - 1]
        if imya.lower().endswith(".csv"):
            df = pd.read_csv(imya, sep=";", encoding="utf-8-sig")
        else:
            df = pd.read_excel(imya)
        tablicy.append(Tablica(imya, df))
        print("Загружено:", imya)

    return tablicy


rabochaya_papka = poluchit_rabochuyu_papku()
os.chdir(rabochaya_papka)

god = input("Введите год работы: ").strip()
while not god.isdigit():
    god = input("Год нужно ввести цифрами. Введите еще раз: ").strip()

mesyac = input("Введите месяц работы: ").strip().lower()
while not mesyac:
    mesyac = input("Введите месяц работы: ").strip().lower()

papka_goda = os.path.join(rabochaya_papka, god)
papka_mesaca = os.path.join(papka_goda, mesyac)
os.makedirs(papka_mesaca, exist_ok=True)

print("Рабочая папка:", papka_mesaca)

tablicy = zagruzit_tablicy(papka_mesaca, papka_goda)

while True:
    print("\nВыберите пункт меню")
    print("1 - добавить тип операции для файла")
    print("2 - добавить исполнителя для файла")
    print("3 - составление премий")
    print("4 - собрать операции по исполнителю")
    print("5 - технические вопросы")
    print("-1 - завершить работу")

    punkt = input("Введите номер пункта: ")

    if punkt == "1":
        istochnik = vybrat_tablicu(tablicy)
        if istochnik == "назад":
            continue
        novaya, rezultat, nerasp = dobavit_tip_operacii(
            istochnik, papka_mesaca
        )
        if rezultat is None:
            continue
        if novaya is not None:
            tablicy.append(novaya)
        tablicy.extend([rezultat, nerasp])
    elif punkt == "2":
        istochnik = vybrat_tablicu(tablicy)
        if istochnik == "назад":
            continue
        novaya, rezultat = dobavit_ispolnitelya(
            istochnik, papka_mesaca
        )
        if novaya is not None:
            tablicy.append(novaya)
        if rezultat is not None:
            tablicy.append(rezultat)
    elif punkt == "3":
        rezultat = sostavit_premii(tablicy, papka_mesaca)
        if rezultat is not None:
            tablicy.append(rezultat)
    elif punkt == "4":
        rezultat, dobavit_v_sessiyu = sobrat_operacii_ispolnitelya(
            tablicy, papka_goda, god
        )
        if rezultat is not None and dobavit_v_sessiyu:
            tablicy.append(rezultat)
    elif punkt == "5":
        technicheskie_voprosy(tablicy, papka_mesaca, papka_goda)
    elif punkt == "-1":
        if tablicy:
            print("\nСохраняю таблицы текущей сессии:")
            sohranit_tablicy(tablicy)
        print("Работа программы завершена")
        break
    else:
        print("Такого пункта пока нет")
