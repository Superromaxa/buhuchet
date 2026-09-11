import os
import sys
import traceback

import pandas as pd

from dobavit_ispolnitelya import nomer_mesaca, prochitat_podgotovlennye_operacii
from itogovaya_summa import dobavit_itogovuyu_summu
from obrabotka_altegra import dobavit_tip_operacii, prochitat_vypisku
from operacii_ispolnitelya import sohranit_excel
from sformirovat_itogovuyu_tablicu import (
    KOLONKI_ITOGA, nazvanie_firmy, normalizovat_tekst as normalizovat_itog,
    oformit_itogovyi_excel, proverit_tablicu, rasschitat_firmu,
)
from sostavlenie_premiy import rasschitat_premii_sotrudnikov
from tablica import Tablica
from tablica_vseh_operaciy import oformit_tablicu_vseh_operaciy, sostavit_avtomaticheski


# Исходные файлы лежат рядом с программой. Расширение может быть .xls или .xlsx.
FIRMY = {
    "Альтэгра": ("Альтэгра", "Альтегра"), "АВК": ("АВК",),
    "Билд": ("Билд",), "Вектор": ("Вектор",), "Макрон": ("Макрон",),
    "Позитрон": ("Позитрон",), "Сити": ("Сити",), "Кит": ("Кит",),
    "Энергопоинт": ("Энергопоинт", "Энерго поинт", "ЭП"),
    "Факторион": ("Факторион",),
}
FAYLY_ISPOLNITELEY = {
    "АлексейБТ": ("Алексей", False), "АлексейК": ("Алексей", True),
    "ВладимирБТ": ("Владимир", False), "ВладимирК": ("Владимир", True),
    "ДмитрийБТ": ("Дмитрий", False), "ДмитрийК": ("Дмитрий", True),
}


def poluchit_rabochuyu_papku():
    return os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.path.dirname(os.path.abspath(__file__))


def nastroit_konsol():
    if os.name == "nt":
        os.system("color F0")
        os.system("cls")


def zapisat_oshibku(papka):
    with open(os.path.join(papka, "error_log.txt"), "a", encoding="utf-8") as zhurnal:
        zhurnal.write(traceback.format_exc() + "\n")


def nayti_fayl(papka, imena):
    """Ищет ровно один Excel-файл по ожидаемому имени без расширения."""
    ozhidaemye = {imya.casefold() for imya in imena}
    kandidaty = []
    for imya in sorted(os.listdir(papka)):
        put = os.path.join(papka, imya)
        osnova, rasshirenie = os.path.splitext(imya)
        if os.path.isfile(put) and rasshirenie.lower() in {".xls", ".xlsx"} and osnova.casefold() in ozhidaemye:
            kandidaty.append(put)
    if len(kandidaty) > 1:
        raise ValueError("найдено несколько файлов: " + ", ".join(kandidaty))
    return kandidaty[0] if kandidaty else None


def proverit_vhodnye_fayly(papka):
    naydennye_firmy, naydennye_ispolniteli, oshibki = {}, {}, []
    for firma, varianty in FIRMY.items():
        try:
            fil = nayti_fayl(papka, varianty)
        except ValueError as oshibka:
            oshibki.append(f"{firma}: {oshibka}")
            continue
        if fil is None:
            oshibki.append(f"{firma}.xls или {firma}.xlsx")
        else:
            naydennye_firmy[firma] = fil
    for imya in FAYLY_ISPOLNITELEY:
        try:
            fil = nayti_fayl(papka, (imya,))
        except ValueError as oshibka:
            oshibki.append(f"{imya}: {oshibka}")
            continue
        if fil is None:
            oshibki.append(f"{imya}.xls или {imya}.xlsx")
        else:
            naydennye_ispolniteli[imya] = fil
    if oshibki:
        print("\nНе найдены или названы неоднозначно следующие входные файлы:")
        for oshibka in oshibki:
            print("-", oshibka)
        print("Положите их рядом с Buhuchet.exe и назовите как указано.")
        return None
    print("\nПроверка входных файлов пройдена.")
    print("Найдены выписки компаний:", ", ".join(naydennye_firmy))
    print("Найдены таблицы исполнителей:", ", ".join(naydennye_ispolniteli))
    return naydennye_firmy, naydennye_ispolniteli


def prochitat_vypiski(fayly_firm):
    vypiski = {}
    for firma, fil in fayly_firm.items():
        try:
            vypiski[firma] = Tablica(fil, prochitat_vypisku(fil))
            print("Прочитана выписка:", os.path.basename(fil))
        except Exception as oshibka:
            print(f"Ошибка чтения выписки {firma}: {oshibka}")
    return vypiski


def obrabotat_tipy(vypiski, papka_mesaca, rabochaya_papka):
    rezultaty = {}
    print("\nОбрабатываю типы операций во всех выписках.")
    for firma, vypiska in vypiski.items():
        try:
            _, rezultat, _ = dobavit_tip_operacii(vypiska, papka_mesaca, podtverzhdat_summy=False)
            if rezultat is not None:
                rezultaty[firma] = rezultat
        except Exception as oshibka:
            print(f"Ошибка при определении типов для {firma}: {oshibka}")
            zapisat_oshibku(rabochaya_papka)
    return rezultaty


def dobavit_itogovye_summy(tablicy_s_tipami, papka_mesaca, rabochaya_papka):
    rezultaty = {}
    print("\nДобавляю итоговые суммы во все выписки.")
    for firma, tablica in tablicy_s_tipami.items():
        try:
            _, rezultat = dobavit_itogovuyu_summu(tablica, papka_mesaca, avtomaticheski=True)
            if rezultat is not None:
                rezultaty[firma] = rezultat
        except Exception as oshibka:
            print(f"Ошибка при расчёте итоговой суммы для {firma}: {oshibka}")
            zapisat_oshibku(rabochaya_papka)
    return rezultaty


def prochitat_tablicy_ispolniteley(fayly):
    return [
        prochitat_podgotovlennye_operacii(fayly[imya], chelovek, konkurs)
        for imya, (chelovek, konkurs) in FAYLY_ISPOLNITELEY.items()
    ]


def zagruzit_gotovye_tablicy(papka_mesaca):
    """Загружает проверенные вручную таблицы, не изменяя их содержимое."""
    gotovye, net_faylov = {}, []
    for firma in FIRMY:
        imya = os.path.join(papka_mesaca, f"{firma}_готовая.xlsx")
        if not os.path.isfile(imya):
            net_faylov.append(imya)
            continue
        try:
            gotovye[firma] = Tablica(imya, pd.read_excel(imya))
        except Exception as oshibka:
            print(f"Ошибка чтения готовой таблицы {firma}: {oshibka}")
            return None
    if net_faylov:
        return None
    print("\nЗагружены готовые таблицы после ручной проверки.")
    return gotovye


def ne_raspoznannye_prishlo(tablica):
    df = tablica.df
    tip = df["Тип операции"].fillna("").astype(str).str.strip().str.lower()
    ispolnitel = df["Исполнитель"].fillna("").astype(str).str.strip()
    return df[(tip == "пришло") & (ispolnitel == "")].copy()


def dobavit_ispolniteley_avtomaticheski(tablicy, operacii, papka_mesaca, god, mesyac, rabochaya_papka):
    from dobavit_ispolnitelya import dobavit_ispolnitelya

    rezultaty, vse_ne_raspoznany = {}, []
    print("\nСопоставляю поступления с таблицами исполнителей.")
    for firma, tablica in tablicy.items():
        try:
            imya_gotovoy = os.path.join(papka_mesaca, f"{firma}_готовая.xlsx")
            _, rezultat = dobavit_ispolnitelya(
                tablica,
                papka_mesaca,
                god,
                mesyac,
                gotovye_operacii=operacii,
                imya_rezultata=imya_gotovoy,
            )
            if rezultat is None:
                continue
            rezultaty[firma] = rezultat
            ne_raspoznano = ne_raspoznannye_prishlo(rezultat)
            if not ne_raspoznano.empty:
                ne_raspoznano.insert(0, "Компания", firma)
                vse_ne_raspoznany.append(ne_raspoznano)
                print(f"{firma}: не распознано поступлений — {len(ne_raspoznano)}")
        except Exception as oshibka:
            print(f"Ошибка при добавлении исполнителей для {firma}: {oshibka}")
            zapisat_oshibku(rabochaya_papka)
    if vse_ne_raspoznany:
        imya = os.path.join(papka_mesaca, "Нераспознанные поступления.xlsx")
        pd.concat(vse_ne_raspoznany, ignore_index=True).to_excel(imya, index=False)
        print("\nЕсть нераспознанные поступления. Файл для проверки:", imya)
    else:
        print("\nВсе операции 'пришло' распознаны.")
    return rezultaty


def klyuchi_strok(df, kolonki):
    return df.reindex(columns=kolonki).fillna("").astype(str).agg("\x1f".join, axis=1)


def sobrat_godovye_tablicy(tablicy, papka_goda, god):
    """Дописывает только новые строки и сохраняет прежние отметки."""
    for chelovek in ("Алексей", "Владимир", "Дмитрий"):
        stroki = []
        for tablica in tablicy.values():
            maska = tablica.df["Исполнитель"].fillna("").astype(str).str.strip().str.lower() == chelovek.lower()
            stroki.append(tablica.df[maska].copy())
        novye = pd.concat(stroki, ignore_index=True)
        novye["проверено"], novye["отслежено"] = "нет", "нет"
        imya = os.path.join(papka_goda, f"{chelovek}_{god}.xlsx")
        staraya = pd.read_excel(imya) if os.path.isfile(imya) else pd.DataFrame()
        for kolonka in ("проверено", "отслежено"):
            if kolonka not in staraya.columns:
                staraya[kolonka] = "нет"
            staraya[kolonka] = staraya[kolonka].fillna("").replace("", "нет")
        kolonki_dannykh = [k for k in novye.columns if k not in {"проверено", "отслежено"}]
        starye_klyuchi = set(klyuchi_strok(staraya, kolonki_dannykh))
        dobavit = novye[~klyuchi_strok(novye, kolonki_dannykh).isin(starye_klyuchi)].copy()
        itog = pd.concat([staraya, dobavit], ignore_index=True, sort=False)
        ostalnye = [k for k in itog.columns if k not in {"проверено", "отслежено"}]
        rezultat = Tablica(imya, itog[ostalnye + ["проверено", "отслежено"]])
        sohranit_excel(rezultat)
        print(f"{chelovek}: добавлено строк — {len(dobavit)}, файл — {rezultat.imya}")


def sostavit_premii_avtomaticheski(tablicy, papka_mesaca, mesyac):
    stroki = []
    for tablica in tablicy.values():
        tipy = tablica.df["Тип операции"].apply(normalizovat_itog)
        stroki.append(tablica.df[tipy.isin({"товар", "пришло"})].copy())
    operacii = pd.concat(stroki, ignore_index=True)
    premii = rasschitat_premii_sotrudnikov(operacii)
    nazvanie = mesyac.strip().lower().replace(" ", "_")
    imya_operacii = os.path.join(papka_mesaca, f"Операции для премий {nazvanie}.xlsx")
    imya_premii = os.path.join(papka_mesaca, f"Премии сотрудников {nazvanie}.xlsx")
    operacii.to_excel(imya_operacii, index=False)
    premii.to_excel(imya_premii, index=False)
    print("Операции для премий:", imya_operacii)
    print("Премии сотрудников:", imya_premii)


def sformirovat_itog_avtomaticheski(tablicy, papka_mesaca, papka_goda, god, mesyac_chislom):
    oshibki = [f"{firma}: {oshibka}" for firma, tablica in tablicy.items() for oshibka in proverit_tablicu(tablica)]
    if oshibki:
        print("\nИтоговая таблица не сформирована:")
        for oshibka in oshibki:
            print("-", oshibka)
        return
    data, stroki, neuchtennye = pd.Timestamp(year=int(god), month=mesyac_chislom, day=1), [], []
    for tablica in tablicy.values():
        stroki_firmy, neuchtennye_firmy = rasschitat_firmu(tablica, data)
        stroki.extend(stroki_firmy)
        neuchtennye.extend(neuchtennye_firmy)
    novye = pd.DataFrame(stroki, columns=KOLONKI_ITOGA)
    imya_itoga = os.path.join(papka_goda, f"Итоговая таблица {god}.xlsx")
    if os.path.isfile(imya_itoga):
        staraya = pd.read_excel(imya_itoga)
        daty = pd.to_datetime(staraya.get("Дата"), errors="coerce")
        firmy = {normalizovat_itog(nazvanie_firmy(t.imya)) for t in tablicy.values()}
        staraya_firmy = staraya.get("Фирма", pd.Series("", index=staraya.index)).apply(normalizovat_itog)
        ostavit = ~((daty.dt.year == int(god)) & (daty.dt.month == mesyac_chislom) & staraya_firmy.isin(firmy))
        novye = pd.concat([staraya[ostavit], novye], ignore_index=True)
    novye.reindex(columns=KOLONKI_ITOGA).to_excel(imya_itoga, index=False)
    oformit_itogovyi_excel(imya_itoga)
    print("Итоговая таблица:", imya_itoga)
    if neuchtennye:
        imya = os.path.join(papka_mesaca, "Неучтенные операции.xlsx")
        pd.DataFrame(neuchtennye).to_excel(imya, index=False)
        oformit_itogovyi_excel(imya, neuchtennye=True)
        print("Неучтённые операции:", imya)


def sformirovat_vse_operacii(tablicy, papka_mesaca):
    df, propusheno = sostavit_avtomaticheski(list(tablicy.values()))
    imya = os.path.join(papka_mesaca, "Все операции.xlsx")
    df.to_excel(imya, index=False)
    oformit_tablicu_vseh_operaciy(imya)
    print("Таблица всех операций:", imya)
    if propusheno:
        print("Пропущено строк с неверной суммой:", propusheno)


def pokazat_menu():
    print("\nВыберите пункт меню")
    print("1 - определить типы операций во всех выписках")
    print("2 - добавить итоговые суммы во все выписки")
    print("3 - добавить исполнителей во все выписки")
    print("4 - собрать годовые таблицы исполнителей")
    print("5 - составить премии")
    print("6 - сформировать итоговую таблицу")
    print("7 - сформировать таблицу всех операций")
    print("8 - повторно проверить входные файлы")
    print("-1 - завершить работу")


def main():
    rabochaya_papka = poluchit_rabochuyu_papku()
    os.chdir(rabochaya_papka)
    nastroit_konsol()
    god = input("Введите год работы: ").strip()
    while not god.isdigit():
        god = input("Год нужно ввести цифрами. Введите ещё раз: ").strip()
    mesyac = input("Введите месяц работы: ").strip().lower()
    while nomer_mesaca(mesyac) is None:
        mesyac = input("Введите номер или название месяца ещё раз: ").strip().lower()
    mesyac_chislom = nomer_mesaca(mesyac)
    papka_goda = os.path.join(rabochaya_papka, god)
    papka_mesaca = os.path.join(papka_goda, mesyac)
    os.makedirs(papka_mesaca, exist_ok=True)
    print("Рабочая папка:", papka_mesaca)
    vhodnye = None
    while vhodnye is None:
        vhodnye = proverit_vhodnye_fayly(rabochaya_papka)
        if vhodnye is not None:
            break
        if input("Добавьте файлы и введите 1 для повторной проверки (-1 — выход): ").strip() == "-1":
            return
    fayly_firm, fayly_ispolniteley = vhodnye
    sostoyanie = {"vypiski": None, "tipy": None, "summy": None, "ispolniteli": None}

    def obespetchit_vypiski():
        if sostoyanie["vypiski"] is None:
            sostoyanie["vypiski"] = prochitat_vypiski(fayly_firm)
        return sostoyanie["vypiski"]

    def obespetchit_tipy():
        if sostoyanie["tipy"] is None:
            sostoyanie["tipy"] = obrabotat_tipy(obespetchit_vypiski(), papka_mesaca, rabochaya_papka)
        return sostoyanie["tipy"]

    def obespetchit_summy():
        if sostoyanie["summy"] is None:
            sostoyanie["summy"] = dobavit_itogovye_summy(obespetchit_tipy(), papka_mesaca, rabochaya_papka)
        return sostoyanie["summy"]

    def obespetchit_ispolniteley(ispolzovat_gotovye=True):
        if sostoyanie["ispolniteli"] is None:
            if ispolzovat_gotovye:
                gotovye = zagruzit_gotovye_tablicy(papka_mesaca)
                if gotovye is not None:
                    sostoyanie["ispolniteli"] = gotovye
                    return gotovye
            try:
                operacii = prochitat_tablicy_ispolniteley(fayly_ispolniteley)
            except Exception as oshibka:
                print("Ошибка чтения таблиц исполнителей:", oshibka)
                zapisat_oshibku(rabochaya_papka)
                return {}
            sostoyanie["ispolniteli"] = dobavit_ispolniteley_avtomaticheski(obespetchit_summy(), operacii, papka_mesaca, god, mesyac, rabochaya_papka)
        return sostoyanie["ispolniteli"]

    while True:
        pokazat_menu()
        punkt = input("Введите номер пункта: ").strip()
        try:
            if punkt == "1":
                obespetchit_tipy()
            elif punkt == "2":
                obespetchit_summy()
            elif punkt == "3":
                # Пункт 3 всегда создаёт новые готовые таблицы из исходных файлов.
                sostoyanie["ispolniteli"] = None
                obespetchit_ispolniteley(ispolzovat_gotovye=False)
            elif punkt == "4":
                sobrat_godovye_tablicy(obespetchit_ispolniteley(), papka_goda, god)
            elif punkt == "5":
                sostavit_premii_avtomaticheski(obespetchit_ispolniteley(), papka_mesaca, mesyac)
            elif punkt == "6":
                sformirovat_itog_avtomaticheski(obespetchit_ispolniteley(), papka_mesaca, papka_goda, god, mesyac_chislom)
            elif punkt == "7":
                sformirovat_vse_operacii(obespetchit_ispolniteley(), papka_mesaca)
            elif punkt == "8":
                vhodnye = proverit_vhodnye_fayly(rabochaya_papka)
                if vhodnye is not None:
                    fayly_firm, fayly_ispolniteley = vhodnye
                    sostoyanie = {"vypiski": None, "tipy": None, "summy": None, "ispolniteli": None}
            elif punkt == "-1":
                print("Работа программы завершена")
                return
            else:
                print("Такого пункта нет")
        except Exception as oshibka:
            print("Ошибка. Программа продолжит работу:", oshibka)
            zapisat_oshibku(rabochaya_papka)


if __name__ == "__main__":
    main()
