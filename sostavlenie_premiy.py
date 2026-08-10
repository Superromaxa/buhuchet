import os

import pandas as pd

from tablica import Tablica


def normalizovat_tekst(znachenie):
    if pd.isna(znachenie):
        return ""
    return " ".join(str(znachenie).strip().lower().replace("ё", "е").split())


def vybrat_tablicy(tablicy):
    dostupnye = [
        tablica
        for tablica in tablicy
        if "Исполнитель" in tablica.df.columns
        and "Тип операции" in tablica.df.columns
    ]

    if not dostupnye:
        print("В текущей сессии нет таблиц с исполнителями и типами операций")
        return None

    print("\nВыберите таблицы для составления премии:")
    for nomer, tablica in enumerate(dostupnye, start=1):
        print(nomer, "-", tablica.imya)
    print("Можно выбрать несколько таблиц, например: 1, 2, 4")
    print("-1 - вернуться в главное меню")

    vybor = input("Введите номера таблиц через запятую: ").strip()
    while True:
        if vybor == "-1":
            return None
        try:
            nomera = [int(nomer.strip()) for nomer in vybor.split(",")]
            nomera = list(dict.fromkeys(nomera))
            if nomera and all(1 <= nomer <= len(dostupnye) for nomer in nomera):
                return [dostupnye[nomer - 1] for nomer in nomera]
        except ValueError:
            pass
        vybor = input("Нет таких номеров. Введите еще раз: ").strip()


def podgotovit_chislovuyu_kolonku(df, kolonka):
    if kolonka not in df.columns:
        df[kolonka] = 0.0
    else:
        df[kolonka] = pd.to_numeric(df[kolonka], errors="coerce").fillna(0.0)


def rasschitat_premii_sotrudnikov(operacii):
    df = operacii.copy()
    for kolonka in ["Итоговая сумма", "Страховка", "Затраты"]:
        podgotovit_chislovuyu_kolonku(df, kolonka)

    if "Конкурс" not in df.columns:
        df["Конкурс"] = ""
    if "Наименование" not in df.columns:
        df["Наименование"] = ""
    if "Номер" not in df.columns:
        df["Номер"] = ""

    kolonka_daty = None
    for variant in ["Дата операции", "Дата"]:
        if variant in df.columns:
            kolonka_daty = variant
            break
    if kolonka_daty is None:
        df["Дата для премии"] = pd.NaT
    else:
        df["Дата для премии"] = pd.to_datetime(
            df[kolonka_daty], dayfirst=True, errors="coerce"
        ).dt.date

    df["Конкурс для расчета"] = df["Конкурс"].apply(normalizovat_tekst)
    df["Наименование для расчета"] = df["Наименование"].apply(
        normalizovat_tekst
    )
    df["Номер для расчета"] = df["Номер"].apply(normalizovat_tekst)
    df["Группа для расчета"] = df["Номер для расчета"].where(
        df["Номер для расчета"] != "",
        df["Наименование для расчета"],
    )

    rezultat = []
    for chelovek in ["Алексей", "Владимир", "Дмитрий"]:
        operacii_cheloveka = df[
            df["Исполнитель"].apply(normalizovat_tekst)
            == normalizovat_tekst(chelovek)
        ]

        bestorgovka = operacii_cheloveka[
            operacii_cheloveka["Конкурс для расчета"].isin(
                {"нет", "no", "0", "false"}
            )
        ]
        summa_bt = bestorgovka["Итоговая сумма"].sum() * 0.03

        konkurs = operacii_cheloveka[
            operacii_cheloveka["Конкурс для расчета"].isin(
                {"да", "yes", "1", "true"}
            )
        ]
        gruppy = konkurs.groupby(
            ["Дата для премии", "Группа для расчета"],
            dropna=False,
        )[["Итоговая сумма", "Затраты", "Страховка"]].sum()
        summa_konkurs = (
            gruppy["Итоговая сумма"]
            - gruppy["Затраты"]
            - gruppy["Страховка"]
        ).sum() * 0.15

        rezultat.extend([
            {
                "Имя": chelovek,
                "Тип премии": "б/т",
                "Сумма": round(float(summa_bt), 2),
            },
            {
                "Имя": chelovek,
                "Тип премии": "конкурс",
                "Сумма": round(float(summa_konkurs), 2),
            },
        ])

    return pd.DataFrame(rezultat, columns=["Имя", "Тип премии", "Сумма"])


def sostavit_premii(tablicy, papka_rezultatov=".", mesyac=""):
    vybrannye_tablicy = vybrat_tablicy(tablicy)
    if vybrannye_tablicy is None:
        return []

    operacii_dlya_premii = []
    for tablica in vybrannye_tablicy:
        tip = tablica.df["Тип операции"].apply(normalizovat_tekst)
        operacii_dlya_premii.append(
            tablica.df[tip.isin({"товар", "пришло"})].copy()
        )

    obshaya_tablica = pd.concat(operacii_dlya_premii, ignore_index=True)

    obyazatelnye = ["Итоговая сумма", "Конкурс", "Наименование"]
    net_kolonok = [
        kolonka for kolonka in obyazatelnye if kolonka not in obshaya_tablica.columns
    ]
    if "Дата операции" not in obshaya_tablica.columns and "Дата" not in obshaya_tablica.columns:
        net_kolonok.append("Дата операции")
    if net_kolonok:
        print(
            "Ошибка: для расчета премии не хватает колонок:",
            ", ".join(net_kolonok),
        )
        return []

    premii_sotrudnikov = rasschitat_premii_sotrudnikov(obshaya_tablica)

    bezopasnyi_mesyac = normalizovat_tekst(mesyac).replace(" ", "_")
    imya_obshaya = os.path.join(
        papka_rezultatov, f"премия_{bezopasnyi_mesyac}.xlsx"
    )
    imya_sotrudniki = os.path.join(
        papka_rezultatov, "премии_сотрудников.xlsx"
    )

    obshaya_tablica.to_excel(imya_obshaya, index=False)
    premii_sotrudnikov.to_excel(imya_sotrudniki, index=False)

    print("\nОбщая таблица премии сохранена:", imya_obshaya)
    print("Премии сотрудников сохранены:", imya_sotrudniki)
    print("\nПремии сотрудников:")
    print(premii_sotrudnikov.to_string(index=False))

    return [
        Tablica(imya_obshaya, obshaya_tablica),
        Tablica(imya_sotrudniki, premii_sotrudnikov),
    ]
