import os

import pandas as pd
import numpy as np

from tablica import Tablica


EDINYE_KOLONKI = [
    "Документ",
    "Дата операции",
    "Корреспондент",
    "ИНН",
    "КПП",
    "Счет",
    "БИК",
    "Наименование банка",
    "Вх.остаток",
    "Оборот Дт",
    "Оборот Кт",
    "Назначение платежа",
]


def normalizovat_nazvanie(znachenie):
    if pd.isna(znachenie):
        return ""
    return str(znachenie).strip().lower().replace("ё", "е")


def prochitat_vypisku(fil, list_excel=0):
    syrye_dannye = pd.read_excel(fil, sheet_name=list_excel, header=None)

    stroka_zagolovka = None
    varianty_documenta = {"документ", "номер документа"}
    varianty_daty = {"дата", "дата операции"}

    # Ищем строку заголовка по двум обязательным названиям, а не по номеру строки.
    for index, stroka in syrye_dannye.head(50).iterrows():
        znacheniya = {normalizovat_nazvanie(znachenie) for znachenie in stroka}
        est_document = bool(znacheniya & varianty_documenta)
        est_data = bool(znacheniya & varianty_daty)

        if est_document and est_data:
            stroka_zagolovka = index
            break

    if stroka_zagolovka is None:
        raise ValueError(
            "Не найдена строка заголовка с колонками "
            "'Документ/Номер документа' и 'Дата/Дата операции'"
        )

    # Заголовок занимает две строки. В первой лежат основные названия,
    # во второй — Наименование, ИНН, КПП, Счет и БИК.
    verhnie_nazvaniya = syrye_dannye.iloc[stroka_zagolovka]
    nizhnie_nazvaniya = syrye_dannye.iloc[stroka_zagolovka + 1]
    nazvaniya_kolonok = []

    for nomer, (verhnee, nizhnee) in enumerate(
        zip(verhnie_nazvaniya, nizhnie_nazvaniya), start=1
    ):
        if normalizovat_nazvanie(verhnee):
            nazvanie = str(verhnee).strip()
        elif normalizovat_nazvanie(nizhnee):
            nazvanie = str(nizhnee).strip()
        else:
            nazvanie = f"Колонка {nomer}"
        nazvaniya_kolonok.append(nazvanie)

    df = syrye_dannye.iloc[stroka_zagolovka + 2:].copy()
    df.columns = nazvaniya_kolonok
    # Приводим два банковских формата к одним названиям колонок.
    df = df.rename(columns={
        "Номер документа": "Документ",
        "Дата": "Дата операции",
        "Дебет": "Оборот Дт",
        "Кредит": "Оборот Кт",
        "Контрагент": "Корреспондент",
        "Счёт": "Счет",
    })

    # Для всех банков используем одинаковые колонки и одинаковый порядок.
    # Если в банковском формате колонки нет, она останется пустой.
    df = df.reindex(columns=EDINYE_KOLONKI)

    return df


def dobavit_tip_operacii(tablica=None, papka_rezultatov="."):
    novaya_tablica = None

    if tablica is None:
        fil = input("Введите название файла (-1 — назад): ")
        if fil.strip() == "-1":
            return None, None, None

        try:
            df = prochitat_vypisku(fil)
        except Exception as oshibka:
            print("Ошибка чтения файла:", oshibka)
            return None, None, None

        # Строку ИТОГО и все строки после нее не берем.
        stroka_itogo = df.index[
            df["Документ"].astype(str).str.strip().str.upper() == "ИТОГО:"
        ]
        if len(stroka_itogo) > 0:
            df = df.iloc[:stroka_itogo[0]]
    else:
        fil = tablica.imya
        df = tablica.df.copy()

    obyazatelnye_kolonki = [
        "Документ",
        "Оборот Дт",
        "Оборот Кт",
        "Назначение платежа",
    ]
    net_kolonok = [
        kolonka for kolonka in obyazatelnye_kolonki if kolonka not in df.columns
    ]
    if net_kolonok:
        print("Ошибка: в таблице нет колонок:", ", ".join(net_kolonok))
        return None, None, None

    df = df.dropna(how="all").reset_index(drop=True)
    df["Оборот Дт"] = pd.to_numeric(df["Оборот Дт"], errors="coerce")
    df["Оборот Кт"] = pd.to_numeric(df["Оборот Кт"], errors="coerce")

    if tablica is None:
        # Сохраняем отдельно исходную таблицу, подготовленную для обработки.
        imya_ishodnoy = os.path.splitext(os.path.basename(fil))[0] + ".xlsx"
        novaya_tablica = Tablica(
            os.path.join(papka_rezultatov, imya_ishodnoy), df.copy()
        )

    naznachenie = df["Назначение платежа"].fillna("").astype(str)

    conditions = [
        naznachenie.str.contains(r"Оплата по дог|Возм\. по дог\.", case=False),
        naznachenie.str.contains(r"Оплата за тех\. обслуживание", case=False),
        naznachenie.str.contains("Оплата", case=False),
        naznachenie.str.contains("Выплата процентов согласно депозитного договора", case=False),
        naznachenie.str.contains("Пополнение счета согласно депозитного договора", case=False),
        naznachenie.str.contains("Возврат депозитн|Возврат согласно депозитн", case=False),
        naznachenie.str.contains("аренд", case=False),
        naznachenie.str.contains("Единый налоговый платеж|Единый социальный налог|Страховые взносы|пени", case=False),
        naznachenie.str.contains("заработной платы|заработная плата", case=False),
        naznachenie.str.contains("Комиссия", case=False),
        naznachenie.str.contains("РАД|Плата оператору", case=False),
        naznachenie.str.contains("Займ|займ", case=False),
        naznachenie.str.contains("Выдача денежных средств", case=False),
        naznachenie.str.contains("Перевод собственных средств", case=False),
        naznachenie.str.contains("Возврат средств|Возврат денежных средств", case=False),
    ]

    tipy = {
        1: "пришло",
        2: "товар",
        3: "% депозит",
        4: "депозит",
        5: "депозит возврат",
        6: "аренда",
        7: "налог",
        8: "ЗП",
        9: "банк комиссия",
        10: "РАД",
        11: "займ",
        12: "ПО",
        13: "снятие дс",
        14: "перевод сс",
        15: "возврат средств"
    }

    # Типы здесь идут в том же порядке, что и автоматические условия выше.
    avtomaticheskie_tipy = [
        "пришло",
        "ПО",
        "товар",
        "% депозит",
        "депозит",
        "депозит возврат",
        "аренда",
        "налог",
        "ЗП",
        "банк комиссия",
        "РАД",
        "займ",
        "снятие дс",
        "перевод сс",
        "возврат средств",
    ]

    df["Тип операции"] = np.select(
        conditions,
        avtomaticheskie_tipy,
        default="другое",
    )

    # Товар не может быть приходом.
    df.loc[
        (df["Тип операции"] == "товар")
        & (df["Оборот Кт"].fillna(0) != 0),
        "Тип операции",
    ] = "пришло"

    # Сначала сохраняем нераспознанные операции, до ручного распределения.
    nerasp = df[df["Тип операции"] == "другое"]
    imya_fayla = os.path.basename(fil)
    imya_bez_rasshireniya = os.path.splitext(imya_fayla)[0]
    fail_nerasp = os.path.join(
        papka_rezultatov, imya_bez_rasshireniya + "_нераспознанные.xlsx"
    )
    nerasp[["Оборот Дт", "Оборот Кт", "Назначение платежа"]].to_excel(
        fail_nerasp, index=False
    )

    print("\nТипы операций:")
    for nomer, tip in tipy.items():
        print(nomer, "-", tip)
    print("\nСохрани таблицу или сфоткай")
    print("Нераспознанные операции сохранены в файле:", fail_nerasp)

    # Ручное распределение оставшихся операций.
    for index in df.index[df["Тип операции"] == "другое"]:
        print("\n----------------------------------------")
        print("Дт:", df.loc[index, "Оборот Дт"])
        print("Кт:", df.loc[index, "Оборот Кт"])
        print("Назначение:", df.loc[index, "Назначение платежа"])

        nomer = input("Введите номер типа операции: ")
        while not nomer.isdigit() or int(nomer) not in tipy:
            nomer = input("Нет такого номера. Введите еще раз: ")

        df.loc[index, "Тип операции"] = tipy[int(nomer)]

    df["Исполнитель"] = np.nan

    fail_rezultat = os.path.join(
        papka_rezultatov, imya_bez_rasshireniya + "_с_типами.xlsx"
    )
    df.to_excel(fail_rezultat, index=False)

    print("\nГотово. Результат сохранен в файле:", fail_rezultat)

    return (
        novaya_tablica,
        Tablica(fail_rezultat, df),
        Tablica(fail_nerasp, nerasp.copy()),
    )
