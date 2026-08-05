import os

import pandas as pd
import numpy as np

from tablica import Tablica


def dobavit_tip_operacii(tablica=None, papka_rezultatov="."):
    novaya_tablica = None

    if tablica is None:
        fil = input("Введите название файла (-1 — назад): ")
        if fil.strip() == "-1":
            return None, None, None

        # Первые 9 строк пропускаем. Строка 10 становится заголовком.
        df = pd.read_excel(fil, skiprows=9)

        # Во второй строке таблицы лежат названия ИНН, КПП, Счет и БИК.
        df = df.rename(columns={
            "Unnamed: 3": "ИНН",
            "Unnamed: 4": "КПП",
            "Unnamed: 5": "Счет",
            "Unnamed: 6": "БИК",
        })
        df = df.iloc[:, :11]
        df = df.drop(0).reset_index(drop=True)

        # Строку ИТОГО и все строки после нее не берем.
        stroka_itogo = df.index[
            df["Документ"].astype(str).str.strip().str.upper() == "ИТОГО:"
        ]
        if len(stroka_itogo) > 0:
            df = df.iloc[:stroka_itogo[0]]
    else:
        fil = tablica.imya
        df = tablica.df.copy()

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
