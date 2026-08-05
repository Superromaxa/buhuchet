import os

import pandas as pd

from tablica import Tablica


def dobavit_ispolnitelya(tablica=None, papka_rezultatov="."):
    novaya_tablica = None

    if tablica is None:
        fil = input("Введите название файла (-1 — назад): ")
        if fil.strip() == "-1":
            return None, None
        df = pd.read_excel(fil)
        imya_ishodnoy = os.path.splitext(os.path.basename(fil))[0] + ".xlsx"
        novaya_tablica = Tablica(
            os.path.join(papka_rezultatov, imya_ishodnoy), df.copy()
        )
    else:
        fil = tablica.imya
        df = tablica.df.copy()

    if "Тип операции" not in df.columns:
        print("В файле нет столбца 'Тип операции'")
        return novaya_tablica, None

    if "Исполнитель" not in df.columns:
        df["Исполнитель"] = None
    df["Исполнитель"] = df["Исполнитель"].astype("object")

    # Исполнитель для товарных операций определяется по букве в назначении.
    for index in df.index[df["Тип операции"] == "товар"]:
        naznachenie = str(df.loc[index, "Назначение платежа"]).upper()

        if "(А)" in naznachenie:
            df.loc[index, "Исполнитель"] = "Алексей"
        elif "(В)" in naznachenie:
            df.loc[index, "Исполнитель"] = "Владимир"
        elif "(Д)" in naznachenie:
            df.loc[index, "Исполнитель"] = "Дмитрий"
        else:
            df.loc[index, "Исполнитель"] = "Андрей"

    lyudi = {
        1: "Алексей",
        2: "Владимир",
        3: "Дмитрий",
        4: "Андрей",
    }

    print("\nИсполнители:")
    for nomer, chelovek in lyudi.items():
        print(nomer, "-", chelovek)

    # Для всех приходов исполнитель вводится вручную.
    for index in df.index[df["Тип операции"] == "пришло"]:
        print("\n----------------------------------------")
        print("Дт:", df.loc[index, "Оборот Дт"])
        print("Кт:", df.loc[index, "Оборот Кт"])
        print("Назначение:", df.loc[index, "Назначение платежа"])

        nomer = input("Введите номер исполнителя: ")
        while not nomer.isdigit() or int(nomer) not in lyudi:
            nomer = input("Нет такого номера. Введите еще раз: ")

        df.loc[index, "Исполнитель"] = lyudi[int(nomer)]

    imya_fayla = os.path.basename(fil)
    imya_bez_rasshireniya = os.path.splitext(imya_fayla)[0]
    fail_rezultat = os.path.join(
        papka_rezultatov,
        imya_bez_rasshireniya + "_с_исполнителями.xlsx",
    )
    df.to_excel(fail_rezultat, index=False)

    print("\nГотово. Результат сохранен в файле:", fail_rezultat)

    return novaya_tablica, Tablica(fail_rezultat, df)
