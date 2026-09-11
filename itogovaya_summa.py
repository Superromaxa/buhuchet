import os
import re

import pandas as pd

from tablica import Tablica


SHABLON_CHISLA = r"[0-9]+(?:[ \u00a0][0-9]{3})*(?:[.,][0-9]+)?"


def preobrazovat_chislo(tekst):
    return float(str(tekst).replace(" ", "").replace("\u00a0", "").replace(",", "."))


def nayti_chislo(tekst, nazvanie):
    sovpadenie = re.search(
        rf"{nazvanie}\.*\s*({SHABLON_CHISLA})",
        str(tekst),
        flags=re.IGNORECASE,
    )
    if sovpadenie is None:
        return None
    return preobrazovat_chislo(sovpadenie.group(1))


def sprosit_chislo(tekst):
    znachenie = input(tekst).strip()
    while True:
        if znachenie == "-1":
            return None
        try:
            return preobrazovat_chislo(znachenie)
        except ValueError:
            znachenie = input(
                "Введите число, например 1250,50, или -1 для возврата: "
            ).strip()


def dobavit_itogovuyu_summu(
    tablica=None, papka_rezultatov=".", avtomaticheski=False
):
    novaya_tablica = None

    if tablica is None:
        fil = input("Введите название файла (-1 — назад): ").strip()
        if fil == "-1":
            return None, None
        try:
            df = pd.read_excel(fil)
        except Exception as oshibka:
            print("Ошибка чтения файла:", oshibka)
            return None, None

        imya_ishodnoy = os.path.splitext(os.path.basename(fil))[0] + ".xlsx"
        novaya_tablica = Tablica(
            os.path.join(papka_rezultatov, imya_ishodnoy), df.copy()
        )
    else:
        fil = tablica.imya
        df = tablica.df.copy()

    obyazatelnye_kolonki = [
        "Оборот Дт",
        "Оборот Кт",
        "Назначение платежа",
    ]
    net_kolonok = [
        kolonka for kolonka in obyazatelnye_kolonki if kolonka not in df.columns
    ]
    if net_kolonok:
        print("Ошибка: в таблице нет колонок:", ", ".join(net_kolonok))
        return novaya_tablica, None

    debit = pd.to_numeric(df["Оборот Дт"], errors="coerce")
    kredit = pd.to_numeric(df["Оборот Кт"], errors="coerce")
    df["Итоговая сумма"] = debit.combine_first(kredit)
    if "Комиссия" not in df.columns:
        df["Комиссия"] = 0.0
    else:
        df["Комиссия"] = pd.to_numeric(df["Комиссия"], errors="coerce").fillna(0)
    if "НДС" not in df.columns:
        df["НДС"] = 0.0
    else:
        df["НДС"] = pd.to_numeric(df["НДС"], errors="coerce").fillna(0)

    naznacheniya = df["Назначение платежа"].fillna("").astype(str)
    vozmeshenie = naznacheniya.str.contains(
        r"возм\.*\s*по\s*дог\.*", case=False, regex=True
    )

    for index in df.index[vozmeshenie]:
        naznachenie = naznacheniya.loc[index]
        osnovnaya_summa = df.loc[index, "Итоговая сумма"]
        komissiya = nayti_chislo(naznachenie, "ком")
        nds = nayti_chislo(naznachenie, "ндс")

        if pd.isna(osnovnaya_summa) or komissiya is None or nds is None:
            print("\n----------------------------------------")
            print("Назначение:", naznachenie)

        if pd.isna(osnovnaya_summa):
            if avtomaticheski:
                raise ValueError(
                    "не найдена сумма в Дт/Кт для возмещения "
                    f"в строке Excel {index + 2}"
                )
            osnovnaya_summa = sprosit_chislo(
                "Не найдена сумма в Дт/Кт. Введите сумму (-1 — назад): "
            )
            if osnovnaya_summa is None:
                return novaya_tablica, None

        if komissiya is None:
            if avtomaticheski:
                raise ValueError(
                    "не найдена комиссия для возмещения "
                    f"в строке Excel {index + 2}"
                )
            komissiya = sprosit_chislo(
                "Не найдена комиссия после 'Ком.'. Введите комиссию (-1 — назад): "
            )
            if komissiya is None:
                return novaya_tablica, None

        if nds is None:
            if avtomaticheski:
                raise ValueError(
                    "не найден НДС для возмещения "
                    f"в строке Excel {index + 2}"
                )
            nds = sprosit_chislo(
                "Не найдено число после 'НДС'. Введите НДС (-1 — назад): "
            )
            if nds is None:
                return novaya_tablica, None

        df.loc[index, "Комиссия"] = komissiya
        df.loc[index, "НДС"] = nds
        df.loc[index, "Итоговая сумма"] = osnovnaya_summa + komissiya + nds

    imya_fayla = os.path.basename(fil)
    imya_bez_rasshireniya = os.path.splitext(imya_fayla)[0]
    fail_rezultat = os.path.join(
        papka_rezultatov,
        imya_bez_rasshireniya + "_с_итоговой_суммой.xlsx",
    )
    df.to_excel(fail_rezultat, index=False)

    print("\nГотово. Результат сохранен в файле:", fail_rezultat)
    return novaya_tablica, Tablica(fail_rezultat, df)
