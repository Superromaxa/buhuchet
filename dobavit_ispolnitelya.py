import os
import re
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP

import pandas as pd

from tablica import Tablica


def normalizovat_tekst(tekst):
    return " ".join(str(tekst).strip().lower().replace("ё", "е").split())


def nayti_kolonku(kolonki, obyazatelnye_slova):
    for kolonka in kolonki:
        nazvanie = normalizovat_tekst(kolonka)
        if all(slovo in nazvanie for slovo in obyazatelnye_slova):
            return kolonka
    return None


def prochitat_operacii_cheloveka(fil, chelovek):
    excel = pd.ExcelFile(fil)

    for list_excel in excel.sheet_names:
        df = pd.read_excel(fil, sheet_name=list_excel)

        kolonka_summa = nayti_kolonku(
            df.columns, ["сумма", "оплат", "заказчик"]
        )
        kolonka_data = nayti_kolonku(
            df.columns, ["дата", "оплат", "заказчик"]
        )
        kolonka_postavshik = nayti_kolonku(
            df.columns, ["фирма", "постав"]
        )
        kolonka_nomenklatura = nayti_kolonku(
            df.columns, ["номенклатур"]
        )

        naydennye = [
            kolonka_summa,
            kolonka_data,
            kolonka_postavshik,
            kolonka_nomenklatura,
        ]
        if all(kolonka is not None for kolonka in naydennye):
            rezultat = df[naydennye].copy()
            rezultat.columns = [
                "Сумма оплаты заказчиком",
                "Дата оплаты заказчиком",
                "Фирма поставщик",
                "Номенклатура",
            ]
            rezultat["Исполнитель"] = chelovek
            rezultat["Сумма оплаты заказчиком"] = pd.to_numeric(
                rezultat["Сумма оплаты заказчиком"], errors="coerce"
            )
            rezultat = rezultat.dropna(subset=["Сумма оплаты заказчиком"])
            rezultat = rezultat[
                rezultat[["Фирма поставщик", "Номенклатура"]]
                .notna()
                .any(axis=1)
            ]
            return rezultat

    raise ValueError(
        "Не найдены колонки с суммой, датой оплаты, "
        "фирмой поставщиком и номенклатурой"
    )


def sprosit_operacii_cheloveka(chelovek):
    while True:
        fil = input(
            f"Введите название файла с операциями {chelovek} (-1 — назад): "
        ).strip()
        if fil == "-1":
            return None, None

        try:
            operacii = prochitat_operacii_cheloveka(fil, chelovek)
            return fil, operacii
        except Exception as oshibka:
            print("Ошибка чтения файла:", oshibka)


def tochnaya_summa(summa):
    try:
        return Decimal(str(summa)).quantize(
            Decimal("0.01"), rounding=ROUND_HALF_UP
        )
    except (InvalidOperation, ValueError, TypeError):
        return None


def nomer_mesaca(mesyac):
    mesyacy = {
        "1": 1, "01": 1, "январь": 1, "января": 1,
        "2": 2, "02": 2, "февраль": 2, "февраля": 2,
        "3": 3, "03": 3, "март": 3, "марта": 3,
        "4": 4, "04": 4, "апрель": 4, "апреля": 4,
        "5": 5, "05": 5, "май": 5, "мая": 5,
        "6": 6, "06": 6, "июнь": 6, "июня": 6,
        "7": 7, "07": 7, "июль": 7, "июля": 7,
        "8": 8, "08": 8, "август": 8, "августа": 8,
        "9": 9, "09": 9, "сентябрь": 9, "сентября": 9,
        "10": 10, "октябрь": 10, "октября": 10,
        "11": 11, "ноябрь": 11, "ноября": 11,
        "12": 12, "декабрь": 12, "декабря": 12,
    }
    return mesyacy.get(normalizovat_tekst(mesyac))


def poluchit_daty(znachenie):
    if pd.isna(znachenie):
        return set()

    if isinstance(znachenie, pd.Timestamp):
        return {znachenie.date()}

    daty = set()
    tekst = str(znachenie)
    shablony = re.findall(
        r"\d{1,2}\.\d{1,2}\.\d{4}|\d{4}-\d{1,2}-\d{1,2}", tekst
    )
    for tekst_daty in shablony:
        if "-" in tekst_daty:
            data = pd.to_datetime(
                tekst_daty, format="%Y-%m-%d", errors="coerce"
            )
        else:
            data = pd.to_datetime(
                tekst_daty, format="%d.%m.%Y", errors="coerce"
            )
        if not pd.isna(data):
            daty.add(data.date())

    if not daty:
        data = pd.to_datetime(znachenie, dayfirst=True, errors="coerce")
        if not pd.isna(data):
            daty.add(data.date())
    return daty


def poluchit_summu_osnovnoy_tablicy(df, index):
    if "Итоговая сумма" in df.columns:
        return tochnaya_summa(df.loc[index, "Итоговая сумма"])

    summa = df.loc[index, "Оборот Дт"]
    if pd.isna(summa):
        summa = df.loc[index, "Оборот Кт"]
    return tochnaya_summa(summa)


def zapisat_odinochnoe_sovpadenie(df, index, naydeno):
    df.loc[index, "Исполнитель"] = naydeno["Исполнитель"]
    df.loc[index, "Номенклатура"] = naydeno["Номенклатура"]
    df.loc[index, "Фирма поставщик"] = naydeno["Фирма поставщик"]


def nayti_kombinacii(indexy, nuzhnaya_summa, spravochnik, ispolzovannye):
    dostupnye = []
    for index in indexy:
        if index in ispolzovannye:
            continue
        summa = spravochnik.loc[index, "Точная сумма"]
        if summa is not None and summa > 0 and summa <= nuzhnaya_summa:
            dostupnye.append((index, summa))

    dostupnye.sort(key=lambda para: para[1], reverse=True)
    rezultaty = []

    def perebor(poziciya, tekushaya_summa, vybrannye):
        if len(rezultaty) > 1:
            return
        if tekushaya_summa == nuzhnaya_summa:
            if len(vybrannye) >= 2:
                rezultaty.append(tuple(vybrannye))
            return
        if tekushaya_summa > nuzhnaya_summa:
            return

        for nomer in range(poziciya, len(dostupnye)):
            index, summa = dostupnye[nomer]
            novaya_summa = tekushaya_summa + summa
            if novaya_summa <= nuzhnaya_summa:
                perebor(nomer + 1, novaya_summa, vybrannye + [index])

    perebor(0, Decimal("0.00"), [])
    return rezultaty


def dobavit_ispolnitelya(
    tablica=None, papka_rezultatov=".", god=None, mesyac=None
):
    novye_tablicy = []

    if tablica is None:
        fil = input("Введите название файла (-1 — назад): ").strip()
        if fil == "-1":
            return [], None
        try:
            df = pd.read_excel(fil)
        except Exception as oshibka:
            print("Ошибка чтения файла:", oshibka)
            return [], None

        imya_ishodnoy = os.path.splitext(os.path.basename(fil))[0] + ".xlsx"
        novye_tablicy.append(Tablica(
            os.path.join(papka_rezultatov, imya_ishodnoy), df.copy()
        ))
    else:
        fil = tablica.imya
        df = tablica.df.copy()

    obyazatelnye_kolonki = [
        "Тип операции",
        "Назначение платежа",
        "Оборот Дт",
        "Оборот Кт",
    ]
    net_kolonok = [
        kolonka for kolonka in obyazatelnye_kolonki if kolonka not in df.columns
    ]
    if net_kolonok:
        print("Ошибка: в таблице нет колонок:", ", ".join(net_kolonok))
        return novye_tablicy, None

    if "Исполнитель" not in df.columns:
        df["Исполнитель"] = None
    df["Исполнитель"] = df["Исполнитель"].astype("object")

    if "Номенклатура" not in df.columns:
        df["Номенклатура"] = None
    if "Фирма поставщик" not in df.columns:
        df["Фирма поставщик"] = None

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

    stroki_prishlo = df.index[df["Тип операции"] == "пришло"]
    if len(stroki_prishlo) > 0:
        vse_operacii = []
        for chelovek in ["Алексей", "Владимир", "Дмитрий"]:
            _, operacii = sprosit_operacii_cheloveka(chelovek)
            if operacii is None:
                return novye_tablicy, None
            vse_operacii.append(operacii)

        spravochnik = pd.concat(vse_operacii, ignore_index=True)
        mesyac_chislom = nomer_mesaca(mesyac)
        if mesyac_chislom is None or not str(god).isdigit():
            print("Ошибка: не удалось определить рабочий год или месяц")
            return novye_tablicy, None

        spravochnik["Даты оплаты"] = spravochnik[
            "Дата оплаты заказчиком"
        ].apply(poluchit_daty)
        spravochnik = spravochnik[
            spravochnik["Даты оплаты"].apply(
                lambda daty: any(
                    data.year == int(god) and data.month == mesyac_chislom
                    for data in daty
                )
            )
        ].copy()
        spravochnik.reset_index(drop=True, inplace=True)

        spravochnik["Точная сумма"] = spravochnik[
            "Сумма оплаты заказчиком"
        ].apply(tochnaya_summa)
        spravochnik["Поставщик для сравнения"] = spravochnik[
            "Фирма поставщик"
        ].apply(normalizovat_tekst)

        ispolzovannye_stroki = set()
        neraspredelennye_stroki = []

        # Сначала распределяем все одиночные точные совпадения.
        for index in stroki_prishlo:
            summa = poluchit_summu_osnovnoy_tablicy(df, index)
            kandidaty = spravochnik[
                (spravochnik["Точная сумма"] == summa)
                & (~spravochnik.index.isin(ispolzovannye_stroki))
            ]

            if len(kandidaty) == 1:
                naydeno = kandidaty.iloc[0]
                zapisat_odinochnoe_sovpadenie(df, index, naydeno)
                ispolzovannye_stroki.add(kandidaty.index[0])
            else:
                neraspredelennye_stroki.append(index)

        # Для каждого человека, каждой даты и каждого поставщика формируем
        # отдельную группу. Люди и даты между собой никогда не смешиваются.
        gruppy = {}
        for spravochnik_index, stroka in spravochnik.iterrows():
            postavshik = stroka["Поставщик для сравнения"]
            if not postavshik:
                continue
            for data in stroka["Даты оплаты"]:
                if data.year != int(god) or data.month != mesyac_chislom:
                    continue
                kluch = (stroka["Исполнитель"], data, postavshik)
                gruppy.setdefault(kluch, []).append(spravochnik_index)

        # Затем для оставшихся операций перебираем комбинации из двух,
        # трех и большего числа еще не использованных покупок.
        vse_eshe_ne_naydennye = []

        for index in neraspredelennye_stroki:
            nuzhnaya_summa = poluchit_summu_osnovnoy_tablicy(df, index)
            podhodyashie_kombinacii = []

            if nuzhnaya_summa is not None:
                for kluch, indexy_gruppy in gruppy.items():
                    kombinacii = nayti_kombinacii(
                        indexy_gruppy,
                        nuzhnaya_summa,
                        spravochnik,
                        ispolzovannye_stroki,
                    )
                    for kombinaciya in kombinacii:
                        podhodyashie_kombinacii.append((kluch, kombinaciya))
                        if len(podhodyashie_kombinacii) > 1:
                            break
                    if len(podhodyashie_kombinacii) > 1:
                        break

            if len(podhodyashie_kombinacii) == 1:
                kluch, kombinaciya = podhodyashie_kombinacii[0]
                chelovek, _, _ = kluch
                pervaya = spravochnik.loc[kombinaciya[0]]

                nomenklatura = None
                for spravochnik_index in kombinaciya:
                    tekushaya = spravochnik.loc[spravochnik_index, "Номенклатура"]
                    if not pd.isna(tekushaya) and str(tekushaya).strip():
                        nomenklatura = tekushaya
                        break

                df.loc[index, "Исполнитель"] = chelovek
                df.loc[index, "Номенклатура"] = nomenklatura
                df.loc[index, "Фирма поставщик"] = pervaya["Фирма поставщик"]
                ispolzovannye_stroki.update(kombinaciya)
            else:
                vse_eshe_ne_naydennye.append(index)

        print(
            "\nОпераций 'пришло', которые нужно проверить вручную:",
            len(vse_eshe_ne_naydennye),
        )

    imya_fayla = os.path.basename(fil)
    imya_bez_rasshireniya = os.path.splitext(imya_fayla)[0]
    fail_rezultat = os.path.join(
        papka_rezultatov,
        imya_bez_rasshireniya + "_с_исполнителями.xlsx",
    )
    df.to_excel(fail_rezultat, index=False)

    print("\nГотово. Результат сохранен в файле:", fail_rezultat)
    return novye_tablicy, Tablica(fail_rezultat, df)
