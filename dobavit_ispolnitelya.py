import os
import re
from bisect import bisect_right
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP

import pandas as pd

from tablica import Tablica


def normalizovat_tekst(tekst):
    if pd.isna(tekst):
        return ""
    return " ".join(str(tekst).strip().lower().replace("ё", "е").split())


def vybrat_format_tablic_ispolniteley():
    print("\nКакие таблицы исполнителей использовать?")
    print("1 - подготовленные таблицы")
    print("2 - старые таблицы")
    print("-1 - вернуться в главное меню")

    vybor = input("Введите номер: ").strip()
    while vybor not in {"1", "2", "-1"}:
        vybor = input("Нет такого номера. Введите еще раз: ").strip()
    return vybor


def prochitat_staryi_istochnik(
    fil,
    list_excel,
    chelovek,
    konkurs,
    kolonka_daty,
    kolonka_naimenovaniya,
    kolonka_firmy,
    kolonka_summy,
):
    df = pd.read_excel(fil, sheet_name=list_excel)
    kolonki = {
        normalizovat_tekst(kolonka): kolonka for kolonka in df.columns
    }
    trebovaniya = {
        "Дата оплаты": kolonka_daty,
        "Наименование": kolonka_naimenovaniya,
        "Фирма": kolonka_firmy,
        "Сумма оплаты": kolonka_summy,
    }

    net_kolonok = [
        ishodnoe
        for ishodnoe in trebovaniya.values()
        if normalizovat_tekst(ishodnoe) not in kolonki
    ]
    if net_kolonok:
        raise ValueError(
            f"в листе '{list_excel}' файла '{os.path.basename(fil)}' "
            f"не найдены колонки: {', '.join(net_kolonok)}"
        )

    rezultat = pd.DataFrame({
        novoe: df[kolonki[normalizovat_tekst(ishodnoe)]]
        for novoe, ishodnoe in trebovaniya.items()
    })
    rezultat["Исполнитель"] = chelovek
    rezultat["Конкурс"] = "да" if konkurs else "нет"
    rezultat["Сумма оплаты"] = pd.to_numeric(
        rezultat["Сумма оплаты"], errors="coerce"
    )
    return rezultat.dropna(subset=["Сумма оплаты"])


def prochitat_starye_tablicy(papka_staryh):
    fayly = {
        "Влад8.xlsx": os.path.join(papka_staryh, "Влад8.xlsx"),
        "ВладК8.xlsx": os.path.join(papka_staryh, "ВладК8.xlsx"),
        "Дима8.xlsx": os.path.join(papka_staryh, "Дима8.xlsx"),
        "Леша8.xlsx": os.path.join(papka_staryh, "Леша8.xlsx"),
    }
    net_faylov = [imya for imya, put in fayly.items() if not os.path.isfile(put)]
    if net_faylov:
        print(
            "Ошибка: в папке 'Старые таблицы' не найдены файлы:",
            ", ".join(net_faylov),
        )
        return None

    bt = {
        "kolonka_daty": "Дата оплаты заказчиком",
        "kolonka_naimenovaniya": "Номенклатура по нашей отгрузочной",
        "kolonka_firmy": "Фирма Поставщик",
        "kolonka_summy": "Сумма оплаты заказчиком",
    }

    try:
        istochniki = [
            prochitat_staryi_istochnik(
                fayly["Влад8.xlsx"], 0, "Владимир", False, **bt
            ),
            prochitat_staryi_istochnik(
                fayly["ВладК8.xlsx"], 0, "Владимир", True,
                kolonka_daty="Дата оплаты Россетями",
                kolonka_naimenovaniya="наменование",
                kolonka_firmy="Фирма",
                kolonka_summy="Сумма оплаты с НДС",
            ),
            prochitat_staryi_istochnik(
                fayly["Дима8.xlsx"], 0, "Дмитрий", False, **bt
            ),
            prochitat_staryi_istochnik(
                fayly["Леша8.xlsx"], 0, "Алексей", False, **bt
            ),
            prochitat_staryi_istochnik(
                fayly["Леша8.xlsx"], 1, "Алексей", True,
                kolonka_daty="Дата оплаты Россетями",
                kolonka_naimenovaniya="наменование",
                kolonka_firmy="Фирма",
                kolonka_summy="Сумма",
            ),
        ]
    except Exception as oshibka:
        print("Ошибка чтения старых таблиц:", oshibka)
        return None

    print("Старые таблицы успешно прочитаны из папки:", papka_staryh)
    return istochniki


def prochitat_podgotovlennye_operacii(fil, chelovek, konkurs):
    excel = pd.ExcelFile(fil)

    for list_excel in excel.sheet_names:
        df = pd.read_excel(fil, sheet_name=list_excel)
        kolonki = {
            normalizovat_tekst(kolonka): kolonka for kolonka in df.columns
        }
        obyazatelnye = [
            "Дата оплаты",
            "Номер",
            "Наименование",
            "Фирма",
            "Сумма оплаты",
            "страховка",
            "сумма закупки",
        ]

        if all(normalizovat_tekst(kolonka) in kolonki for kolonka in obyazatelnye):
            ishodnye_kolonki = [
                kolonki[normalizovat_tekst(kolonka)]
                for kolonka in obyazatelnye
            ]
            rezultat = df[ishodnye_kolonki].copy()
            rezultat.columns = obyazatelnye
            rezultat["Исполнитель"] = chelovek
            rezultat["Конкурс"] = "да" if konkurs else "нет"
            rezultat["Сумма оплаты"] = pd.to_numeric(
                rezultat["Сумма оплаты"], errors="coerce"
            )
            rezultat["страховка"] = pd.to_numeric(
                rezultat["страховка"], errors="coerce"
            )
            rezultat["сумма закупки"] = pd.to_numeric(
                rezultat["сумма закупки"], errors="coerce"
            )
            rezultat = rezultat.dropna(subset=["Сумма оплаты"])
            return rezultat

    raise ValueError(
        "Не найдены колонки: Дата оплаты, Номер, Наименование, "
        "Фирма, Сумма оплаты, страховка и сумма закупки"
    )


def sprosit_podgotovlennye_operacii(chelovek, konkurs):
    vid = "конкурс" if konkurs else "б/т"
    while True:
        fil = input(
            f"Введите название файла с операциями {chelovek} "
            f"({vid}) (-1 — назад): "
        ).strip()
        if fil == "-1":
            return None, None

        try:
            operacii = prochitat_podgotovlennye_operacii(
                fil, chelovek, konkurs
            )
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


def zapisat_sovpadenie(df, index, naydeno, stroki_sovpadeniya=None):
    df.loc[index, "Исполнитель"] = naydeno["Исполнитель"]
    df.loc[index, "Номер"] = naydeno.get("Номер")
    df.loc[index, "Наименование"] = naydeno["Наименование"]
    df.loc[index, "Фирма"] = naydeno["Фирма"]
    df.loc[index, "Конкурс"] = naydeno["Конкурс"]

    if stroki_sovpadeniya is None:
        df.loc[index, "Страховка"] = naydeno.get("страховка")
        df.loc[index, "Затраты"] = naydeno.get("сумма закупки")
    else:
        def summa_kolonki(kolonka):
            if kolonka not in stroki_sovpadeniya.columns:
                return float("nan")
            chisla = pd.to_numeric(
                stroki_sovpadeniya[kolonka], errors="coerce"
            )
            return chisla.sum(min_count=1)

        df.loc[index, "Страховка"] = summa_kolonki("страховка")
        df.loc[index, "Затраты"] = summa_kolonki("сумма закупки")


def nayti_kombinacii(indexy, nuzhnaya_summa, spravochnik, ispolzovannye):
    dostupnye = []
    for index in indexy:
        if index in ispolzovannye:
            continue
        summa = spravochnik.loc[index, "Точная сумма"]
        if summa is not None and summa > 0 and summa <= nuzhnaya_summa:
            dostupnye.append((index, summa))

    # Проверяем только комбинации из двух и трех строк. Для троек заранее
    # группируем позиции по сумме, поэтому полный перебор 2^N не возникает.
    rezultaty = []
    pozicii_po_summe = {}
    for poziciya, (_, summa) in enumerate(dostupnye):
        pozicii_po_summe.setdefault(summa, []).append(poziciya)

    def dobavit_rezultat(pozicii):
        kombinaciya = tuple(dostupnye[poziciya][0] for poziciya in pozicii)
        if kombinaciya not in rezultaty:
            rezultaty.append(kombinaciya)
        return len(rezultaty) >= 2

    # Сначала пары.
    for pervaya in range(len(dostupnye)):
        ostatok = nuzhnaya_summa - dostupnye[pervaya][1]
        pozicii = pozicii_po_summe.get(ostatok, [])
        nachalo = bisect_right(pozicii, pervaya)
        for vtoraya in pozicii[nachalo:nachalo + 2]:
            if dobavit_rezultat((pervaya, vtoraya)):
                return rezultaty

    # Затем тройки.
    for pervaya in range(len(dostupnye)):
        for vtoraya in range(pervaya + 1, len(dostupnye)):
            ostatok = (
                nuzhnaya_summa
                - dostupnye[pervaya][1]
                - dostupnye[vtoraya][1]
            )
            pozicii = pozicii_po_summe.get(ostatok, [])
            nachalo = bisect_right(pozicii, vtoraya)
            for tretya in pozicii[nachalo:nachalo + 2]:
                if dobavit_rezultat((pervaya, vtoraya, tretya)):
                    return rezultaty

    return rezultaty


def dobavit_ispolnitelya(
    tablica=None,
    papka_rezultatov=".",
    god=None,
    mesyac=None,
    gotovye_operacii=None,
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

    for kolonka in ["Номер", "Наименование", "Фирма", "Конкурс"]:
        if kolonka not in df.columns:
            df[kolonka] = None
        df[kolonka] = df[kolonka].astype("object")

    for kolonka in ["Страховка", "Затраты"]:
        if kolonka not in df.columns:
            df[kolonka] = float("nan")
        df[kolonka] = pd.to_numeric(df[kolonka], errors="coerce")

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
        if gotovye_operacii is None:
            vse_operacii = []
            for chelovek in ["Алексей", "Владимир", "Дмитрий"]:
                for konkurs in [True, False]:
                    _, operacii = sprosit_podgotovlennye_operacii(
                        chelovek, konkurs
                    )
                    if operacii is None:
                        return novye_tablicy, None
                    vse_operacii.append(operacii)
        else:
            vse_operacii = gotovye_operacii

        vse_operacii = [operacii.copy() for operacii in vse_operacii]
        for nomer_istochnika, operacii in enumerate(vse_operacii):
            operacii["Источник"] = nomer_istochnika

        spravochnik = pd.concat(vse_operacii, ignore_index=True)
        mesyac_chislom = nomer_mesaca(mesyac)
        if mesyac_chislom is None or not str(god).isdigit():
            print("Ошибка: не удалось определить рабочий год или месяц")
            return novye_tablicy, None

        spravochnik["Даты оплаты"] = spravochnik["Дата оплаты"].apply(
            poluchit_daty
        )
        spravochnik = spravochnik[
            spravochnik["Даты оплаты"].apply(
                lambda daty: any(
                    data.year == int(god) and data.month == mesyac_chislom
                    for data in daty
                )
            )
        ].copy()
        spravochnik.reset_index(drop=True, inplace=True)

        spravochnik["Точная сумма"] = spravochnik["Сумма оплаты"].apply(
            tochnaya_summa
        )
        spravochnik["Наименование для сравнения"] = spravochnik[
            "Наименование"
        ].apply(normalizovat_tekst)
        if "Номер" not in spravochnik.columns:
            spravochnik["Номер"] = None
        spravochnik["Номер для сравнения"] = spravochnik["Номер"].apply(
            normalizovat_tekst
        )
        spravochnik["Группа для сравнения"] = spravochnik[
            "Номер для сравнения"
        ].where(
            spravochnik["Номер для сравнения"] != "",
            spravochnik["Наименование для сравнения"],
        )

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
                zapisat_sovpadenie(df, index, naydeno)
                ispolzovannye_stroki.add(kandidaty.index[0])
            else:
                neraspredelennye_stroki.append(index)

        # В новых таблицах комбинации группируются по номеру и дате. Для
        # старых таблиц без номера временно используется наименование.
        gruppy = {}
        for spravochnik_index, stroka in spravochnik.iterrows():
            gruppa = stroka["Группа для сравнения"]
            if not gruppa:
                continue
            for data in stroka["Даты оплаты"]:
                if data.year != int(god) or data.month != mesyac_chislom:
                    continue
                kluch = (stroka["Источник"], data, gruppa)
                gruppy.setdefault(kluch, []).append(spravochnik_index)

        # Затем для оставшихся операций проверяем комбинации из двух и трех
        # еще не использованных покупок.
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
                _, kombinaciya = podhodyashie_kombinacii[0]
                pervaya = spravochnik.loc[kombinaciya[0]]
                zapisat_sovpadenie(
                    df,
                    index,
                    pervaya,
                    spravochnik.loc[list(kombinaciya)],
                )
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
