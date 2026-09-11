import os

import pandas as pd
import numpy as np

from tablica import Tablica
from tablica_vseh_operaciy import nayti_nashu_kompaniyu


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


def eto_itogovaya_yacheyka(znachenie):
    """Распознаёт отдельную служебную надпись ИТОГО с пунктуацией."""
    tekst = normalizovat_nazvanie(znachenie)
    tekst = tekst.rstrip(":;. ")
    return tekst in {"итого", "итоги", "всего"}


def ubrat_itog_i_stroki_posle(df):
    """Отсекает первую строку итогов и всё, что находится ниже неё."""
    maska_itogo = df.apply(
        lambda kolonka: kolonka.map(eto_itogovaya_yacheyka)
    ).any(axis=1)
    pozicii = np.flatnonzero(maska_itogo.to_numpy())
    if len(pozicii) == 0:
        return df
    return df.iloc[:pozicii[0]].copy()


def podtverdit_itogovye_summy(df):
    """Показывает контрольные суммы выписки и просит подтвердить их."""
    summa_debet = df["Оборот Дт"].fillna(0).sum()
    summa_kredit = df["Оборот Кт"].fillna(0).sum()

    def format_summy(summa):
        return f"{summa:,.2f}".replace(",", " ").replace(".", ",")

    print("\nПроверьте итоговые суммы по исходной банковской выписке:")
    print("Дт:", format_summy(summa_debet))
    print("Кт:", format_summy(summa_kredit))

    polozhitelnye_otvety = {"да", "д", "yes", "y"}
    otricatelnye_otvety = {"нет", "н", "no", "n"}
    dopustimye_otvety = polozhitelnye_otvety | otricatelnye_otvety

    otvet = input("Суммы совпадают? Введите да/нет или yes/no: ").strip().lower()
    while otvet not in dopustimye_otvety:
        otvet = input("Введите да/нет или yes/no: ").strip().lower()

    if otvet in otricatelnye_otvety:
        print("Обработка отменена. Возвращаемся в главное меню.")
        return False
    return True


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

    # В некоторых выписках ИТОГО находится в объединенной ячейке колонки,
    # которая не входит в единый формат. Отсекаем итог до reindex, пока эта
    # служебная надпись еще не потеряна.
    df = ubrat_itog_i_stroki_posle(df)

    # Для всех банков используем одинаковые колонки и одинаковый порядок.
    # Если в банковском формате колонки нет, она останется пустой.
    df = df.reindex(columns=EDINYE_KOLONKI)

    return df


def dobavit_tip_operacii(
    tablica=None, papka_rezultatov=".", podtverzhdat_summy=True
):
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

    else:
        fil = tablica.imya
        df = tablica.df.copy()

    # Итоги исключаются как из нового файла, так и из таблицы, которая уже
    # была загружена в начале сессии. Надпись может стоять в любой колонке.
    df = ubrat_itog_i_stroki_posle(df)

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

    if podtverzhdat_summy and not podtverdit_itogovye_summy(df):
        return None, None, None

    if tablica is None:
        # Сохраняем отдельно исходную таблицу, подготовленную для обработки.
        imya_ishodnoy = os.path.splitext(os.path.basename(fil))[0] + ".xlsx"
        novaya_tablica = Tablica(
            os.path.join(papka_rezultatov, imya_ishodnoy), df.copy()
        )

    naznachenie = df["Назначение платежа"].fillna("").astype(str)
    oplata_ili_vozmeshenie = naznachenie.str.contains(
        r"Оплата по дог|Возм\. по дог\.", case=False
    )
    obespechenie_ili_garantiya = naznachenie.str.contains(
        r"обеспеч|гарант", case=False
    )
    bankovskaya_komissiya = naznachenie.str.contains(
        r"комиссия|ком\.\s*за", case=False
    )
    vozvrat_depozita = (
        naznachenie.str.contains(r"возврат", case=False)
        & naznachenie.str.contains(r"депозит", case=False)
    )
    kontragent = df["Корреспондент"].fillna("").astype(str)
    kontragent_rad = kontragent.str.contains(r"РАД", case=False)
    kontragent_okean_servis = kontragent.str.contains(
        r"океан[\s-]*сервис", case=False
    )
    oplata_za_tehobsluzhivanie = (
        naznachenie.str.contains(
            r"Оплата за тех\. обслуживание", case=False
        )
        | (
            naznachenie.str.contains(
                r"оплата\s+за\s+тех\.?\s*обслуж", case=False
            )
            & kontragent_okean_servis
        )
    )
    est_debet = df["Оборот Дт"].fillna(0) != 0
    est_kredit = df["Оборот Кт"].fillna(0) != 0

    conditions = [
        vozvrat_depozita,
        obespechenie_ili_garantiya & kontragent_rad & ~bankovskaya_komissiya,
        obespechenie_ili_garantiya & ~kontragent_rad & ~bankovskaya_komissiya,
        oplata_ili_vozmeshenie & est_kredit,
        oplata_ili_vozmeshenie & est_debet,
        oplata_za_tehobsluzhivanie,
        naznachenie.str.contains("пени|штраф|взыск|неустойк", case=False),
        naznachenie.str.contains("Выплата процентов согласно депозитного договора|УПЛАТА ПРОЦЕНТОВ ДЕПОЗИТ", case=False),
        naznachenie.str.contains("Пополнение счета согласно депозитного договора|Размещение денежных средств во Вклад", case=False),
        naznachenie.str.contains("аренд", case=False),
        naznachenie.str.contains("налог|Страховые взносы", case=False),
        naznachenie.str.contains("заработной платы|заработная плата", case=False),
        bankovskaya_komissiya,
        naznachenie.str.contains("РАД|Плата оператору", case=False),
        naznachenie.str.contains("Займ|займ", case=False),
        naznachenie.str.contains("Выдача денежных средств|Снятие по карте", case=False),
        naznachenie.str.contains("Перевод собственных средств", case=False),
        naznachenie.str.contains("Возврат средств|Возврат денежных средств", case=False),
        naznachenie.str.contains("топлив", case=False),
        naznachenie.str.contains("Оплата", case=False),
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
        15: "возврат средств",
        16: "ППР",
        17: "штрафы",
        18: "БГ",
    }

    # Типы здесь идут в том же порядке, что и автоматические условия выше.
    avtomaticheskie_tipy = [
        "депозит возврат",
        "РАД",
        "БГ",
        "пришло",
        "ППР",
        "ПО",
        "штрафы",
        "% депозит",
        "депозит",
        "аренда",
        "налог",
        "ЗП",
        "банк комиссия",
        "РАД",
        "займ",
        "снятие дс",
        "перевод сс",
        "возврат средств",
        "ППР",
        "товар",
    ]

    df["Тип операции"] = np.select(
        conditions,
        avtomaticheskie_tipy,
        default="другое",
    )

    # Общая формулировка "Оплата" обычно означает товар. Но если деньги
    # переводятся между нашими компаниями, это собственные средства, а не
    # покупка товара. Проверяем обе стороны: операции в Дт и в Кт.
    nash_kontragent = df["Корреспондент"].apply(nayti_nashu_kompaniyu)
    df.loc[
        (df["Тип операции"] == "товар") & nash_kontragent.notna(),
        "Тип операции",
    ] = "перевод сс"

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
    tablica_nerasp = None
    if not nerasp.empty:
        nerasp[["Оборот Дт", "Оборот Кт", "Назначение платежа"]].to_excel(
            fail_nerasp, index=False
        )
        tablica_nerasp = Tablica(fail_nerasp, nerasp.copy())

    print("\nТипы операций:")
    for nomer, tip in tipy.items():
        print(nomer, "-", tip)
    print("\nСохрани таблицу или сфоткай")
    if tablica_nerasp is not None:
        print("Нераспознанные операции сохранены в файле:", fail_nerasp)
    else:
        print("Все операции распознаны, отдельный файл не создается")

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
        tablica_nerasp,
    )
