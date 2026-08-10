import os
import re
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from itogovaya_summa import nayti_chislo
from tablica import Tablica


NOL = Decimal("0.00")
SHAG = Decimal("0.01")

STROKI_ITOGA = [
    ("Денежные средства", "р/с"),
    ("Денежные средства", "РАД"),
    ("Денежные средства", "ППР"),
    ("Денежные средства", "БГ"),
    ("Денежные средства", "Депозит"),
    ("Товар", "Алексей"),
    ("Товар", "Владимир"),
    ("Товар", "Дмитрий"),
    ("Товар", "Андрей"),
    ("Займ плюс", "Внутренний"),
    ("Займ плюс", "Внешний"),
    ("Собственные средства", ""),
    ("Займ минус", "Внутренний"),
    ("Займ минус", "Внешний"),
    ("Прибыль", "Алексей"),
    ("Прибыль", "Владимир"),
    ("Прибыль", "Дмитрий"),
    ("Прибыль", "% депозит"),
    ("Прибыль", "Возврат налога"),
    ("Убыток", "Эквайринг"),
    ("Убыток", "Эквайринг НДС"),
    ("Убыток", "Офис"),
    ("Убыток", "Снятие"),
]

SOTRUDNIKI_TOVARA = {
    "алексей": "Алексей",
    "владимир": "Владимир",
    "дмитрий": "Дмитрий",
    "андрей": "Андрей",
}
SOTRUDNIKI_PRIBYLI = {
    "алексей": "Алексей",
    "владимир": "Владимир",
    "дмитрий": "Дмитрий",
}
OFISNYE_TIPY = {"банк комиссия", "штрафы", "по", "зп", "аренда"}


def normalizovat_tekst(znachenie):
    if pd.isna(znachenie):
        return ""
    return " ".join(
        str(znachenie).strip().lower().replace("ё", "е").split()
    )


def denezhnoe_chislo(znachenie):
    if pd.isna(znachenie) or str(znachenie).strip() == "":
        return NOL
    try:
        return Decimal(str(znachenie)).quantize(SHAG, rounding=ROUND_HALF_UP)
    except (InvalidOperation, ValueError, TypeError):
        return None


def nazvanie_firmy(imya_fayla):
    osnova = os.path.splitext(os.path.basename(imya_fayla))[0]
    return osnova.split("_", 1)[0].strip()


def oformit_itogovyi_excel(imya_fayla, neuchtennye=False):
    """Делает итоговые файлы читаемыми в старых версиях Excel."""
    kniga = load_workbook(imya_fayla)
    list_excel = kniga.active
    list_excel.freeze_panes = "A2"
    list_excel.auto_filter.ref = list_excel.dimensions

    zalivka = PatternFill(fill_type="solid", fgColor="D9EAF7")
    for yacheyka in list_excel[1]:
        yacheyka.font = Font(bold=True)
        yacheyka.fill = zalivka

    if neuchtennye:
        shiriny = [38, 18, 14, 16, 24, 18, 16, 16, 16, 24, 55, 48]
    else:
        shiriny = [14, 22, 25, 22, 18]
        for yacheyka in list_excel["A"][1:]:
            yacheyka.number_format = "dd.mm.yyyy"
        for yacheyka in list_excel["E"][1:]:
            yacheyka.number_format = "#,##0.00"

    for nomer, shirina in enumerate(shiriny, start=1):
        list_excel.column_dimensions[get_column_letter(nomer)].width = shirina
    kniga.save(imya_fayla)


def eto_vnutrennyaya_kompaniya(kontragent):
    tekst = normalizovat_tekst(kontragent)
    if not tekst:
        return None

    nazvaniya = [
        "альтегра",
        "альтэгра",
        "авк",
        "билд",
        "вектор",
        "макрон",
        "позитрон",
        "сити",
        "кит",
        "энергопоинт",
        "энерго поинт",
        "факторион",
    ]
    if any(nazvanie in tekst for nazvanie in nazvaniya):
        return True
    return bool(re.search(r"(^|\W)эп($|\W)", tekst, flags=re.IGNORECASE))


def vybrat_neskolko_tablic(tablicy):
    istochniki = [
        tablica
        for tablica in tablicy
        if {"Тип операции", "Оборот Дт", "Оборот Кт"}.issubset(
            tablica.df.columns
        )
    ]
    if not istochniki:
        print(
            "В текущей сессии нет подходящих таблиц. "
            "Сначала добавьте типы операций и исполнителей."
        )
        return None

    print("\nВыберите таблицы компаний:")
    for nomer, tablica in enumerate(istochniki, start=1):
        print(nomer, "-", tablica.imya)
    print("-1 - вернуться в главное меню")

    vybor = input("Введите номера таблиц через запятую: ").strip()
    while True:
        if vybor == "-1":
            return None
        try:
            nomera = [int(nomer.strip()) for nomer in vybor.split(",")]
            nomera = list(dict.fromkeys(nomera))
            if nomera and all(1 <= nomer <= len(istochniki) for nomer in nomera):
                return [istochniki[nomer - 1] for nomer in nomera]
        except ValueError:
            pass
        vybor = input("Нет таких номеров. Введите еще раз: ").strip()


def poluchit_komissiyu_i_nds(stroka, kolonki, naznachenie):
    rezultaty = []
    for kolonka, metka in [("Комиссия", "ком"), ("НДС", "ндс")]:
        if kolonka in kolonki and not pd.isna(stroka.get(kolonka)):
            chislo = denezhnoe_chislo(stroka.get(kolonka))
        else:
            naydeno = nayti_chislo(naznachenie, metka)
            chislo = denezhnoe_chislo(naydeno) if naydeno is not None else None
        rezultaty.append(chislo)
    return tuple(rezultaty)


def proverit_tablicu(tablica):
    df = tablica.df
    obyazatelnye = [
        "Тип операции",
        "Оборот Дт",
        "Оборот Кт",
        "Назначение платежа",
        "Исполнитель",
        "Затраты",
    ]
    net_kolonok = [kolonka for kolonka in obyazatelnye if kolonka not in df.columns]
    if net_kolonok:
        return [
            "нет колонок: " + ", ".join(net_kolonok)
            + ". Сначала выполните пункты 1–3 для этой таблицы"
        ]

    oshibki = []
    tipy = df["Тип операции"].apply(normalizovat_tekst)
    ispolniteli = df["Исполнитель"].apply(normalizovat_tekst)
    nuzhen_ispolnitel = tipy.isin({"товар", "возврат средств", "пришло"})

    net_ispolnitelya = df.index[nuzhen_ispolnitel & (ispolniteli == "")]
    if len(net_ispolnitelya):
        stroki = ", ".join(str(index + 2) for index in net_ispolnitelya[:10])
        oshibki.append(
            "не указан исполнитель в строках Excel: " + stroki
            + ". Добавьте исполнителей везде"
        )

    neizvestnye = df.index[
        nuzhen_ispolnitel
        & (ispolniteli != "")
        & (~ispolniteli.isin(SOTRUDNIKI_TOVARA))
    ]
    if len(neizvestnye):
        stroki = ", ".join(str(index + 2) for index in neizvestnye[:10])
        oshibki.append("неизвестный исполнитель в строках Excel: " + stroki)

    zatraty = pd.to_numeric(df["Затраты"], errors="coerce")
    net_zatrat = df.index[(tipy == "пришло") & zatraty.isna()]
    if len(net_zatrat):
        stroki = ", ".join(str(index + 2) for index in net_zatrat[:10])
        oshibki.append(
            "не заполнены затраты для операций 'пришло' в строках Excel: "
            + stroki
        )

    vozmeshenie = df["Назначение платежа"].fillna("").astype(str).str.contains(
        r"возм\.*\s*по\s*дог\.*", case=False, regex=True
    )
    for index in df.index[vozmeshenie]:
        komissiya, nds = poluchit_komissiyu_i_nds(
            df.loc[index], df.columns, df.loc[index, "Назначение платежа"]
        )
        if komissiya is None or nds is None:
            oshibki.append(
                f"в строке Excel {index + 2} не найдены комиссия или НДС. "
                "Сначала снова выполните пункт 2 для этой таблицы"
            )
    return oshibki


def dobavit_neuchtennuyu(neuchtennye, tablica, firma, index, stroka, prichina):
    dt = denezhnoe_chislo(stroka.get("Оборот Дт"))
    kt = denezhnoe_chislo(stroka.get("Оборот Кт"))
    summa = kt if kt not in {None, NOL} else dt
    neuchtennye.append({
        "Файл": tablica.imya,
        "Фирма": firma,
        "Строка Excel": index + 2,
        "Документ": stroka.get("Документ"),
        "Контрагент": stroka.get("Корреспондент"),
        "Тип операции": stroka.get("Тип операции"),
        "Исполнитель": stroka.get("Исполнитель"),
        "Оборот Дт": float(dt) if dt is not None else None,
        "Оборот Кт": float(kt) if kt is not None else None,
        "Сумма для ручного добавления": float(summa) if summa is not None else None,
        "Назначение платежа": stroka.get("Назначение платежа"),
        "Причина": prichina,
    })


def rasschitat_firmu(tablica, data_itoga):
    firma = nazvanie_firmy(tablica.imya)
    df = tablica.df.copy()
    itogi = {kluch: NOL for kluch in STROKI_ITOGA}
    neuchtennye = []

    dt_vse = pd.to_numeric(df["Оборот Дт"], errors="coerce").fillna(0).sum()
    kt_vse = pd.to_numeric(df["Оборот Кт"], errors="coerce").fillna(0).sum()
    itogi[("Денежные средства", "р/с")] = denezhnoe_chislo(kt_vse - dt_vse)

    for index, stroka in df.iterrows():
        dt = denezhnoe_chislo(stroka.get("Оборот Дт"))
        kt = denezhnoe_chislo(stroka.get("Оборот Кт"))
        if dt is None or kt is None:
            dobavit_neuchtennuyu(
                neuchtennye, tablica, firma, index, stroka,
                "сумма в Дт или Кт не является числом",
            )
            continue
        if dt == NOL and kt == NOL:
            continue
        if dt != NOL and kt != NOL:
            dobavit_neuchtennuyu(
                neuchtennye, tablica, firma, index, stroka,
                "одновременно заполнены Дт и Кт",
            )
            continue

        tip = normalizovat_tekst(stroka.get("Тип операции"))
        ispolnitel = normalizovat_tekst(stroka.get("Исполнитель"))
        naznachenie = str(stroka.get("Назначение платежа", ""))
        summa = dt if dt != NOL else kt

        # Комиссия и НДС — отдельные части возмещения. Они не входят в
        # прибыль сотрудника, поэтому учитываются отдельно в убытке.
        vozmeshenie = bool(re.search(
            r"возм\.*\s*по\s*дог\.*", naznachenie, flags=re.IGNORECASE
        ))
        if vozmeshenie:
            komissiya, nds = poluchit_komissiyu_i_nds(
                stroka, df.columns, naznachenie
            )
            if komissiya is not None:
                itogi[("Убыток", "Эквайринг")] += komissiya
            if nds is not None:
                itogi[("Убыток", "Эквайринг НДС")] += nds

        if tip in {"рад", "ппр", "бг"}:
            podtip = {"рад": "РАД", "ппр": "ППР", "бг": "БГ"}[tip]
            itogi[("Денежные средства", podtip)] += summa
        elif tip == "депозит":
            itogi[("Денежные средства", "Депозит")] += summa
        elif tip == "депозит возврат":
            itogi[("Денежные средства", "Депозит")] -= summa
        elif tip == "товар":
            chelovek = SOTRUDNIKI_TOVARA.get(ispolnitel)
            itogi[("Товар", chelovek)] += summa
        elif tip == "возврат средств":
            chelovek = SOTRUDNIKI_TOVARA.get(ispolnitel)
            itogi[("Товар", chelovek)] -= summa
        elif tip == "займ":
            vnutrennii = eto_vnutrennyaya_kompaniya(stroka.get("Корреспондент"))
            if vnutrennii is None:
                dobavit_neuchtennuyu(
                    neuchtennye, tablica, firma, index, stroka,
                    "не указан контрагент для определения вида займа",
                )
                continue
            podtip = "Внутренний" if vnutrennii else "Внешний"
            est_perevod = bool(re.search(
                r"перевод|перечислен", naznachenie, flags=re.IGNORECASE
            ))
            est_vozvrat = bool(re.search(
                r"возврат", naznachenie, flags=re.IGNORECASE
            ))
            if dt != NOL and est_perevod:
                itogi[("Займ плюс", podtip)] += dt
            elif kt != NOL and est_vozvrat:
                itogi[("Займ плюс", podtip)] -= kt
            elif kt != NOL and est_perevod:
                itogi[("Займ минус", podtip)] += kt
            elif dt != NOL and est_vozvrat:
                itogi[("Займ минус", podtip)] -= dt
            else:
                dobavit_neuchtennuyu(
                    neuchtennye, tablica, firma, index, stroka,
                    "для займа не найдено нужное слово перевод/перечисление/возврат",
                )
        elif tip == "перевод сс":
            itogi[("Собственные средства", "")] += dt - kt
        elif tip == "% депозит":
            itogi[("Прибыль", "% депозит")] += summa
        elif tip == "налог" and kt != NOL:
            itogi[("Прибыль", "Возврат налога")] += kt
        elif tip == "пришло":
            chelovek = SOTRUDNIKI_PRIBYLI.get(ispolnitel)
            if chelovek is None:
                dobavit_neuchtennuyu(
                    neuchtennye, tablica, firma, index, stroka,
                    "для этого исполнителя нет строки в разделе прибыли",
                )
                continue
            if kt == NOL:
                dobavit_neuchtennuyu(
                    neuchtennye, tablica, firma, index, stroka,
                    "операция 'пришло' должна находиться в Кт",
                )
                continue
            zatraty = denezhnoe_chislo(stroka.get("Затраты"))
            itogi[("Прибыль", chelovek)] += kt - zatraty
        elif tip == "налог" and dt != NOL:
            itogi[("Убыток", "Офис")] += dt
        elif tip in OFISNYE_TIPY:
            itogi[("Убыток", "Офис")] += summa
        elif tip in {"снятие сс", "снятие дс"}:
            itogi[("Убыток", "Снятие")] += summa
        else:
            dobavit_neuchtennuyu(
                neuchtennye, tablica, firma, index, stroka,
                "тип операции не входит в правила итоговой таблицы",
            )

    stroki = [
        {
            "Дата": data_itoga,
            "Фирма": firma,
            "Тип": tip,
            "Подтип": podtip,
            "Сумма": float(itogi[(tip, podtip)].quantize(SHAG)),
        }
        for tip, podtip in STROKI_ITOGA
    ]
    return stroki, neuchtennye


def sformirovat_itogovuyu_tablicu(tablicy, papka_mesaca, god, mesyac, mesyac_chislom):
    vybrannye = vybrat_neskolko_tablic(tablicy)
    if vybrannye is None:
        return []

    firmy = [normalizovat_tekst(nazvanie_firmy(tablica.imya)) for tablica in vybrannye]
    if len(firmy) != len(set(firmy)):
        print(
            "Ошибка: выбрано несколько таблиц одной фирмы. "
            "Выберите по одному итоговому файлу для каждой фирмы."
        )
        return []

    vse_oshibki = []
    for tablica in vybrannye:
        for oshibka in proverit_tablicu(tablica):
            vse_oshibki.append(f"{tablica.imya}: {oshibka}")
    if vse_oshibki:
        print("\nИтоговая таблица не сформирована:")
        for oshibka in vse_oshibki:
            print("-", oshibka)
        print("Исправьте таблицы и запустите пункт ещё раз.")
        return []

    data_itoga = pd.Timestamp(year=int(god), month=mesyac_chislom, day=1)
    vse_stroki = []
    vse_neuchtennye = []
    for tablica in vybrannye:
        stroki, neuchtennye = rasschitat_firmu(tablica, data_itoga)
        vse_stroki.extend(stroki)
        vse_neuchtennye.extend(neuchtennye)

    itogovyi_df = pd.DataFrame(
        vse_stroki, columns=["Дата", "Фирма", "Тип", "Подтип", "Сумма"]
    )
    imya_itoga = os.path.join(papka_mesaca, f"итоговая_таблица_{mesyac}.xlsx")
    itogovyi_df.to_excel(imya_itoga, index=False)
    oformit_itogovyi_excel(imya_itoga)
    rezultaty = [Tablica(imya_itoga, itogovyi_df)]
    print("\nИтоговая таблица сохранена:", imya_itoga)

    if vse_neuchtennye:
        neuchtennye_df = pd.DataFrame(vse_neuchtennye)
        imya_neuchtennyh = os.path.join(
            papka_mesaca, f"итоговая_таблица_{mesyac}_неучтенные.xlsx"
        )
        neuchtennye_df.to_excel(imya_neuchtennyh, index=False)
        oformit_itogovyi_excel(imya_neuchtennyh, neuchtennye=True)
        rezultaty.append(Tablica(imya_neuchtennyh, neuchtennye_df))
        print(
            "Неучтённые операции сохранены отдельно:",
            imya_neuchtennyh,
        )
        print("Количество неучтённых операций:", len(vse_neuchtennye))
    else:
        print("Все операции учтены по заданным правилам.")

    return rezultaty
