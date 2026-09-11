import contextlib
import io
from pathlib import Path
import shutil
import subprocess
import sys
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

import pandas as pd

from dobavit_ispolnitelya import dobavit_ispolnitelya, normalizovat_nomer
from obrabotka_altegra import dobavit_tip_operacii
from operacii_ispolnitelya import sobrat_operacii_ispolnitelya
from sformirovat_itogovuyu_tablicu import rasschitat_firmu
from tablica import Tablica


def employee_rows(amounts, numbers=None, dates=None):
    return pd.DataFrame({
        "Дата оплаты": dates or ["15.09.2026"] * len(amounts),
        "Номер": numbers or ["A"] * len(amounts),
        "Наименование": [f"Позиция {i}" for i in range(len(amounts))],
        "Фирма": "Тест",
        "Сумма оплаты": amounts,
        "страховка": 1,
        "сумма закупки": 2,
        "Исполнитель": "Алексей",
        "Конкурс": "да",
    })


def bank_rows(amounts):
    return pd.DataFrame({
        "Тип операции": "пришло",
        "Назначение платежа": "Оплата по договору",
        "Оборот Дт": float("nan"),
        "Оборот Кт": amounts,
        "Итоговая сумма": amounts,
    })


class SessionAndMatchingTests(unittest.TestCase):
    def setUp(self):
        self.temp = TemporaryDirectory(prefix="buhuchet-test-")
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)
        self.output = contextlib.redirect_stdout(io.StringIO())
        self.output.__enter__()
        self.addCleanup(self.output.__exit__, None, None, None)

    def match(self, amounts, sources=None, cache=None):
        return dobavit_ispolnitelya(
            Tablica("Тест.xlsx", bank_rows(amounts)), str(self.folder),
            "2026", "сентябрь", gotovye_operacii=sources,
            kesh_novyh_tablic=cache,
        )[1]

    def test_all_singles_before_groups_and_first_unused_match(self):
        source = employee_rows([30, 70, 100, 100, 40, 60],
                               ["A", "A", "B", "C", "D", "D"])
        result = self.match([100, 100, 100, 30], [source]).df
        # The last bank row reserves 30 before the earlier unmatched 100
        # can use 30+70. The remaining 100 therefore uses group D.
        self.assertEqual(result["Номер"].tolist(), ["B", "C", "D", "A"])
        self.assertEqual(result["Затраты"].tolist(), [2, 2, 4, 2])

    def test_full_group_and_pairs_and_triples(self):
        for amounts in ([20, 30], [10, 20, 30], [10, 20, 30, 40]):
            with self.subTest(amounts=amounts):
                result = self.match([sum(amounts)], [employee_rows(amounts)]).df
                self.assertEqual(result.loc[0, "Исполнитель"], "Алексей")
                self.assertEqual(result.loc[0, "Затраты"], 2 * len(amounts))

    def test_groups_do_not_cross_number_date_or_source(self):
        cases = [
            [employee_rows([30, 70], ["A", "B"])],
            [employee_rows([30, 70], dates=["15.09.2026", "16.09.2026"])],
            [employee_rows([30]), employee_rows([70])],
            [employee_rows([30, 70], dates=["15.08.2026"] * 2)],
            [employee_rows([20, 80, 30, 70])],  # two possible pairs
        ]
        for sources in cases:
            with self.subTest(sources=len(sources)):
                result = self.match([100], sources).df
                self.assertTrue(pd.isna(result.loc[0, "Исполнитель"]))

    def test_cache_reuse_and_cancelled_replacement(self):
        sources = [employee_rows([100 + i]) for i in range(6)]
        cache = {}
        with patch("dobavit_ispolnitelya.sprosit_podgotovlennye_operacii",
                   side_effect=[("file.xlsx", source) for source in sources]) as loader:
            first = self.match([100], cache=cache)
        self.assertEqual(loader.call_count, 6)
        with patch("dobavit_ispolnitelya.sprosit_podgotovlennye_operacii",
                   side_effect=AssertionError("Must use memory")):
            second = self.match([100], cache["операции"])
        pd.testing.assert_frame_equal(first.df, second.df)
        for original, cached in zip(sources, cache["операции"]):
            pd.testing.assert_frame_equal(original, cached)
        previous = cache["операции"]
        with patch("dobavit_ispolnitelya.sprosit_podgotovlennye_operacii",
                   side_effect=[("new.xlsx", employee_rows([999])), (None, None)]):
            self.assertIsNone(self.match([100], cache=cache))
        self.assertIs(cache["операции"], previous)

    def test_tracking_column_new_and_existing_employee_files(self):
        for number, name in enumerate(["Алексей", "Владимир", "Дмитрий", "Андрей"], 1):
            with self.subTest(name=name):
                source = Tablica("Источник.xlsx", pd.DataFrame({"Исполнитель": [name]}))
                with patch("builtins.input", side_effect=["1", str(number), "1"]):
                    result, _ = sobrat_operacii_ispolnitelya([source], str(self.folder), "2026")
                saved = pd.read_excel(result.imya)
                self.assertEqual(saved.columns[-1], "отслежено")
                self.assertEqual(saved["отслежено"].tolist(), ["нет"])
                saved.loc[0, "отслежено"] = "да"
                saved.to_excel(result.imya, index=False)
                with patch("builtins.input", side_effect=["1", str(number), "1"]):
                    result, _ = sobrat_operacii_ispolnitelya([source], str(self.folder), "2026")
                self.assertEqual(pd.read_excel(result.imya)["отслежено"].tolist(), ["да", "нет"])
                pd.DataFrame({"Исполнитель": [name]}).to_excel(result.imya, index=False)
                with patch("builtins.input", side_effect=["1", str(number), "1"]):
                    result, _ = sobrat_operacii_ispolnitelya([source], str(self.folder), "2026")
                self.assertEqual(pd.read_excel(result.imya)["отслежено"].tolist(), ["нет", "нет"])

    def test_main_reports_missing_automatic_input_files(self):
        project = Path(__file__).resolve().parents[1]
        for path in project.glob("*.py"):
            shutil.copy2(path, self.folder / path.name)
        answers = ["2026", "сентябрь", "-1"]
        process = subprocess.run(
            [sys.executable, "-B", str(self.folder / "buhuchet.py")],
            input="\n".join(answers) + "\n", text=True, capture_output=True,
            cwd=self.folder, timeout=30,
        )
        self.assertEqual(process.returncode, 0, process.stdout + process.stderr)
        self.assertIn("Не найдены", process.stdout)
        self.assertIn("АлексейБТ", process.stdout)

    def test_operation_number_is_text_identifier(self):
        self.assertEqual(normalizovat_nomer("0012"), "0012")
        self.assertEqual(normalizovat_nomer("A-12"), "a-12")
        self.assertEqual(normalizovat_nomer(12.0), "12")

    def test_main_automatically_processes_named_files(self):
        project = Path(__file__).resolve().parents[1]
        for path in project.glob("*.py"):
            shutil.copy2(path, self.folder / path.name)

        header = [
            "Документ", "Дата операции", "Корреспондент", "Оборот Дт",
            "Оборот Кт", "Назначение платежа",
        ]
        bank = pd.DataFrame([
            header, [None] * len(header),
            [1, "15.09.2026", "Тест", None, 100, "Оплата по договору"],
        ])
        companies = [
            "Альтэгра", "АВК", "Билд", "Вектор", "Макрон", "Позитрон",
            "Сити", "Кит", "Энергопоинт", "Факторион",
        ]
        for company in companies:
            bank.to_excel(self.folder / f"{company}.xlsx", header=False, index=False)

        for filename in [
            "АлексейБТ", "АлексейК", "ВладимирБТ", "ВладимирК",
            "ДмитрийБТ", "ДмитрийК",
        ]:
            employee_rows([100], dates=["15.09.2026"]).to_excel(
                self.folder / f"{filename}.xlsx", index=False
            )

        first_process = subprocess.run(
            [sys.executable, "-B", str(self.folder / "buhuchet.py")],
            input="2026\nсентябрь\n3\n-1\n",
            text=True, capture_output=True, cwd=self.folder, timeout=60,
        )
        self.assertEqual(
            first_process.returncode, 0,
            first_process.stdout + first_process.stderr,
        )
        month = self.folder / "2026" / "сентябрь"
        self.assertTrue((month / "Кит_готовая.xlsx").is_file())
        ready_path = month / "Кит_готовая.xlsx"
        checked = pd.read_excel(ready_path)
        checked.loc[0, "Исполнитель"] = "Владимир"
        checked.to_excel(ready_path, index=False)

        second_process = subprocess.run(
            [sys.executable, "-B", str(self.folder / "buhuchet.py")],
            input="2026\nсентябрь\n4\n5\n6\n7\n-1\n",
            text=True, capture_output=True, cwd=self.folder, timeout=60,
        )
        self.assertEqual(
            second_process.returncode, 0,
            second_process.stdout + second_process.stderr,
        )
        self.assertTrue((month / "Премии сотрудников сентябрь.xlsx").is_file())
        self.assertTrue((month / "Все операции.xlsx").is_file())
        self.assertTrue((self.folder / "2026" / "Итоговая таблица 2026.xlsx").is_file())
        saved = pd.read_excel(ready_path)
        self.assertEqual(saved.loc[0, "Исполнитель"], "Владимир")

    def test_classification_and_separate_fines_in_summary(self):
        source = pd.DataFrame({
            "Документ": [1, 2, 3, 4],
            "Корреспондент": ["Океан Сервис", "Тестовый банк", "Тест", "Тест"],
            "Оборот Дт": [100, 0, 0, 67.89],
            "Оборот Кт": [0, 1000, 123.45, 0],
            "Назначение платежа": ["оплата за техобслуж",
                "возврат средств согласно депозитн договора",
                "Штраф по договору", "Оплата штрафа"],
        })
        with patch("builtins.input", return_value="да"):
            _, result, unknown = dobavit_tip_operacii(
                Tablica("Тест.xlsx", source), str(self.folder)
            )
        self.assertIsNone(unknown)
        saved = pd.read_excel(result.imya)
        self.assertEqual(saved["Тип операции"].tolist(),
                         ["ПО", "депозит возврат", "штрафы", "штрафы"])
        result.df["Затраты"] = 0
        rows, unknown = rasschitat_firmu(result, pd.Timestamp("2026-09-01"))
        self.assertFalse(unknown)
        totals = {(row["Тип"], row["Подтип"]): row["Сумма"] for row in rows}
        self.assertEqual(totals["Прибыль", "Штраф"], 123.45)
        self.assertEqual(totals["Убыток", "Штраф"], 67.89)
        self.assertEqual(totals["Убыток", "Офис"], 100)
        self.assertEqual(totals["Денежные средства", "Депозит"], -1000)


if __name__ == "__main__":
    unittest.main()
