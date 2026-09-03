"""Check the actual Windows executable with synthetic Excel input."""
import os
from pathlib import Path
import shutil
import subprocess
import sys
from tempfile import TemporaryDirectory

import pandas as pd


def main():
    distribution = Path(sys.argv[1]).resolve()
    with TemporaryDirectory(prefix="buhuchet-exe-") as temp:
        folder = Path(temp) / "Бухучёт проверка"
        shutil.copytree(distribution, folder)
        columns = ["Документ", "Дата операции", "Корреспондент",
                   "Оборот Дт", "Оборот Кт", "Назначение платежа"]
        rows = [
            columns, [None] * len(columns),
            [1, "15.09.2026", "Океан Сервис", 100, None, "оплата за техобслуж"],
            [2, "15.09.2026", "Тестовый банк", None, 1000,
             "возврат средств согласно депозитн договора"],
            [3, "15.09.2026", "Тест", 50, None, "Оплата штрафа"],
            [4, "15.09.2026", "Тест", None, 75, "Штраф по договору"],
            ["ИТОГО", None, None, 150, 1075, None],
        ]
        pd.DataFrame(rows).to_excel(folder / "Выписка.xlsx", header=False, index=False)
        answers = "2026\nсентябрь\n1\n1\nВыписка.xlsx\nда\n-1\n"
        result = subprocess.run(
            [str(folder / "Buhuchet.exe")], input=answers,
            text=True, encoding="utf-8", capture_output=True,
            cwd=folder, env={**os.environ, "PYTHONUTF8": "1"}, timeout=60,
        )
        if result.returncode != 0 or "Ошибка" in result.stdout:
            raise RuntimeError(result.stdout + result.stderr)
        output = pd.read_excel(folder / "2026/сентябрь/Выписка_с_типами.xlsx")
        expected = ["ПО", "депозит возврат", "штрафы", "штрафы"]
        if output["Тип операции"].tolist() != expected:
            raise AssertionError(output["Тип операции"].tolist())
        print("Packaged executable: startup, Cyrillic paths, Excel import/classification/export OK")


if __name__ == "__main__":
    main()
