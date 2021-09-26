import os
import subprocess
import sys

import pandas as pd

from helpers import format_excel_worksheet


def potentially_inspect(dataframe, sheet, filename_with_xlsx, look_at=None):
    if look_at:
        if not look_at.endswith('.xlsx'):
            look_at += '.xlsx'
        if look_at == filename_with_xlsx:
            satisfied = False
            while not satisfied:
                try:
                    results_name = f'Inspection of {look_at}'
                    with pd.ExcelWriter(results_name) as writer:
                        dataframe.to_excel(writer, sheet_name=sheet, index=False)
                        format_excel_worksheet(writer.sheets[sheet], dataframe)

                    if sys.platform == "win32":
                        os.startfile(results_name)
                    else:
                        opener = "open" if sys.platform == "darwin" else "xdg-open"
                        subprocess.call([opener, results_name])
                except Exception:
                    pass
                else:
                    satisfied = True
