"""
Write 10 forumlae into a new Text doc and save as a pdf file.
"""
import pytest
import random
from pathlib import Path

if __name__ == "__main__":
    pytest.main([__file__])
import uno

from ooodev.utils.lo import Lo
from ooodev.write import Write
from ooodev.write import WriteDoc
from ooodev.utils.date_time_util import DateUtil


def test_math_questions(loader, tmp_path_fn):
    visible = False
    delay = 0  # 2_000
    doc = WriteDoc(Write.create_doc(loader))
    fnm_pdf = Path(tmp_path_fn, "mathQuestions.pdf")
    try:
        if visible:
            doc.set_visible()
        # lock screen updates for first test
        cursor = doc.get_cursor()
        cursor.append_para("Math Questions")
        cursor.style_prev_paragraph("Heading 1")

        cursor.append_para("Solve the following formulae for x:\n")

        # lock screen updating
        with Lo.ControllerLock():
            for _ in range(10):  # generate 10 random formulae
                iA = random.randint(0, 7) + 2
                iB = random.randint(0, 7) + 2
                iC = random.randint(0, 8) + 1
                iD = random.randint(0, 7) + 2
                iE = random.randint(0, 8) + 1
                iF1 = random.randint(0, 7) + 2

                choice = random.randint(0, 2)

                if choice == 0:
                    formula = f"[[[sqrt[{iA}x]] over {iB}] + [{iC} over {iD}]=[{iE} over {iF1} ]]"
                elif choice == 1:
                    formula = f"[[[{iA}x] over {iB}] + [{iC} over {iD}]=[{iE} over {iF1}]]"
                else:
                    formula = f"[{iA}x + {iB} = {iC}]"
                cursor.add_formula(formula.replace("[", "{").replace("]", "}"))
                cursor.end_paragraph()

        cursor.append_para(f"Timestamp: {DateUtil.time_stamp()}")
        Lo.delay(delay)
        doc.save_doc(fnm=fnm_pdf)
        assert fnm_pdf.exists()
    finally:
        doc.close_doc()
