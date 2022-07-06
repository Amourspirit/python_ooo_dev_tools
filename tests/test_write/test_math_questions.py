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
from ooodev.office.write import Write
from ooodev.utils.gui import GUI
from ooodev.utils.date_time_util import DateUtil


def test_math_questions(loader, tmp_path_fn):
    visible = False
    delay = 0 # 2_000
    doc = Write.create_doc(loader)
    fnm_pdf = Path(tmp_path_fn, "mathQuestions.pdf")
    try:
        if visible:
            GUI.set_visible(visible, doc)
        # lock screen updates for first test
        cursor = Write.get_cursor(doc)
        Write.append_para(cursor, "Math Questions")
        Write.style_prev_paragraph(cursor, "Heading 1")
        
        Write.append_para(cursor, "Solve the following formulae for x:\n")

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
                Write.add_formula(cursor, formula.replace("[", '{').replace(']', '}'))
                Write.end_paragraph(cursor)
        
        Write.append_para(cursor, f"Timestamp: {DateUtil.time_stamp()}")
        Lo.delay(delay)
        Write.save_doc(text_doc=doc, fnm=fnm_pdf)
        assert fnm_pdf.exists()
    finally:
        Lo.close_doc(doc, False)


