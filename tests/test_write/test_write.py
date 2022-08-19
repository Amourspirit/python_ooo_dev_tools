import os
import pytest
from pathlib import Path
# from ooodev.office.write import Write
if __name__ == "__main__":
    pytest.main([__file__])

def test_writer_lines(loader, fix_writer_path, tmp_path_fn):
    from ooodev.utils.lo import Lo
    from ooodev.office.write import Write
    from ooodev.utils.gui import GUI
    test_doc = fix_writer_path("hello_sunny.odt")
    assert loader is not None
    doc = Write.create_doc(loader)
    # if doc is None:
    #     Lo.close_office()
    assert doc is not None
    GUI.set_visible(is_visible=True, odoc=doc)
    
    lines = ['Hello World!', 'More sunny days will appear. Blue skies are fantastic.', 'Reportedly the distance from Earth to the Sun and equal to 150 million kilometres (93 million miles) or 8.3 light minutes.',
            'The actual distance from Earth to the Sun varies by about 3% as Earth orbits the Sun, from a maximum (aphelion) to a minimum (perihelion) and back again once each year.',
            "The astronomical unit was originally conceived as the average of Earth's aphelion and perihelion; however, since 2012 it has been defined as exactly 149597870700 m."]
    cursor = Write.get_cursor(doc)
    for line in lines:
        Write.append_para(cursor=cursor, text=line)
    Lo.delay(100)
    fnm = Path(tmp_path_fn, 'example.odt')
    Write.save_doc(text_doc=doc, fnm=str(fnm))
    Lo.close_doc(doc, False)
    tmp_doc = Write.open_doc(fnm=str(fnm), loader=loader)
    cursor = Write.get_cursor(tmp_doc)
    text = Write.get_all_text(cursor=cursor).rstrip()
    Lo.close_doc(tmp_doc, False)
    txt_lines = os.linesep.join(lines).rstrip()
    assert txt_lines == text, "Origin document text does not match temp document text."