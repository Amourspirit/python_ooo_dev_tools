from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])
from pathlib import Path
from ooodev.adapter.util.path_settings_properties_partial import PathSettingsPropertiesPartial
from ooodev.utils.partial.auto_attribute import AutoAttribute
from ooodev.utils.string.str_list import StrList


def test_path_settings(loader) -> None:
    obj = AutoAttribute()
    ps = PathSettingsPropertiesPartial(obj)
    first = "/home/user/me/.config/LibreOffice/4/value1"
    second = "/home/user/me/.config/LibreOffice/4/value2"
    expected1 = StrList.from_str("file:///home/user/me/.config/LibreOffice/4/value1")
    expected2 = StrList.from_str(
        "file:///home/user/me/.config/LibreOffice/4/value1;file:///home/user/me/.config/LibreOffice/4/value2"
    )

    p1 = Path(first)
    p2 = Path(second)
    p_comb = (p1, p2)

    ps.addin = first
    assert ps.addin == expected1
    ps.addin = p1
    assert ps.addin == expected1
    assert ps.addin.to_string() == str(expected1)

    ps.auto_correct = first
    assert ps.auto_correct == expected1
    ps.auto_correct = p1
    assert ps.auto_correct == expected1
    ps.auto_correct = p_comb
    assert ps.auto_correct == expected2
    assert ps.auto_correct.to_string() == str(expected2)

    ps.auto_text = first
    assert ps.auto_text == expected1
    ps.auto_text = p1
    assert ps.auto_text == expected1

    ps.backup = first
    assert ps.backup == expected1
    ps.backup = p1
    assert ps.backup == expected1

    ps.base_path_share_layer = first
    assert ps.base_path_share_layer == expected1
    ps.base_path_share_layer = p1
    assert ps.base_path_share_layer == expected1

    ps.base_path_user_layer = first
    assert ps.base_path_user_layer == expected1
    ps.base_path_user_layer = p1
    assert ps.base_path_user_layer == expected1

    ps.basic = first
    assert ps.basic == expected1
    ps.basic = p1
    assert ps.basic == expected1
    ps.basic = p_comb
    assert ps.basic == expected2

    ps.bitmap = first
    assert ps.bitmap == expected1
    ps.bitmap = p1
    assert ps.bitmap == expected1

    ps.dictionary = first
    assert ps.dictionary == expected1
    ps.dictionary = p1
    assert ps.dictionary == expected1

    ps.favorite = first
    assert ps.favorite == expected1
    ps.favorite = p1
    assert ps.favorite == expected1

    ps.filter = first
    assert ps.filter == expected1
    ps.filter = p1
    assert ps.filter == expected1

    ps.gallery = first
    assert ps.gallery == expected1
    ps.gallery = p1
    assert ps.gallery == expected1
    ps.gallery = p_comb
    assert ps.gallery == expected2

    ps.graphic = first
    assert ps.graphic == expected1
    ps.graphic = p1
    assert ps.graphic == expected1

    ps.help = first
    assert ps.help == expected1
    ps.help = p1
    assert ps.help == expected1

    ps.linguistic = first
    assert ps.linguistic == expected1
    ps.linguistic = p1
    assert ps.linguistic == expected1
    ps.linguistic = p_comb
    assert ps.linguistic == expected2

    ps.palette = first
    assert ps.palette == expected1
    ps.palette = p1
    assert ps.palette == expected1
    ps.palette = p_comb
    assert ps.palette == expected2

    ps.plugin = first
    assert ps.plugin == expected1
    ps.plugin = p1
    assert ps.plugin == expected1
    ps.plugin = p_comb
    assert ps.plugin == expected2

    ps.storage = first
    assert ps.storage == expected1
    ps.storage = p1
    assert ps.storage == expected1

    ps.temp = first
    assert ps.temp == expected1
    ps.temp = p1
    assert ps.temp == expected1

    ps.template = first
    assert ps.template == expected1
    ps.template = p1
    assert ps.template == expected1
    ps.template = p_comb
    assert ps.template == expected2

    ps.ui_config = first
    assert ps.ui_config == expected1
    ps.ui_config = p1
    assert ps.ui_config == expected1
    ps.ui_config = p_comb
    assert ps.ui_config == expected2

    ps.user_config = first
    assert ps.user_config == expected1
    ps.user_config = p1
    assert ps.user_config == expected1

    ps.user_dictionary = first
    assert ps.user_dictionary == expected1
    ps.user_dictionary = p1
    assert ps.user_dictionary == expected1

    ps.work = first
    assert ps.work == expected1
    ps.work = p1
    assert ps.work == expected1
