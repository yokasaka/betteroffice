import io
import json
import re
import zipfile

import pytest

import betteroffice_pptx as bo
from conftest import ROOT

JAPANESE = "日本語の見出し"
NOTO_JP = ROOT / "packages" / "fonts-cjk" / "assets" / "NotoSansJP-Regular.otf"


@pytest.fixture(scope="module")
def japanese_deck(sample_bytes) -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(sample_bytes)) as source, zipfile.ZipFile(buffer, "w") as target:
        for part in source.infolist():
            data = source.read(part.filename)
            if part.filename == "ppt/slides/slide1.xml":
                data, replaced = re.subn(
                    rb"<a:t>[^<]+</a:t>", f"<a:t>{JAPANESE}</a:t>".encode(), data, count=1
                )
                assert replaced == 1
            target.writestr(part, data)
    return buffer.getvalue()


@pytest.fixture(scope="module")
def noto_jp() -> bytes:
    if not NOTO_JP.exists():
        pytest.skip(f"fixture missing: {NOTO_JP}")
    return NOTO_JP.read_bytes()


def families(deck: bo.Presentation) -> str:
    return json.dumps(deck.render_slide(0).to_dict(), ensure_ascii=False)


def test_a_fallback_font_draws_glyphs_the_run_face_lacks(japanese_deck, font_bytes, noto_jp):
    deck = bo.Presentation.open(japanese_deck)
    deck.register_font("Arial", font_bytes)
    without = deck.render_png(0).bytes
    assert "Noto Sans JP" not in families(deck)

    assert isinstance(deck.register_fallback_font("Noto Sans JP", noto_jp), int)

    assert "Noto Sans JP" in families(deck)
    assert deck.render_png(0).bytes != without


def test_a_fallback_font_is_not_written_into_the_saved_deck(japanese_deck, font_bytes, noto_jp):
    deck = bo.Presentation.open(japanese_deck)
    deck.register_font("Arial", font_bytes)
    deck.register_fallback_font("Noto Sans JP", noto_jp)
    deck.render_png(0)
    assert b"Noto Sans JP" not in zipfile.ZipFile(io.BytesIO(deck.save())).read("ppt/slides/slide1.xml")


def test_a_fallback_font_rejects_bytes_that_are_not_a_font(japanese_deck):
    deck = bo.Presentation.open(japanese_deck)
    with pytest.raises(bo.RenderError):
        deck.register_fallback_font("Broken", b"not a font")
