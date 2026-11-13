from types import SimpleNamespace
from unittest.mock import patch

import pytest

from hydraulic import config
from hydraulic.report import generate_morfostvor_report


class ReportConstructionReached(Exception):
    pass


def test_graphics_titles_text_reads_existing_lowercase_graph_attributes():
    def graph():
        return SimpleNamespace(fig=object(), _ax_title_text="Заголовок")

    morfostvor = SimpleNamespace(
        fig_profile=SimpleNamespace(fig=object(), morfostvor=SimpleNamespace(title="Профиль")),
        fig_qh=graph(),
        fig_qhv=graph(),
        fig_qv=graph(),
        fig_vh=graph(),
        fig_qf=graph(),
        fig_fh=graph(),
        fig_qwvh=graph(),
    )

    with (
        patch.object(config, "GRAPHICS_TITLES_TEXT", True),
        patch("hydraulic.report.Report", side_effect=ReportConstructionReached),
        pytest.raises(ReportConstructionReached),
    ):
        generate_morfostvor_report(morfostvor, "unused.docx")
