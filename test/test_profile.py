from unittest.mock import patch
from pathlib import Path

import numpy as np
import pandas as pd
import pytest

from hydraulic import config
from hydraulic.models import ProfileSector
from hydraulic.profile import ComputeCVQ, Morfostvor


def make_sector(sector_id, name, start, end, x, y):
    return ProfileSector(
        id=sector_id,
        name=name,
        start_point=start,
        end_point=end,
        roughness=0.035,
        slope=3.0,
        coord=(x, y),
    )


def test_use_h_instead_r_uses_depth_in_shezi_and_velocity():
    with patch.object(config, "USE_H_INSTEAD_R", True), patch.object(config, "SHEZI_TYPE", "Маннинга"):
        compute = ComputeCVQ(n=0.035, i=3.0, h=1.5, a=10.0, p=9.0, r=1.11)

    expected_shezi = (1 / 0.035) * 1.5 ** (1 / 6)
    assert compute.r == pytest.approx(1.5)
    assert compute.shezi == pytest.approx(expected_shezi)
    assert compute.v == pytest.approx(expected_shezi * np.sqrt(1.5 * 0.003))


def test_morfostvor_rejects_unknown_calculation_type():
    with patch.object(config, "CALC_TYPE", 999), pytest.raises(SystemExit):
        Morfostvor()


def test_probability_validation_rejects_empty_probability_name():
    morfostvor = Morfostvor()
    morfostvor.title = "Тестовый створ"
    morfostvor.probability = [["", 10.0]]

    with pytest.raises(ValueError, match="не может быть пустым"):
        morfostvor._Morfostvor__check_probability()


def test_probability_validation_rejects_non_numeric_flow():
    morfostvor = Morfostvor()
    morfostvor.title = "Тестовый створ"
    morfostvor.probability = [[1, "не расход"]]

    with pytest.raises(ValueError, match="Некорректное значение"):
        morfostvor._Morfostvor__check_probability()


def test_read_xls_loads_profile_sectors_and_probability_table():
    filename = Path(__file__).parents[1] / "example" / "example_profile.xlsx"
    morfostvor = Morfostvor()

    morfostvor.read_xls(filename, page=0)

    assert len(morfostvor.x) == len(morfostvor.y) == 6
    assert len(morfostvor.sectors) == 3
    assert len(morfostvor.probability) == 5
    assert morfostvor.sectors[0].roughness == pytest.approx(0.04)
    assert morfostvor.sectors[0].slope == pytest.approx(1.5)


def test_read_xls_rejects_missing_file():
    morfostvor = Morfostvor()

    with pytest.raises(SystemExit) as error:
        morfostvor.read_xls("missing-profile.xlsx")

    assert error.value.code == 33


def test_read_xls_rejects_page_out_of_range():
    filename = Path(__file__).parents[1] / "example" / "example_profile.xlsx"
    morfostvor = Morfostvor()

    with pytest.raises(SystemExit) as error:
        morfostvor.read_xls(filename, page=999)

    assert error.value.code == 34


def test_get_q_max_returns_highest_flow_and_its_probability():
    morfostvor = Morfostvor()
    morfostvor.probability = [[1, 120.0], [5, 80.0], [10, 150.0]]

    assert morfostvor.get_q_max() == (10, 150.0)


def test_get_min_sector_returns_sector_with_lowest_bed_elevation():
    left = make_sector(1, "левый", 0, 1, [0.0, 1.0], [3.0, 2.0])
    channel = make_sector(2, "русло", 1, 2, [1.0, 2.0, 3.0], [2.0, 0.0, 2.0])
    morfostvor = Morfostvor(sectors=[left, channel])

    assert morfostvor.get_min_sector() == (1, channel)


def test_get_prob_table_interpolates_results_for_requested_flows():
    morfostvor = Morfostvor(probability=[[5, 15.0]])
    curve = pd.DataFrame(
        {"Q": [10.0, 20.0], "V": [2.0, 4.0], "F": [4.0, 8.0]},
        index=[1.0, 2.0],
    )

    result = morfostvor.get_prob_table(curve)

    assert result.loc[0, "P"] == 5
    assert result.loc[0, "Q"] == 15.0
    assert result.loc[0, "H"] == pytest.approx(1.5)
    assert result.loc[0, "V"] == pytest.approx(3.0)
    assert result.loc[0, "F"] == pytest.approx(6.0)


def test_get_topography_table_maps_points_to_sectors():
    first = make_sector(1, "левый", 0, 1, [0.0, 1.0], [2.0, 1.0])
    second = make_sector(2, "русло", 1, 2, [1.0, 2.0], [1.0, 0.0])
    morfostvor = Morfostvor(x=[0.0, 1.0, 2.0], y=[2.0, 1.0, 0.0], sectors=[first, second])

    result = morfostvor.get_topography_table()

    assert result["x"].tolist() == [0.0, 1.0, 2.0]
    assert result["sectors"].tolist() == ["левый", "русло", "русло"]
    assert result["roughness"].tolist() == [0.035, 0.035, 0.035]
    assert result["slope"].tolist() == [3.0, 3.0, 3.0]


def test_calc_by_sectors_records_section_parameters_and_flow():
    with patch.object(config, "SHEZI_TYPE", "Маннинга"), patch.object(config, "CALC_TYPE", 1):
        sector = make_sector(1, "русло", 0, 2, [0.0, 1.0, 2.0], [1.0, 0.0, 1.0])
        morfostvor = Morfostvor(sectors=[sector])
        columns = pd.Index(["Участок", "УВ", "F", "B", "W", "Hср", "Hмакс", "R", "V", "Q", "Shezi"])
        calculated_flows = []
        result = morfostvor._calc_by_sectors(0.5, columns, pd.DataFrame(columns=columns), calculated_flows)

    assert len(result) == 1
    assert result.loc[0, "Участок"] == "русло"
    assert result.loc[0, "УВ"] == 0.5
    assert result.loc[0, "F"] > 0
    assert result.loc[0, "V"] > 0
    assert result.loc[0, "Q"] == pytest.approx(result.loc[0, "F"] * result.loc[0, "V"])
    assert calculated_flows == pytest.approx([result.loc[0, "Q"]])
