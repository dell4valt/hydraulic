from types import SimpleNamespace

import numpy as np
import pytest

from hydraulic import lib
from hydraulic.models import ProfileSector, SituationSector, WaterSection


def make_sector(name="русло", x=None, y=None, start=0, end=2):
    return ProfileSector(
        id=1,
        name=name,
        start_point=start,
        end_point=end,
        roughness=0.035,
        slope=3.0,
        coord=(x or [0.0, 1.0, 2.0], y or [1.0, 0.0, 1.0]),
    )


def test_profile_sector_calculates_length_and_known_color():
    sector = make_sector()

    assert sector.length == 2.0
    assert sector.color == [0, 0.5, 1]


@pytest.mark.parametrize("field,value", [("id", "1"), ("roughness", "0.035"), ("slope", "3")])
def test_profile_sector_rejects_invalid_types(field, value):
    params = {
        "id": 1,
        "name": "русло",
        "start_point": 0,
        "end_point": 2,
        "roughness": 0.035,
        "slope": 3.0,
        "coord": ([0.0, 1.0, 2.0], [1.0, 0.0, 1.0]),
    }
    params[field] = value

    with pytest.raises(TypeError):
        ProfileSector(**params)


@pytest.mark.parametrize(
    "surface,expected",
    [("луг", "honeydew"), ("бетон", "gainsboro"), ("поле", "burlywood"), ("неизвестно", "white")],
)
def test_situation_sector_maps_surface_to_color(surface, expected):
    assert SituationSector(1, surface, 0, 1).get_color() == expected


def test_water_section_calculates_trapezoidal_section_parameters():
    section = WaterSection([0.0, 1.0, 2.0], [1.0, 0.0, 1.0], water_level=1.0)

    assert section.width == pytest.approx(2.0)
    assert section.area == pytest.approx(1.0)
    assert section.average_depth == pytest.approx(0.5)
    assert section.max_depth == pytest.approx(1.0)
    assert section.wet_perimeter == pytest.approx(round(2 * np.sqrt(2), 3))
    assert section.r_hydraulic == pytest.approx(round(1 / round(2 * np.sqrt(2), 3), 3))


def test_water_section_combines_disconnected_pools():
    section = WaterSection(
        [0.0, 1.0, 2.0, 3.0, 4.0],
        [2.0, 0.0, 2.0, 0.0, 2.0],
        water_level=1.0,
    )

    assert len(section.segments) == 2
    assert section.width == pytest.approx(2.0)
    assert section.area == pytest.approx(1.0)
    assert section.average_depth == pytest.approx(0.5)


def test_water_section_filters_segments_by_start_point():
    section = WaterSection(
        [0.0, 1.0, 2.0, 3.0, 4.0],
        [2.0, 0.0, 2.0, 0.0, 2.0],
        water_level=1.0,
        start_point=3.0,
    )

    assert len(section.segments) == 1
    assert section.width == pytest.approx(1.0)
    assert section.area == pytest.approx(0.5)


def test_water_section_rejects_level_below_profile():
    with pytest.raises(ValueError, match="ниже низшей точки"):
        WaterSection([0.0, 1.0], [1.0, 0.0], water_level=-0.1)


def test_water_section_rejects_start_point_outside_water_segment():
    with pytest.raises(ValueError, match="не попадает ни в один сегмент"):
        WaterSection([0.0, 1.0, 2.0], [1.0, 0.0, 1.0], water_level=1.0, start_point=99.0)


def test_get_water_sections_returns_sections_for_non_overflow_mode():
    sector = make_sector()
    morfostvor = SimpleNamespace(sectors=[sector], get_min_sector=lambda: (0, sector))

    sections, sectors = lib.get_water_sections(morfostvor, water_level=1.0)

    assert len(sections) == 1
    assert sectors == [sector]
    assert sections[0].area == pytest.approx(1.0)


def test_get_water_sections_returns_empty_for_dry_sectors():
    sector = make_sector(y=[2.0, 1.5, 2.0])
    morfostvor = SimpleNamespace(sectors=[sector], get_min_sector=lambda: (0, sector))

    sections, sectors = lib.get_water_sections(morfostvor, water_level=1.0)

    assert sections == []
    assert sectors == []


def test_get_water_sections_starts_from_lowest_sector_in_overflow_mode():
    sector = make_sector()
    morfostvor = SimpleNamespace(sectors=[sector], get_min_sector=lambda: (0, sector))

    sections, sectors = lib.get_water_sections(morfostvor, water_level=1.0, overflow=True)

    assert len(sections) == 1
    assert sectors == [sector]
    assert sections[0].area == pytest.approx(1.0)
