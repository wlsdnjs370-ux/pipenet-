"""Duration units, failed-trial safeguards, and optional water demand."""
from dataclasses import replace
import pytest
from src.pipenet_converter.hydraulics.models import Settings
from src.pipenet_converter.hydraulics.water_supply import water_demand


@pytest.mark.parametrize('duration,expected', [(20, 30), (40, 60), (60, 90)])
def test_scenario_volume(duration, expected):
    result = water_demand(1500, duration, feasible=True)
    assert result['required_volume_m3'] == expected
    assert result['usable'] and not result['regulatory_compliance']


def test_failed_or_missing_duration_not_a_required_spec():
    assert water_demand(500, 20, feasible=False)['required_volume_m3'] is None
    assert water_demand(500, 20, feasible=False)['trial_volume_m3'] == 10
    assert not water_demand(500, None, feasible=True)['usable']


@pytest.mark.parametrize('duration', [0, -1, float('nan'), float('inf'), True, '20'])
def test_invalid_duration(duration):
    with pytest.raises(ValueError, match='지속시간'):
        replace(Settings(), duration_minutes=duration).validate()
