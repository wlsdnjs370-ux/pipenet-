# -*- coding: utf-8 -*-
import sys

sys.path.insert(0, ".")
from core.design.schema import floor_index, BuildingDraft  # noqa: E402

for s in ("1F", "2F", "B1F", "지하2층", "지상5층", "옥탑", "지하주차장 1층", "", "x"):
    print(repr(s), floor_index(s))

d = BuildingDraft.from_dict({"source": {"floors": [
    {"label": "B2F", "height_mm": 3000}, {"label": "B1F", "height_mm": 4000},
    {"label": "1F", "height_mm": 4500}, {"label": "2F", "height_mm": 3200},
    {"label": "3F"}, {"label": "4F", "height_mm": 3200}, {"label": "옥상"}]}})
print(d.floor_elevations())
