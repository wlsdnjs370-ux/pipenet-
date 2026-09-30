"""Explicit H nozzle properties shared by review sizing and SDF/SLF export.

K is L/min/sqrt(bar). Required pressure is gauge bar. SLF pressure limits
are absolute Pa; the established file format uses standard atmosphere 101325 Pa.
Only marked H corrections are applied; original library definitions stay intact.
"""
from __future__ import annotations

from copy import deepcopy
import hashlib
import math
from pathlib import Path
import re
from typing import Mapping, Sequence
import xml.etree.ElementTree as ET

ATM_PA = 101325.0
POLICY = 'drawing_first_v1'


def corrected_spec(row: Mapping, base: Mapping) -> dict:
    """Resolve a corrected head against its existing SLF nozzle definition."""
    k=float(row.get('k_factor_si',float(base['k_si'])*60000*math.sqrt(100000)))
    pressure=float(row.get('required_pressure_bar',max(0,(base['min_p_pa']-ATM_PA)/100000)))
    maximum=max(0,(float(base['max_p_pa'])-ATM_PA)/100000)
    if not math.isfinite(k) or k<=0 or not math.isfinite(pressure) or pressure<0:
        raise ValueError('헤드 K 값·최소 압력을 확인하세요.')
    if maximum and pressure>maximum:
        raise ValueError('헤드 최소 압력이 라이브러리 최대 압력을 초과합니다.')
    return dict(k_factor_si=k,min_bar=pressure,max_bar=maximum)


def apply_nozzle_overrides(sdf_path: str | Path, slf_path: str | Path,
                           rows: Sequence[Mapping]) -> None:
    """Clone per-property SLF items and bind only the corresponding output heads."""
    targets=[r for r in rows if r.get('definition_policy')==POLICY]
    if not targets:
        return
    sdf_path,slf_path=Path(sdf_path),Path(slf_path)
    sdf,slf=ET.parse(sdf_path),ET.parse(slf_path)
    section=slf.getroot().find('.//Nozzle-section')
    if section is None:
        raise ValueError('SLF 노즐 라이브러리를 찾지 못했습니다.')
    definitions={n.findtext('Item-name'):n for n in section.findall('Nozzle-definition')}
    heads={n.get('label'):n for n in sdf.getroot().iter('Nozzle')}
    for row in targets:
        original=definitions.get(row.get('lib'))
        output=heads.get(str(row['label']))
        if original is None or output is None:
            raise ValueError(f"헤드 {row['label']}: 출력 노즐/라이브러리 대응 실패")
        spec=corrected_spec(row,dict(k_si=original.get('k-value'),
            min_p_pa=float(original.get('minimum-pressure','0')),
            max_p_pa=float(original.get('maximum-pressure','0'))))
        digest=hashlib.sha256(repr((row['lib'],spec)).encode()).hexdigest()[:14]
        name='H-HEAD-'+digest
        if name not in definitions:
            clone=deepcopy(original)
            clone.find('Item-name').text=name
            clone.set('k-value',format(spec['k_factor_si']/60000/math.sqrt(100000),'.12g'))
            clone.set('minimum-pressure',format(spec['min_bar']*100000+ATM_PA,'.12g'))
            section.append(clone);definitions[name]=clone
        item=output.find('Library-item')
        if item is None:
            raise ValueError('출력 노즐 라이브러리 참조가 없습니다.')
        item.text=name
    for path,tree in ((sdf_path,sdf),(slf_path,slf)):
        # Retain the existing root/doctype (SDF root varies by template version).
        original_text=path.read_text(encoding='utf-8')
        match=re.search(r'<!DOCTYPE[^>]*>',original_text)
        prefix='<?xml version="1.0" encoding="UTF-8"?>\n'+((match.group(0)+'\n') if match else '')
        path.write_text(prefix+ET.tostring(tree.getroot(),encoding='unicode'),encoding='utf-8')
