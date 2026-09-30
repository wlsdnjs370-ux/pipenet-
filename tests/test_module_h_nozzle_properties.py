"""H batch head properties survive both standalone and merged SDF output."""
import math
import xml.etree.ElementTree as ET
import pytest
from test_module_f_sizing import tables
from routes.module_f.hydraulic_sizing import prepare


@pytest.mark.parametrize('writer',['plan','merged'])
def test_custom_nozzle_binding_roundtrip(tmp_path,writer):
    tbl=tables()
    tbl.nozzles[0].update(definition_policy='drawing_first_v1',k_factor_si=90,required_pressure_bar=1.5)
    net,_,_=prepare(tbl,{})
    assert net.nozzles[0].k_lpm_sqrt_bar==90
    assert net.nozzles[0].min_pressure_bar==1.5
    path=tmp_path/'heads.sdf'
    if writer=='plan':
        from services.cad_import.design.emit import emit_design_sdf
        emit_design_sdf(tbl,path)
    else:
        from remote30_prototype import emit_sdf
        emit_sdf(tbl,path)
    head=next(ET.parse(path).getroot().iter('Nozzle'))
    library=ET.parse(path.with_suffix('.slf')).getroot()
    definition=next(n for n in library.iter('Nozzle-definition') if n.findtext('Item-name')==head.findtext('Library-item'))
    assert head.findtext('Library-item').startswith('H-HEAD-')
    assert float(definition.get('k-value'))*60000*math.sqrt(100000)==pytest.approx(90)
    assert (float(definition.get('minimum-pressure'))-101325)/100000==1.5
    assert any(n.findtext('Item-name')=='SP-HEAD' for n in library.iter('Nozzle-definition'))
