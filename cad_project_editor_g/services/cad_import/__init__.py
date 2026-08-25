"""CAD 임포트 엔진. UI·i18n·PySide 없음.

화면은 DTO와 파일만 주고받는다. 찍기·손질·파이프라인 1~6·변환 엔진.
입력은 DXF. DWG 변환은 넣지 않는다.

공개 API는 하위 모듈에서 직접 가져온다.

    from services.cad_import.pick import PickSession
    from services.cad_import.edit import EditSession
    from services.cad_import.convert.engine import convert_to_kfp
    from services.cad_import.convert.preflight import preflight_kfp_convert
    from services.cad_import.dto import default_dto
    from services.cad_import.pipeline.expand import stage1_body
    from services.cad_import.pipeline.flow import pipeline
"""
