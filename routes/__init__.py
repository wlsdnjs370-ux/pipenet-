# -*- coding: utf-8 -*-
"""도메인/기능 축 라우트 모듈 패키지.

각 모듈은 `register(app, ...)` 함수를 노출한다. `대조 서버.py` 는 앱과
공유 헬퍼를 모두 정의한 뒤, 파일 끝에서 각 도메인의 `register()` 를 호출해
`@app.route` 를 실제 app 에 등록한다. Blueprint 대신 이 패턴을 쓰는 이유는
엔드포인트명(`login_page` 등)을 접두사 없이 그대로 보존해 `url_for`·템플릿·
route 인벤토리가 리팩토링 전후로 바이트 동일하게 유지되기 때문이다.
"""
import os
import traceback


def traceback_for_client() -> str | None:
    """활성 예외의 traceback — 외부 노출 환경에서는 None.

    fncadnet.com 으로 열려 있어 traceback 본문(파일 경로·업로드 파일명·코드
    스니펫)이 그대로 나가면 안 된다. 호출부는 반환값이 None 이면 키 자체를
    응답에서 빼고, 본문은 서버 로그에만 남긴다. EXPOSE_TRACEBACK=1 로 국소 복원.
    """
    if os.environ.get("EXPOSE_TRACEBACK") == "1":
        return traceback.format_exc()[-2000:]
    return None
