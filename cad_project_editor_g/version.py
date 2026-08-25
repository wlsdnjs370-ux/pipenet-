"""
앱 버전 단일 진실 공급원 (SSOT)

버전 업 시 APP_VERSION 한 줄만 수정하면
모든 참조 위치(스플래시, About, PDF 푸터, 설치파일)에 자동 반영됩니다.

PROJECT_SCHEMA_VERSION은 프로젝트 파일 포맷 식별자입니다.
앱 버전과 독립적으로 관리하며, 파일 포맷 구조가 바뀔 때만 올립니다.
"""

APP_VERSION = "5.5"
# 프로젝트 .kfp 파일 포맷 식별자. 파일 포맷 구조가 바뀔 때만 올린다.
# major 4 = 단일 nodes_meta_runtime 블록 포맷(저장포맷 마이그레이션 완료).
# 로더는 major 3(레거시 4-블록)도 영구 지원하므로 기존 사용자 .kfp는 그대로 열린다.
PROJECT_SCHEMA_VERSION = "4.0-NFPA13-EQL"
# Water-only projects intentionally remain on the exact 4.0 format.  A project
# that explicitly contains an antifreeze analysis case uses major 5 so older
# major-4 builds reject it instead of silently dropping the case on re-save.
ANTIFREEZE_PROJECT_SCHEMA_VERSION = "5.0-NFPA13-AF-DW"
