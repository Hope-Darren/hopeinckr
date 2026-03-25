#!/bin/bash
set -euo pipefail

# 웹 환경(Claude Code on the web)에서만 실행
if [ "${CLAUDE_CODE_REMOTE:-}" != "true" ]; then
  exit 0
fi

# Python 의존성 설치 (엑셀 파일 분석용)
pip3 install openpyxl --quiet

echo "재료수불 관리 환경 준비 완료" >&2
