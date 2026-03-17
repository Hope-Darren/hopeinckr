#!/bin/bash
set -euo pipefail

# 웹 환경(Claude Code on the web)에서만 실행
if [ "${CLAUDE_CODE_REMOTE:-}" != "true" ]; then
  exit 0
fi

# Python 의존성 설치 (openpyxl - 엑셀 파일 분석용)
if ! python3 -c "import openpyxl" 2>/dev/null; then
  pip3 install openpyxl --quiet
fi

# 간단한 HTTP 서버 제공을 위한 확인
echo "재료수불 관리 환경 준비 완료" >&2
