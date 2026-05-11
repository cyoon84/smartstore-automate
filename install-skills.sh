#!/bin/bash
# .cowork-skills/ 안의 모든 스킬 폴더를 ~/.claude/skills/ 로 설치 (또는 갱신)
# 한 번 실행하면 모든 새 Cowork 세션에서 자동으로 인식됨.
set -e

SRC_DIR="$(cd "$(dirname "$0")" && pwd)/.cowork-skills"
DST_DIR="$HOME/.claude/skills"

if [ ! -d "$SRC_DIR" ]; then
  echo "❌ $SRC_DIR 폴더가 없습니다."
  exit 1
fi

mkdir -p "$DST_DIR"

installed=()
for skill in "$SRC_DIR"/*/; do
  [ -d "$skill" ] || continue
  name="$(basename "$skill")"
  rm -rf "$DST_DIR/$name"
  # rsync 가 있으면 캐시 제외하면서 복사, 없으면 cp 후 정리
  if command -v rsync >/dev/null 2>&1; then
    rsync -a --exclude '__pycache__' --exclude '*.pyc' "$skill" "$DST_DIR/$name/"
  else
    cp -R "$skill" "$DST_DIR/$name"
    find "$DST_DIR/$name" -name __pycache__ -type d -exec rm -rf {} + 2>/dev/null || true
    find "$DST_DIR/$name" -name "*.pyc" -delete 2>/dev/null || true
  fi
  if [ -d "$DST_DIR/$name/scripts" ]; then
    chmod +x "$DST_DIR/$name/scripts/"*.py 2>/dev/null || true
  fi
  installed+=("$name")
done

if [ ${#installed[@]} -eq 0 ]; then
  echo "⚠️  설치할 스킬 폴더가 .cowork-skills/ 에 없습니다."
  exit 1
fi

echo "✅ 설치 완료: $DST_DIR"
for n in "${installed[@]}"; do
  echo "   • $n"
done
echo ""
echo "이제 모든 새 Cowork 세션에서 다음 표현으로 호출 가능:"
echo "  • '한미플로우 시작' / '한미택배 발송 처리'"
echo "  • '우체국 발송' / '우체국택배 처리'"
