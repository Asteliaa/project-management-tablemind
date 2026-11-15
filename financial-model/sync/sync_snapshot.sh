#!/usr/bin/env bash
# Снимок финансовой модели TableMind -> папка financial-model/ клона репозитория.
# Источник ТОЛЬКО читается (cp/sha256sum/grep); в источник ничего не записывается.
# Использование:
#   bash sync_snapshot.sh <SRC_ROOT> <DEST_FINANCIAL_MODEL_DIR>
# где SRC_ROOT       — корень локального проекта (на Windows в Git Bash: /d/TableMInd),
#     DEST_...DIR    — папка financial-model/ в СВЕЖЕМ клоне репозитория (не в D:\TableMInd).
set -euo pipefail
SRC="${1:?укажите SRC_ROOT}"; DEST="${2:?укажите DEST (financial-model/)}"
M="$SRC/TableMind_Model/TableMind"
[ -d "$M" ] || { echo "нет $M" >&2; exit 1; }
case "$(cd "$DEST" 2>/dev/null && pwd)" in "$(cd "$SRC" && pwd)"*) echo "DEST внутри SRC — запрещено" >&2; exit 1;; esac

# Содержимое DEST/{model,journals,scripts,governance} пересоздаётся целиком (файлы из старого снимка не копятся).
rm -rf "$DEST/model" "$DEST/journals" "$DEST/scripts" "$DEST/governance" "$DEST/audit"
mkdir -p "$DEST/model/_recalc" "$DEST/journals" "$DEST/scripts" "$DEST/governance" "$DEST/audit"

# 1. Рабочие копии книг (не оригиналы преподавателя, не «Модель (1)»)
find "$M" -maxdepth 1 -type f \( -name '*_TableMind.xlsx' -o -name '*_TableMind.docx' \) ! -name '.~lock*' -exec cp -p {} "$DEST/model/" \;
# 2. Результаты пересчёта (без lock- и tmp-файлов LibreOffice)
find "$M/_recalc" -maxdepth 1 -type f -name '*_TableMind.xlsx' ! -name '.~lock*' -exec cp -p {} "$DEST/model/_recalc/" \;
# 3. Журналы (без *.tmp)
find "$M/_журналы" -maxdepth 1 -type f ! -name '*.tmp' -exec cp -p {} "$DEST/journals/" \;
# 4. Скрипты расчёта и пересчёта
find "$M/_scripts" -maxdepth 1 -type f \( -name '*.py' -o -name '*.sh' \) -exec cp -p {} "$DEST/scripts/" \;
# 5. Управляющие документы
for f in CLAUDE.md PROJECT_MAP.md PROJECT_STATE.md OPEN_QUESTIONS.md; do cp -p "$SRC/$f" "$DEST/governance/$f"; done
cp -p "$SRC/TableMind_Model/README.md" "$DEST/governance/TableMind_Model_README.md"
# 6. Выписка: трасса индекса барьеров входа = раздел 5 PROJECT_STATE.md (от '## 5.' до '## 6.')
{
  echo "# Расхождения в исследованиях и трасса индекса барьеров входа"
  echo
  echo "> Выписка (дословно) из раздела 5 \`governance/PROJECT_STATE.md\` на момент снимка. Источник истины — PROJECT_STATE.md."
  echo
  awk '/^## 5\./{p=1;next} /^## 6\./{p=0} p' "$SRC/PROJECT_STATE.md"
} > "$DEST/audit/barrier_index_trace.md"

# 7. Манифест: SHA-256 всех файлов снимка (кроме README/SYNC/MANIFEST самой папки)
( cd "$DEST" && find model journals scripts governance audit -type f -print0 | sort -z | xargs -0 sha256sum > MANIFEST.sha256 )
date -u +"снимок (UTC): %Y-%m-%dT%H:%M:%SZ" > "$DEST/SNAPSHOT_TIME.txt"
echo "готово: $(wc -l < "$DEST/MANIFEST.sha256") файлов в манифесте"
