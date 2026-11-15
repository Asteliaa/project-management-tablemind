#!/bin/bash
# Пересчёт рабочих книг в LibreOffice (headless). Рабочие копии НЕ изменяются: результат копируется в ..\_recalc.
# Конвертация идёт во временную папку вне диска D (иначе LibreOffice оставляет .~lock и lu*.tmp, которые нельзя удалить без разрешения).
# Использование: bash recalc_libreoffice.sh "<имя_книги.xlsx>" [...]   (без аргументов - все .xlsx рабочей папки)
cd "$(dirname "$0")/.." || exit 1
T="$(mktemp -d)"; files=("$@"); [ ${#files[@]} -eq 0 ] && files=(*.xlsx)
for f in "${files[@]}"; do
  cp "$f" "$T/in.xlsx"
  if timeout 170 soffice --headless --calc --convert-to xlsx:"Calc MS Excel 2007 XML" --outdir "$T/out" "$T/in.xlsx" >/dev/null 2>&1; then cp -f "$T/out/in.xlsx" "_recalc/$f" && echo "ok  $f"; else echo "FAIL $f"; fi
  rm -rf "$T/out"
done
rm -rf "$T"
