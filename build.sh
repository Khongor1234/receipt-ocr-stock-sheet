#!/bin/sh
# Builds ONE file with every script in it, for a single paste into Apps Script.
# The explanations (comment lines) stay in the .gs files of this folder and are left out of
# the pasted file, which is about a fifth shorter that way.
#
# Usage: sh build.sh          the main workbook   -> neg-file/all-in-one.gs
#        sh build.sh ocr      receipt OCR (its own Apps Script project) -> neg-file/ocr-all-in-one.gs
#
# profile.gs (or profiles/<name>.gs) comes first and config.gs second: the other files read
# their values while they load.
set -e
cd "$(dirname "$0")"
mkdir -p neg-file
profile_name=${1:-main}
workbook="config year-books year-rollover layout-migration gift-migration year-setup sheet-guard year-sync price-sync instant-sync tab-order period-headers calc-formulas charts dividers protect-sheets trim-rows header-guard view-columns name-fixes product-list stock-sheets count-sheets row-copies carry-products"
case "$profile_name" in
  main)
    profile_file=profile.gs
    out=neg-file/all-in-one.gs
    files="$workbook tosgon-names link-common price-link official-names transfer-link"
    title="бараа орлого зарлага"
    ;;
  ocr)
    profile_file=profile.gs
    out=neg-file/ocr-all-in-one.gs
    files="config year-books ocr/ocr-config ocr/ocr-validate ocr/ocr-review ocr/ocr-gemini ocr/ocr-catalog ocr/ocr-files ocr/ocr-sheet ocr/ocr-post ocr/ocr-setup ocr/ocr-watch ocr/ocr-mirror ocr/ocr-orphans ocr/ocr-spot ocr/ocr-instant ocr/ocr-split ocr/ocr-reread ocr/ocr-photo-links ocr/ocr-check ocr/ocr-elsewhere ocr/ocr-reconcile ocr/ocr-layout ocr/ocr-calendar ocr/ocr-calendar-rebuild ocr/ocr-suggest ocr/ocr-backfill ocr/ocr-report ocr/ocr-count-read ocr/ocr-count-report ocr/ocr-twins ocr/ocr-main dividers row-copies ocr/ocr-new-product ocr/ocr-relink"
    title="Баримтын OCR"
    ;;
  *)
    echo "unknown profile: $profile_name (main or ocr)" >&2
    exit 1
    ;;
esac
version=$(date +%Y%m%d-%H%M%S)
# Drops /* ... */ blocks and lines that hold a comment only; code lines are left as they are.
strip_comments() {
  awk '
    inblock { if ($0 ~ /\*\//) inblock = 0; next }
    /^[[:space:]]*\/\*/ { if ($0 !~ /\*\//) inblock = 1; next }
    /^[[:space:]]*\/\// { next }
    /^[[:space:]]*$/ { if (blank) next; blank = 1; print; next }
    { blank = 0; print }
  ' "$1"
}
{
  printf '// %s: бүх script нэг файлд (build %s).\n' "$title" "$version"
  printf '// DRY_RUN = true: юу ч өөрчлөхгүй, зөвхөн log-д бичнэ. false: жинхэнээр хийнэ.\n'
  printf '// Тайлбартай эх код: config.gs, year-setup.gs, ... тус тусын файлд.\n'
  printf '\n// ==================== %s ====================\n\n' "$profile_file"
  strip_comments "$profile_file"
  for f in $files; do
    printf '\n// ==================== %s.gs ====================\n\n' "$f"
    strip_comments "$f.gs"
  done
} > "$out"
# Only the workbook builds have setupYear and its version stamp.
if grep -q "const SETUP_VERSION = 'dev';" "$out"; then
  sed -i '' "s/const SETUP_VERSION = 'dev';/const SETUP_VERSION = '$version';/" "$out"
  grep -c "const SETUP_VERSION = '$version';" "$out" > /dev/null
fi
echo "built $out ($version, $(wc -l < "$out" | tr -d ' ') lines)"
