#!/bin/bash
# Build script: assembles the single-file offline HTML
# Concatenates library scripts inline into the final index.html

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
LIBS_DIR="$SCRIPT_DIR/libs"
SRC_DIR="$SCRIPT_DIR/src"
OUTPUT="$SCRIPT_DIR/index.html"

echo "🔨 Building offline single-file app..."

# Check libs exist
if [ ! -f "$LIBS_DIR/xlsx.full.min.js" ] || [ ! -f "$LIBS_DIR/jspdf.umd.min.js" ] || [ ! -f "$LIBS_DIR/jspdf-autotable.min.js" ] || [ ! -f "$LIBS_DIR/Roboto-Regular.ttf" ] || [ ! -f "$LIBS_DIR/Roboto-Bold.ttf" ]; then
  echo "❌ Library files missing in libs/ directory. Downloading..."
  mkdir -p "$LIBS_DIR"
  curl -sL "https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js" -o "$LIBS_DIR/xlsx.full.min.js"
  curl -sL "https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js" -o "$LIBS_DIR/jspdf.umd.min.js"
  curl -sL "https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.8.4/jspdf.plugin.autotable.min.js" -o "$LIBS_DIR/jspdf-autotable.min.js"
  # Download Fonts
  curl -sL "https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.66/fonts/Roboto/Roboto-Regular.ttf" -o "$LIBS_DIR/Roboto-Regular.ttf"
  curl -sL "https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.66/fonts/Roboto/Roboto-Medium.ttf" -o "$LIBS_DIR/Roboto-Bold.ttf"
  echo "✅ Libraries downloaded."
fi

# Read src files
APP_CSS=$(cat "$SRC_DIR/style.css")
APP_JS=$(cat "$SRC_DIR/app.js")

# Read library files
XLSX_JS=$(cat "$LIBS_DIR/xlsx.full.min.js" | sed '/sourceMappingURL/d')
JSPDF_JS=$(cat "$LIBS_DIR/jspdf.umd.min.js" | sed '/sourceMappingURL/d')
AUTOTABLE_JS=$(cat "$LIBS_DIR/jspdf-autotable.min.js" | sed '/sourceMappingURL/d')
QRCODE_JS=$(cat "$LIBS_DIR/qrcode.min.js")

# Read HTML template
HTML_TEMPLATE=$(cat "$SRC_DIR/template.html")

# Encode logo as base64
LOGO_B64=""
if [ -f "$SRC_DIR/logo.png" ]; then
  LOGO_B64="data:image/png;base64,$(base64 -i "$SRC_DIR/logo.png")"
  echo "📸 Logo embedded as base64"
fi

# Encode sample Excel files as base64
SAMPLE_STUDENTS_B64=""
if [ -f "$SCRIPT_DIR/public/sample_students.xlsx" ]; then
  SAMPLE_STUDENTS_B64="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,$(base64 -i "$SCRIPT_DIR/public/sample_students.xlsx")"
  echo "📄 Sample students Excel embedded"
fi

SAMPLE_ROOMS_B64=""
if [ -f "$SCRIPT_DIR/public/sample_rooms.xlsx" ]; then
  SAMPLE_ROOMS_B64="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,$(base64 -i "$SCRIPT_DIR/public/sample_rooms.xlsx")"
  echo "📄 Sample rooms Excel embedded"
fi

# Encode Fonts
FONT_REGULAR_B64=""
if [ -f "$LIBS_DIR/Roboto-Regular.ttf" ]; then
  FONT_REGULAR_B64="$(base64 -i "$LIBS_DIR/Roboto-Regular.ttf")"
  echo "🔤 Roboto-Regular font embedded"
fi

FONT_BOLD_B64=""
if [ -f "$LIBS_DIR/Roboto-Bold.ttf" ]; then
  FONT_BOLD_B64="$(base64 -i "$LIBS_DIR/Roboto-Bold.ttf")"
  echo "🔤 Roboto-Bold font embedded"
fi

# Build final HTML by replacing placeholders
# We use awk for multi-line replacement
{
echo '<!DOCTYPE html>'
echo '<html lang="tr">'
echo '<head>'
echo '  <meta charset="UTF-8" />'
echo '  <meta name="viewport" content="width=device-width, initial-scale=1.0" />'
echo '  <title>Otomatik Sınav Salon Dağıtım Uygulaması</title>'
echo '  <meta name="description" content="Sınav dağıtım işlemlerinizi kolayca yönetin. Öğrenci listesi yükleyin, salonları tanımlayın, otomatik dağıtım yapın." />'
echo '  <style>'
echo "$APP_CSS"
echo '  </style>'
echo '</head>'
echo '<body>'
echo "$HTML_TEMPLATE"
echo '<script>'
echo "$XLSX_JS"
echo '</script>'
echo '<script>'
echo "$JSPDF_JS"
echo '</script>'
echo '<script>'
echo "$AUTOTABLE_JS"
echo '</script>'
echo '<script>'
echo "$QRCODE_JS"
echo '</script>'
echo '<script>'
echo "$APP_JS"
echo '</script>'
if [ -n "$LOGO_B64" ]; then
  echo '<script>'
  echo "document.getElementById('header-logo').src = '$LOGO_B64';"
  echo '</script>'
fi
if [ -n "$SAMPLE_STUDENTS_B64" ]; then
  echo '<script>'
  echo "document.getElementById('download-student-template').href = '$SAMPLE_STUDENTS_B64';"
  echo '</script>'
fi
if [ -n "$SAMPLE_ROOMS_B64" ]; then
  echo '<script>'
  echo "document.getElementById('download-room-template').href = '$SAMPLE_ROOMS_B64';"
  echo '</script>'
fi
if [ -n "$FONT_REGULAR_B64" ]; then
  echo '<script>'
  echo "window.fontRobotoRegular = '$FONT_REGULAR_B64';"
  echo "window.fontRobotoBold = '$FONT_BOLD_B64';"
  echo '</script>'
fi
echo '</body>'
echo '</html>'
} > "$OUTPUT"

FILE_SIZE=$(wc -c < "$OUTPUT" | tr -d ' ')
echo "✅ Build complete: $OUTPUT ($FILE_SIZE bytes)"
echo "📁 File can be opened directly in any browser — works fully offline!"
