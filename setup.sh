#!/bin/bash
# ═══════════════════════════════════════════════════════
# Script de configuración del complemento Claude AI Excel
# Ejecuta: bash setup.sh tu-usuario tu-repo
# ═══════════════════════════════════════════════════════

if [ $# -lt 2 ]; then
  echo ""
  echo "  Uso: bash setup.sh <github-username> <repo-name>"
  echo ""
  echo "  Ejemplo: bash setup.sh miempresa excel-claude-addin"
  echo ""
  exit 1
fi

USERNAME=$1
REPO=$2

echo "Configurando complemento para: https://${USERNAME}.github.io/${REPO}/"
echo ""

# Reemplazar en manifest.xml
sed -i "s/YOUR_GITHUB_USERNAME/${USERNAME}/g" manifest.xml
sed -i "s/YOUR_REPO_NAME/${REPO}/g" manifest.xml

echo "✅ manifest.xml actualizado"
echo ""
echo "Próximos pasos:"
echo "  1. git add . && git commit -m 'Configurar URLs'"
echo "  2. git push origin main"
echo "  3. Activar GitHub Pages en Settings → Pages → main branch"
echo "  4. Esperar 1-2 min y verificar: https://${USERNAME}.github.io/${REPO}/src/taskpane.html"
echo "  5. Cargar manifest.xml en Excel (Insertar → Complementos → Cargar mi complemento)"
echo ""
