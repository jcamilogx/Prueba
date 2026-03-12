# Claude AI para Excel — Complemento Empresarial

Complemento (Add-in) de Excel que integra Claude Sonnet vía **Azure Foundry** con tu API corporativa. Lee todo el libro de Excel y puede modificarlo según tus instrucciones.

## Características

- **Lectura completa del libro**: Lee todas las hojas, datos, fórmulas y estructura
- **Modificación inteligente**: Escribe celdas, crea hojas, aplica formatos, inserta fórmulas
- **Azure Foundry**: Se conecta a tu deployment corporativo de Claude Sonnet
- **Chat interactivo**: Interfaz conversacional integrada en el panel lateral de Excel
- **Acciones rápidas**: Resumir libro, analizar datos, crear dashboards, limpiar datos
- **Sin servidor propio**: Todo corre en el navegador, se despliega con GitHub Pages

## Requisitos

1. **Cuenta de Azure** con acceso a Azure AI Foundry
2. **Deployment de Claude** activo en Azure Foundry (Sonnet 4.5, 4.6, Opus, o Haiku)
3. **API Key** de tu recurso Azure Foundry
4. **Excel** (Desktop Windows/Mac, Excel Online, o Excel en iPad)

## Despliegue con GitHub Pages

### Paso 1: Crear el repositorio

```bash
# Clonar o crear repositorio
git init excel-claude-addin
cd excel-claude-addin

# Copiar todos los archivos de este proyecto
# (manifest.xml, src/, assets/)
```

### Paso 2: Configurar URLs en el manifest

Edita `manifest.xml` y reemplaza **todas** las ocurrencias de:
- `YOUR_GITHUB_USERNAME` → tu usuario de GitHub
- `YOUR_REPO_NAME` → el nombre de tu repositorio

Por ejemplo, si tu usuario es `miempresa` y el repo es `excel-claude-addin`:
```
https://miempresa.github.io/excel-claude-addin/src/taskpane.html
```

### Paso 3: Subir a GitHub y activar Pages

```bash
git add .
git commit -m "Claude AI Excel Add-in"
git branch -M main
git remote add origin https://github.com/TU_USUARIO/TU_REPO.git
git push -u origin main
```

Luego en GitHub:
1. Ve a **Settings** → **Pages**
2. En **Source** selecciona **Deploy from a branch**
3. Selecciona **main** y **/ (root)**
4. Clic en **Save**
5. Espera 1-2 minutos hasta que la URL esté activa

### Paso 4: Instalar el complemento en Excel

#### Opción A: Sideloading (desarrollo/pruebas)

**Excel Desktop (Windows):**
1. Abre Excel
2. Ve a **Insertar** → **Complementos** → **Mis complementos**
3. Clic en **Cargar mi complemento**
4. Selecciona el archivo `manifest.xml` de tu equipo local

**Excel Online:**
1. Abre Excel en el navegador (office.com)
2. Ve a **Insertar** → **Complementos de Office**
3. Clic en **Cargar mi complemento**
4. Sube el `manifest.xml`

**Excel para Mac:**
1. Abre Finder
2. Ve a `/Users/<tu-usuario>/Library/Containers/com.microsoft.Excel/Data/Documents/wef/`
3. Copia el `manifest.xml` ahí
4. Reinicia Excel

#### Opción B: Despliegue centralizado (producción)

Para desplegar a toda la empresa:
1. Ve al [Centro de administración de Microsoft 365](https://admin.microsoft.com)
2. **Configuración** → **Aplicaciones integradas** → **Complementos**
3. Clic en **Implementar complemento** → **Cargar aplicación personalizada**
4. Sube el `manifest.xml`
5. Asigna usuarios o grupos

### Paso 5: Configurar la conexión

1. Abre Excel y haz clic en el botón **"Abrir Claude"** en la pestaña Inicio
2. Haz clic en el ícono de ⚙️ configuración
3. Ingresa:
   - **Nombre del recurso Azure**: (la parte antes de `.services.ai.azure.com`)
   - **API Key**: tu clave de Azure Foundry
   - **Modelo**: selecciona el deployment que tengas activo
4. Clic en **Guardar y conectar**

## Estructura del proyecto

```
excel-claude-addin/
├── manifest.xml          # Manifiesto del add-in (config para Office)
├── src/
│   └── taskpane.html     # Aplicación completa del panel (HTML + CSS + JS)
├── assets/
│   ├── icon-16.png       # Iconos del complemento
│   ├── icon-32.png
│   ├── icon-64.png
│   └── icon-80.png
└── README.md
```

## Operaciones Excel soportadas

Claude puede ejecutar estas operaciones sobre tu libro:

| Operación | Descripción |
|---|---|
| `write_cells` | Escribe un bloque de datos en un rango |
| `write_cell` | Escribe una celda individual |
| `set_formula` | Inserta una fórmula de Excel |
| `create_sheet` | Crea una nueva hoja |
| `delete_rows` | Elimina filas |
| `insert_rows` | Inserta filas |
| `format_range` | Aplica formato (negrita, color, bordes, etc.) |
| `auto_fit` | Autoajusta columnas y filas |
| `sort_range` | Ordena un rango |
| `clear_range` | Limpia un rango |

## Cómo funciona la lectura del libro

Cuando envías un mensaje, el complemento:

1. **Lee todas las hojas** del libro actual usando la API de Office.js
2. **Extrae los datos visibles** (texto/valores) de cada hoja
3. **Envía el contexto** completo a Claude junto con tu instrucción
4. **Claude analiza** y responde, opcionalmente incluyendo operaciones
5. **El complemento ejecuta** las operaciones sobre el libro

**Límites**: Se leen hasta 500 filas por hoja para mantener el rendimiento. Para hojas más grandes, Claude verá las primeras 500 filas.

## Seguridad

- La API key se almacena solo en `localStorage` del navegador
- Los datos del libro **no se almacenan** — se leen en tiempo real para cada consulta
- La comunicación con Azure Foundry usa HTTPS
- No hay servidor intermedio: la llamada va directo del navegador al endpoint de Azure

## Troubleshooting

**"Error de conexión" al guardar config:**
- Verifica que el nombre del recurso no incluya `.services.ai.azure.com` (solo el nombre)
- Confirma que la API key esté vigente
- Verifica que el modelo coincida con tu deployment en Azure Foundry

**El complemento no aparece en Excel:**
- Verifica que GitHub Pages esté activo y la URL sea accesible
- Comprueba que las URLs en `manifest.xml` sean correctas
- En Excel Online, prueba limpiar caché y recargar

**"Solo funciona en Excel":**
- El complemento detecta que se cargó fuera de Excel. Asegúrate de abrirlo desde Excel.

## Personalización

Puedes modificar `src/taskpane.html` para:
- Cambiar colores y tema (variables CSS al inicio)
- Ajustar las acciones rápidas predefinidas
- Modificar el prompt de sistema por defecto
- Agregar nuevas operaciones de Excel
- Cambiar el límite de filas leídas por hoja

## Licencia

Uso interno empresarial. Modifica según las políticas de tu organización.
