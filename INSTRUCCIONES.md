# AutoFácil Dashboard — Guía de conexión en vivo

## Arquitectura

```
Excel en OneDrive (se actualiza)
        ↓  (cada día a las 6am o manual)
Vercel API route /api/actualizar
  → Descarga el Excel
  → Procesa datos (Python)
  → Sube dashboard_data.json a GitHub
        ↓
Vercel redeploya automáticamente
        ↓
Dashboard HTML lee /data/dashboard_data.json
```

---

## Paso 1 — Estructura del repositorio GitHub

Tu repo debe tener esta estructura:

```
tu-repo/
├── index.html              ← el dashboard (ya lo tienes)
├── data/
│   └── dashboard_data.json ← se genera automáticamente
├── api/
│   └── actualizar.js       ← backend Vercel
├── vercel.json             ← configura el cron
└── README.md
```

---

## Paso 2 — Crear token de GitHub

1. Ve a https://github.com/settings/tokens/new
2. Nombre: `autofacil-dashboard`
3. Expiración: 1 año (o sin expiración)
4. Permisos: marca **`repo`** (acceso completo al repositorio)
5. Clic **"Generate token"**
6. **Copia el token** — solo se muestra una vez (empieza con `ghp_...`)

---

## Paso 3 — Variables de entorno en Vercel

1. Ve a tu proyecto en https://vercel.com
2. **Settings** → **Environment Variables**
3. Agrega estas 4 variables:

| Variable | Valor |
|---|---|
| `ONEDRIVE_URL` | `https://1drv.ms/x/c/2c6ba7ee2485da71/IQC...` (tu URL del Excel) |
| `GITHUB_TOKEN` | `ghp_xxxxxxxxxxxx` (el token que creaste) |
| `GITHUB_REPO` | `tu-usuario/tu-repositorio` |
| `CRON_SECRET` | cualquier texto secreto, ej: `mi-clave-secreta-123` |

---

## Paso 4 — Modificar el dashboard HTML para leer el JSON

Agrega esto al inicio del script del dashboard, reemplazando la carga de datos actual:

```javascript
// Cargar datos desde el JSON generado automáticamente
async function cargarDatos() {
  const r = await fetch('/data/dashboard_data.json');
  const datos = await r.json();
  
  // Mostrar fecha de última actualización
  document.getElementById('ultima-actualizacion').textContent = 
    '🔄 Actualizado: ' + datos.generado_en;
  
  // Pasar datos al dashboard
  window.RAW_DATA    = datos.raw;
  window.DASH_INLINE = calcularResumenInicial(datos.raw);
  window.DASH        = window.DASH_INLINE;
  
  // Reemplazar EJ_PERF con los datos frescos
  Object.assign(EJ_PERF, datos.ej_perf);
  
  buildV1();
}

cargarDatos();
```

---

## Paso 5 — Configurar el cron (actualización automática)

El archivo `vercel.json` ya está configurado para correr **cada día a las 6am UTC** (lunes a viernes):

```json
{
  "crons": [{
    "path": "/api/actualizar",
    "schedule": "0 6 * * 1-5"
  }]
}
```

Para cambiar la hora, modifica el cron. Ejemplos:
- `"0 9 * * 1-5"` → 9am UTC (6am Chile)
- `"0 */4 * * *"` → cada 4 horas
- `"0 * * * *"`   → cada hora

---

## Paso 6 — Actualización manual

Para forzar una actualización manualmente, puedes:

**Opción A — Desde tu PC:**
```bash
# Configura las variables de entorno primero
export GITHUB_TOKEN="ghp_xxxx"
export GITHUB_REPO="usuario/repo"
export ONEDRIVE_URL="https://1drv.ms/..."

python3 actualizar_datos.py
```

**Opción B — Desde el navegador:**
```
https://tu-app.vercel.app/api/actualizar
Authorization: Bearer mi-clave-secreta-123
```

**Opción C — Botón en el dashboard:**
Agrega un botón "Actualizar datos" que llame al endpoint con el CRON_SECRET.

---

## Verificar que funciona

1. Sube todos los archivos al repo de GitHub
2. Vercel detecta el push y redeploya (~30 segundos)
3. Ve a `https://tu-app.vercel.app/api/actualizar` con el header Authorization
4. Si ves `{"ok":true,"registros":8856}` → ¡todo funciona!
5. El dashboard se actualizará con los datos frescos

---

## Troubleshooting

**Error "No se pudo obtener el archivo Excel"**
→ El link de OneDrive caducó o cambió. Genera un nuevo link compartido en OneDrive.

**Error 401 en GitHub**
→ El token venció o tiene permisos insuficientes. Genera uno nuevo con permisos `repo`.

**El cron no corre**
→ Vercel crons requieren plan Pro. En plan gratuito, usa la actualización manual desde tu PC.
