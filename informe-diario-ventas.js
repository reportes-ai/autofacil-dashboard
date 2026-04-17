const nodemailer = require('nodemailer');

const SUPABASE_URL = process.env.SUPABASE_URL;
const SUPABASE_KEY = process.env.SUPABASE_KEY;

async function sbGet(table, params = '') {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/${table}?${params}`, {
    headers: { 'apikey': SUPABASE_KEY, 'Authorization': `Bearer ${SUPABASE_KEY}` }
  });
  if (!res.ok) throw new Error(await res.text());
  return res.json();
}

function fmtMonto(v) {
  if (!v && v !== 0) return '$ 0';
  return '$ ' + Number(v).toLocaleString('es-CL', { maximumFractionDigits: 0 });
}

async function main() {
  // Fecha de corte: ayer (si hoy es lunes, tomar viernes)
  const hoy = new Date();
  hoy.setHours(0, 0, 0, 0);
  const ayer = new Date(hoy);
  ayer.setDate(ayer.getDate() - 1);
  if (hoy.getDay() === 1) ayer.setDate(ayer.getDate() - 2);

  // Desde el día 1 del mes de ayer
  const primerDiaMes = new Date(ayer.getFullYear(), ayer.getMonth(), 1);
  const desdeStr = primerDiaMes.toISOString().split('T')[0];
  const hastaStr = hoy.toISOString().split('T')[0]; // < hoy = hasta ayer inclusive

  console.log(`Período: ${desdeStr} → ${ayer.toISOString().split('T')[0]}`);

  // Todos los créditos del mes (cualquier estado) → determina qué ejecutivos participan
  const todosIngresados = await sbGet(
    'creditos',
    `fecha_otorgamiento=gte.${desdeStr}&fecha_otorgamiento=lt.${hastaStr}&order=ejecutivo.asc&limit=5000`
  );

  // Solo otorgados
  const otorgados = todosIngresados.filter(c =>
    (c.estado_eval || '').toUpperCase() === 'OTORGADO'
  );

  // Set de todos los ejecutivos con al menos un ingreso en el mes
  const todosEjecutivos = new Set(
    todosIngresados.map(c => (c.ejecutivo || 'SIN EJECUTIVO').trim())
  );

  // Agrupar otorgados por ejecutivo → financiera
  const porEjecutivo = {};
  otorgados.forEach(c => {
    const ej  = (c.ejecutivo  || 'SIN EJECUTIVO').trim();
    const fin = (c.financiera || 'SIN FINANCIERA').trim();
    if (!porEjecutivo[ej]) porEjecutivo[ej] = {};
    if (!porEjecutivo[ej][fin]) porEjecutivo[ej][fin] = { ops: 0, monto: 0 };
    porEjecutivo[ej][fin].ops++;
    porEjecutivo[ej][fin].monto += Number(c.monto_financiado) || 0;
  });

  // Incluir ejecutivos con ingresos pero sin otorgados (mapa vacío)
  todosEjecutivos.forEach(ej => {
    if (!porEjecutivo[ej]) porEjecutivo[ej] = {};
  });

  const ejecutivos = Object.keys(porEjecutivo).sort();

  // Labels de fecha
  const mesLabel      = ayer.toLocaleDateString('es-CL', { month: 'long', year: 'numeric' });
  const ayerLabel     = ayer.toLocaleDateString('es-CL', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' });
  const primerLabel   = primerDiaMes.toLocaleDateString('es-CL', { day: 'numeric', month: 'long' });
  const ayerCorto     = ayer.toLocaleDateString('es-CL', { day: 'numeric', month: 'long' });

  let totalOpsGlobal   = 0;
  let totalMontoGlobal = 0;
  let bloques = '';

  ejecutivos.forEach(ej => {
    const fins = Object.keys(porEjecutivo[ej]).sort();
    let totalOpsEj   = 0;
    let totalMontoEj = 0;

    let filas = '';
    if (fins.length === 0) {
      filas = `
        <tr>
          <td style="padding:7px 12px;border-bottom:1px solid #ECEFF1;font-size:13px;color:#90A4AE">—</td>
          <td style="padding:7px 12px;border-bottom:1px solid #ECEFF1;font-size:13px;text-align:center;color:#90A4AE">0</td>
          <td style="padding:7px 12px;border-bottom:1px solid #ECEFF1;font-size:13px;text-align:right;color:#90A4AE">$ 0</td>
        </tr>`;
    } else {
      filas = fins.map(fin => {
        const { ops, monto } = porEjecutivo[ej][fin];
        totalOpsEj   += ops;
        totalMontoEj += monto;
        return `
          <tr>
            <td style="padding:7px 12px;border-bottom:1px solid #ECEFF1;font-size:13px">${fin}</td>
            <td style="padding:7px 12px;border-bottom:1px solid #ECEFF1;font-size:13px;text-align:center;font-weight:600;color:#1565C0">${ops}</td>
            <td style="padding:7px 12px;border-bottom:1px solid #ECEFF1;font-size:13px;text-align:right">${fmtMonto(monto)}</td>
          </tr>`;
      }).join('');
    }

    totalOpsGlobal   += totalOpsEj;
    totalMontoGlobal += totalMontoEj;

    const nombre     = ej.split(' ').map(w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase()).join(' ');
    const sinVentas  = totalOpsEj === 0;
    const colorAcc   = sinVentas ? '#CFD8DC' : '#2196F3';
    const colorHead  = sinVentas ? '#90A4AE' : '#1565C0';
    const colorFoot  = sinVentas ? '#F5F5F5' : '#E3F2FD';
    const colorTxt   = sinVentas ? '#90A4AE' : '#1565C0';

    bloques += `
      <div style="margin-bottom:24px">
        <h3 style="margin:0 0 8px;font-size:14px;color:#1a2a4a;border-left:4px solid ${colorAcc};padding-left:10px">
          ${nombre}${sinVentas ? ' <span style="font-size:11px;color:#90A4AE;font-weight:400">(sin otorgados)</span>' : ''}
        </h3>
        <table style="width:100%;border-collapse:collapse;background:#fff;border-radius:6px;overflow:hidden;border:1px solid #CFD8DC">
          <thead>
            <tr style="background:${colorHead};color:#fff">
              <th style="padding:8px 12px;text-align:left;font-size:12px;font-weight:600">Financiera</th>
              <th style="padding:8px 12px;text-align:center;font-size:12px;font-weight:600">Otorgados</th>
              <th style="padding:8px 12px;text-align:right;font-size:12px;font-weight:600">Monto Financiado</th>
            </tr>
          </thead>
          <tbody>${filas}</tbody>
          <tfoot>
            <tr style="background:${colorFoot}">
              <td style="padding:7px 12px;font-size:13px;font-weight:700;color:${colorTxt}">Total Ejecutivo</td>
              <td style="padding:7px 12px;font-size:13px;font-weight:700;color:${colorTxt};text-align:center">${totalOpsEj}</td>
              <td style="padding:7px 12px;font-size:13px;font-weight:700;color:${colorTxt};text-align:right">${fmtMonto(totalMontoEj)}</td>
            </tr>
          </tfoot>
        </table>
      </div>`;
  });

  const resumen = `
    <div style="margin-top:8px;border-top:2px solid #1565C0;padding-top:16px">
      <table style="width:100%;border-collapse:collapse;background:#1565C0;border-radius:6px;overflow:hidden">
        <tbody>
          <tr>
            <td style="padding:12px 14px;font-size:13px;color:#fff;font-weight:700">TOTAL GENERAL</td>
            <td style="padding:12px 14px;font-size:20px;color:#fff;font-weight:700;text-align:center">${totalOpsGlobal}</td>
            <td style="padding:12px 14px;font-size:13px;color:#fff;font-weight:700;text-align:right">${fmtMonto(totalMontoGlobal)}</td>
          </tr>
        </tbody>
      </table>
    </div>`;

  if (todosIngresados.length === 0) {
    console.log('No hay créditos en el período. No se envía email.');
    return;
  }

  const html = `
    <div style="font-family:'Segoe UI',Arial,sans-serif;max-width:700px;margin:0 auto;color:#333">
      <div style="background:#1a2a4a;color:#fff;padding:20px 24px;border-radius:8px 8px 0 0">
        <div style="font-size:20px;font-weight:700;margin:0">Auto<span style="color:#2196F3">Fácil</span></div>
        <div style="font-size:16px;font-weight:600;margin:6px 0 2px">Informe Diario de Ventas</div>
        <div style="font-size:12px;opacity:0.8">${ayerLabel}</div>
        <div style="font-size:11px;opacity:0.6;margin-top:3px">Acumulado ${primerLabel} al ${ayerCorto} · ${mesLabel}</div>
      </div>
      <div style="background:#fff;padding:20px 24px;border:1px solid #CFD8DC;border-top:none;border-radius:0 0 8px 8px">
        <p style="font-size:13px;color:#546E7A;margin:0 0 20px;line-height:1.6">
          A continuación encontrarán las ventas por Ejecutivo Comercial.<br>
          <strong>Cualquier diferencia levantarla al área de operaciones.</strong>
        </p>
        ${bloques}
        ${resumen}
        <p style="color:#90A4AE;font-size:11px;margin-top:24px;border-top:1px solid #ECEFF1;padding-top:12px">
          Reporte automático generado por el sistema AutoFácil · Crédito Automotriz
        </p>
      </div>
    </div>`;

  const transporter = nodemailer.createTransport({
    service: 'gmail',
    auth: {
      user: process.env.GMAIL_USER,
      pass: process.env.GMAIL_PASSWORD.replace(/\s/g, ''),
    }
  });

  const ayerCL = ayer.toLocaleDateString('es-CL');
  await transporter.sendMail({
    from: `"AutoFácil Reportes" <${process.env.GMAIL_USER}>`,
    to: process.env.EMAIL_TO,
    subject: `Informe Diario de Ventas — ${ayerCL} · ${totalOpsGlobal} créditos acumulados`,
    html,
  });

  console.log(`Email enviado: ${totalOpsGlobal} otorgados, ${ejecutivos.length} ejecutivos.`);
}

main().catch(err => {
  console.error('Error:', err);
  process.exit(1);
});
