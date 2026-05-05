// api/usuarios.js
// Vercel Serverless Function — gestión de usuarios AutoFácil
// Acciones: login, cambiar-clave, listar (admin), guardar (admin)

const https = require('https');

const GITHUB_TOKEN = process.env.GITHUB_TOKEN;
const GITHUB_REPO  = process.env.GITHUB_REPO;
const USERS_FILE   = 'data/usuarios.json';
const PERMISOS_FILE = 'data/tab_permisos.json';

// ── GitHub helpers ────────────────────────────────────────────────────────────

function githubRequest(method, path, body = null) {
  return new Promise((resolve, reject) => {
    const data = body ? JSON.stringify(body) : null;
    const opts = {
      hostname: 'api.github.com',
      path,
      method,
      headers: {
        'Authorization': `token ${GITHUB_TOKEN}`,
        'Accept': 'application/vnd.github.v3+json',
        'User-Agent': 'AutoFacil-Dashboard/1.0',
        'Content-Type': 'application/json',
        ...(data ? { 'Content-Length': Buffer.byteLength(data) } : {}),
      },
    };
    const req = https.request(opts, res => {
      const chunks = [];
      res.on('data', c => chunks.push(c));
      res.on('end', () => {
        try { resolve({ status: res.statusCode, body: JSON.parse(Buffer.concat(chunks).toString() || '{}') }); }
        catch(e) { resolve({ status: res.statusCode, body: {} }); }
      });
    });
    req.on('error', reject);
    if (data) req.write(data);
    req.end();
  });
}

const USUARIOS_INICIALES = [
  {nombre:"Ademir Norambuena",     usuario:"ademir.norambuena@autofacilchile.cl",     clave:"AF2026",     perfil:"USUARIO",       estado:"NUNCA INGRESADO", ultimoIngreso:null},
  {nombre:"Bernardo Ponce",        usuario:"bernardo.ponce@autofacilchile.cl",        clave:"AF2026",     perfil:"USUARIO",       estado:"NUNCA INGRESADO", ultimoIngreso:null},
  {nombre:"Bryan Saavedra",        usuario:"bryan.saavedra@autofacilchile.cl",        clave:"AF2026",     perfil:"USUARIO",       estado:"NUNCA INGRESADO", ultimoIngreso:null},
  {nombre:"Constanza Piedrahita",  usuario:"constanza.piedrahita@autofacilchile.cl",  clave:"AF2026",     perfil:"USUARIO",       estado:"NUNCA INGRESADO", ultimoIngreso:null},
  {nombre:"Cristina Peña",         usuario:"cristina.pena@autofacilchile.cl",         clave:"AF2026",     perfil:"USUARIO",       estado:"NUNCA INGRESADO", ultimoIngreso:null},
  {nombre:"Dagoberto Irribarra",   usuario:"dagoberto.irribarra@autofacilchile.cl",   clave:"AF2026",     perfil:"USUARIO",       estado:"NUNCA INGRESADO", ultimoIngreso:null},
  {nombre:"Eduardo Abad",          usuario:"eduardo.abad@autofacilchile.cl",          clave:"AF2026",     perfil:"USUARIO",       estado:"NUNCA INGRESADO", ultimoIngreso:null},
  {nombre:"Fabiola Torres",        usuario:"fabiola.torres@autofacilchile.cl",        clave:"AF2026",     perfil:"USUARIO",       estado:"NUNCA INGRESADO", ultimoIngreso:null},
  {nombre:"Jorge Vargas",          usuario:"jorge.vargas@autofacilchile.cl",          clave:"AF2026",     perfil:"USUARIO",       estado:"NUNCA INGRESADO", ultimoIngreso:null},
  {nombre:"Juan Manuel Bustamante",usuario:"juan.bustamante@autofacilchile.cl",       clave:"AF2026",     perfil:"USUARIO",       estado:"NUNCA INGRESADO", ultimoIngreso:null},
  {nombre:"Karina Díaz",           usuario:"karina.diaz@autofacilchile.cl",           clave:"AF2026",     perfil:"USUARIO",       estado:"NUNCA INGRESADO", ultimoIngreso:null},
  {nombre:"Katherine Trillo",      usuario:"katherine.trillo@autofacilchile.cl",      clave:"AF2026",     perfil:"USUARIO",       estado:"NUNCA INGRESADO", ultimoIngreso:null},
  {nombre:"Leonardo Sevilla",      usuario:"leonardo.sevilla@autofacilchile.cl",      clave:"AF2026",     perfil:"USUARIO",       estado:"NUNCA INGRESADO", ultimoIngreso:null},
  {nombre:"Noelia González",       usuario:"noelia.gonzalez@autofacilchile.cl",       clave:"AF2026",     perfil:"USUARIO",       estado:"NUNCA INGRESADO", ultimoIngreso:null},
  {nombre:"Patricio Escobar",      usuario:"patricio.escobar@autofacilchile.cl",      clave:"Admin2026!", perfil:"ADMINISTRADOR", estado:"ACTIVO",          ultimoIngreso:null},
  {nombre:"Sandra Ayala",          usuario:"sandra.ayala@autofacilchile.cl",          clave:"AF2026",     perfil:"USUARIO",       estado:"NUNCA INGRESADO", ultimoIngreso:null},
];

async function leerUsuarios() {
  const apiPath = `/repos/${GITHUB_REPO}/contents/${USERS_FILE}`;
  const res = await githubRequest('GET', apiPath);
  if (res.status === 404 || !res.body.content) {
    return { usuarios: USUARIOS_INICIALES.map(u => ({...u})), sha: null };
  }
  try {
    const content = Buffer.from(res.body.content.replace(/\n/g, ''), 'base64').toString('utf-8');
    const usuarios = JSON.parse(content);
    // Migración: asegurar que todos tengan campo primerIngreso
    usuarios.forEach(u => {
      if (u.primerIngreso === undefined) {
        u.primerIngreso = u.estado === 'NUNCA INGRESADO';
      }
    });
    return { usuarios, sha: res.body.sha };
  } catch(e) {
    console.error('Error parseando usuarios.json:', e.message);
    return { usuarios: USUARIOS_INICIALES.map(u => ({...u})), sha: null };
  }
}

async function leerPermisos() {
  const apiPath = `/repos/${GITHUB_REPO}/contents/${PERMISOS_FILE}`;
  const res = await githubRequest('GET', apiPath);
  if (res.status === 404 || !res.body.content) return { permisos: {}, sha: null };
  try {
    const content = Buffer.from(res.body.content.replace(/\n/g, ''), 'base64').toString('utf-8');
    return { permisos: JSON.parse(content), sha: res.body.sha };
  } catch(e) { return { permisos: {}, sha: null }; }
}

async function guardarPermisos(permisos, sha) {
  const apiPath = `/repos/${GITHUB_REPO}/contents/${PERMISOS_FILE}`;
  if (!sha) {
    const check = await githubRequest('GET', apiPath);
    if (check.body.sha) sha = check.body.sha;
  }
  const contenidoB64 = Buffer.from(JSON.stringify(permisos, null, 2)).toString('base64');
  const result = await githubRequest('PUT', apiPath, {
    message: 'admin: actualizar permisos de pestañas',
    content: contenidoB64,
    branch: 'main',
    ...(sha ? { sha } : {}),
  });
  if (result.status !== 200 && result.status !== 201) {
    throw new Error(`GitHub PUT permisos falló: ${result.status}`);
  }
}

async function guardarUsuarios(usuarios, sha, mensaje = 'usuarios: actualización') {
  const apiPath = `/repos/${GITHUB_REPO}/contents/${USERS_FILE}`;
  // Si no tenemos sha, obtenerlo antes de guardar para evitar conflicto
  if (!sha) {
    const check = await githubRequest('GET', apiPath);
    if (check.body.sha) sha = check.body.sha;
  }
  const contenidoB64 = Buffer.from(JSON.stringify(usuarios, null, 2)).toString('base64');
  const result = await githubRequest('PUT', apiPath, {
    message: mensaje,
    content: contenidoB64,
    branch: 'main',
    ...(sha ? { sha } : {}),
  });
  if (result.status !== 200 && result.status !== 201) {
    throw new Error(`GitHub PUT falló: ${result.status} — ${JSON.stringify(result.body)}`);
  }
}

// ── CORS headers ──────────────────────────────────────────────────────────────

function setCors(res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');
}

// ── Handler ───────────────────────────────────────────────────────────────────

module.exports = async function handler(req, res) {
  setCors(res);
  if (req.method === 'OPTIONS') return res.status(200).end();

  try {
    const { accion, usuario, clave, nueva_clave, perfil, estado, nombre } = req.body || {};

    // ── LOGIN ─────────────────────────────────────────────────────────────────
    if (accion === 'login') {
      if (!usuario || !clave) return res.status(400).json({ ok: false, error: 'Faltan datos' });

      const { usuarios, sha } = await leerUsuarios();
      const idx = usuarios.findIndex(u => u.usuario.toLowerCase() === usuario.toLowerCase());
      if (idx === -1) return res.status(401).json({ ok: false, error: 'Usuario no encontrado' });

      const u = usuarios[idx];
      if (u.estado === 'BLOQUEADO')   return res.status(403).json({ ok: false, error: 'Usuario bloqueado' });
      if (u.estado === 'SUSPENDIDO')  return res.status(403).json({ ok: false, error: 'Usuario suspendido' });
      if (u.clave !== clave)          return res.status(401).json({ ok: false, error: 'Contraseña incorrecta' });

      const primerIngreso = u.primerIngreso !== false && u.estado === 'NUNCA INGRESADO';
      usuarios[idx].ultimoIngreso = new Date().toISOString();
      usuarios[idx].estado = 'ACTIVO';
      usuarios[idx].primerIngreso = false; // marcar que ya no es primer ingreso

      await guardarUsuarios(usuarios, sha, `login: ${usuario}`);

      const { permisos } = await leerPermisos();
      return res.status(200).json({
        ok: true,
        primerIngreso,
        sesion: { nombre: u.nombre, usuario: u.usuario, perfil: u.perfil },
        permisos,
      });
    }

    // ── CAMBIAR CLAVE ─────────────────────────────────────────────────────────
    if (accion === 'cambiar-clave') {
      if (!usuario || !nueva_clave) return res.status(400).json({ ok: false, error: 'Faltan datos' });

      const { usuarios, sha } = await leerUsuarios();
      const idx = usuarios.findIndex(u => u.usuario.toLowerCase() === usuario.toLowerCase());
      if (idx === -1) return res.status(404).json({ ok: false, error: 'Usuario no encontrado' });

      usuarios[idx].clave = nueva_clave;
      usuarios[idx].estado = 'ACTIVO';
      usuarios[idx].ultimoIngreso = new Date().toISOString();

      await guardarUsuarios(usuarios, sha, `clave: ${usuario}`);
      return res.status(200).json({ ok: true });
    }

    // ── LISTAR (admin) ────────────────────────────────────────────────────────
    if (accion === 'listar') {
      const authHeader = req.headers.authorization || '';
      if (!authHeader.startsWith('Bearer ')) return res.status(401).json({ ok: false, error: 'No autorizado' });

      const { usuarios } = await leerUsuarios();
      // No devolver claves al frontend
      const seguros = usuarios.map(u => ({
        nombre: u.nombre, usuario: u.usuario, perfil: u.perfil,
        estado: u.estado, ultimoIngreso: u.ultimoIngreso || null,
      }));
      return res.status(200).json({ ok: true, usuarios: seguros });
    }

    // ── GUARDAR usuario (admin) ───────────────────────────────────────────────
    if (accion === 'guardar') {
      const authHeader = req.headers.authorization || '';
      if (!authHeader.startsWith('Bearer ')) return res.status(401).json({ ok: false, error: 'No autorizado' });

      const { usuarios, sha } = await leerUsuarios();
      const idx = usuarios.findIndex(u => u.usuario.toLowerCase() === usuario.toLowerCase());
      const { primerIngreso } = req.body || {};

      if (idx === -1) {
        usuarios.push({ nombre, usuario, clave: clave || 'AF2026', perfil: perfil || 'USUARIO', estado: estado || 'NUNCA INGRESADO', primerIngreso: true, ultimoIngreso: null });
      } else {
        if (nombre  !== undefined) usuarios[idx].nombre  = nombre;
        if (clave   !== undefined && clave)  usuarios[idx].clave   = clave;
        if (perfil  !== undefined) usuarios[idx].perfil  = perfil;
        if (estado  !== undefined) usuarios[idx].estado  = estado;
        if (primerIngreso !== undefined) usuarios[idx].primerIngreso = primerIngreso;
      }

      await guardarUsuarios(usuarios, sha, `admin: guardar ${usuario}`);
      return res.status(200).json({ ok: true });
    }

    // ── ELIMINAR (admin) ──────────────────────────────────────────────────────
    if (accion === 'eliminar') {
      const authHeader = req.headers.authorization || '';
      if (!authHeader.startsWith('Bearer ')) return res.status(401).json({ ok: false, error: 'No autorizado' });

      const { usuarios, sha } = await leerUsuarios();
      const admins = usuarios.filter(u => u.perfil === 'ADMINISTRADOR');
      const target = usuarios.find(u => u.usuario.toLowerCase() === usuario.toLowerCase());
      if (target?.perfil === 'ADMINISTRADOR' && admins.length <= 1) {
        return res.status(400).json({ ok: false, error: 'No puedes eliminar el único administrador' });
      }
      const nuevos = usuarios.filter(u => u.usuario.toLowerCase() !== usuario.toLowerCase());
      await guardarUsuarios(nuevos, sha, `admin: eliminar ${usuario}`);
      return res.status(200).json({ ok: true });
    }

    // ── OBTENER PERMISOS (público) ────────────────────────────────────────────
    if (accion === 'obtener-permisos') {
      const { permisos } = await leerPermisos();
      return res.status(200).json({ ok: true, permisos });
    }

    // ── GUARDAR PERMISOS (admin) ──────────────────────────────────────────────
    if (accion === 'guardar-permisos') {
      const authHeader = req.headers.authorization || '';
      if (!authHeader.startsWith('Bearer ')) return res.status(401).json({ ok: false, error: 'No autorizado' });
      const { permisos } = req.body || {};
      if (!permisos || typeof permisos !== 'object') return res.status(400).json({ ok: false, error: 'Permisos inválidos' });
      const { sha } = await leerPermisos();
      await guardarPermisos(permisos, sha);
      return res.status(200).json({ ok: true });
    }

    return res.status(400).json({ ok: false, error: 'Acción no reconocida' });

  } catch (err) {
    console.error('Error usuarios.js:', err.message);
    return res.status(500).json({ ok: false, error: err.message });
  }
};
