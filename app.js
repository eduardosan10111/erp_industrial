const express = require('express');
const sqlite3 = require('sqlite3').verbose();
const multer = require('multer');
const XLSX = require('xlsx');
const session = require('express-session');
const path = require('path');
const fs = require('fs');
const app = express();
const port = 3000;

app.use(express.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, 'public')));

// Creación de carpetas necesarias
if (!fs.existsSync(path.join(__dirname, 'uploads/ocs'))) fs.mkdirSync(path.join(__dirname, 'uploads/ocs'), { recursive: true });
if (!fs.existsSync(path.join(__dirname, 'uploads/pagos'))) fs.mkdirSync(path.join(__dirname, 'uploads/pagos'), { recursive: true });
if (!fs.existsSync(path.join(__dirname, 'uploads/gastos'))) fs.mkdirSync(path.join(__dirname, 'uploads/gastos'), { recursive: true });
if (!fs.existsSync(path.join(__dirname, 'uploads/csf'))) fs.mkdirSync(path.join(__dirname, 'uploads/csf'), { recursive: true });
if (!fs.existsSync(path.join(__dirname, 'uploads/facturas'))) fs.mkdirSync(path.join(__dirname, 'uploads/facturas'), { recursive: true });
if (!fs.existsSync(path.join(__dirname, 'uploads/almacen'))) fs.mkdirSync(path.join(__dirname, 'uploads/almacen'), { recursive: true });
if (!fs.existsSync(path.join(__dirname, 'uploads/personal'))) fs.mkdirSync(path.join(__dirname, 'uploads/personal'), { recursive: true });
if (!fs.existsSync(path.join(__dirname, 'uploads/productos'))) fs.mkdirSync(path.join(__dirname, 'uploads/productos'), { recursive: true });
if (!fs.existsSync(path.join(__dirname, 'uploads/temp'))) fs.mkdirSync(path.join(__dirname, 'uploads/temp'), { recursive: true });

app.use('/ocs', express.static(path.join(__dirname, 'uploads/ocs')));
app.use('/pagos', express.static(path.join(__dirname, 'uploads/pagos')));
app.use('/gastos', express.static(path.join(__dirname, 'uploads/gastos')));
app.use('/csf', express.static(path.join(__dirname, 'uploads/csf')));
app.use('/facturas', express.static(path.join(__dirname, 'uploads/facturas')));
app.use('/almacen', express.static(path.join(__dirname, 'uploads/almacen')));
app.use('/personal', express.static(path.join(__dirname, 'uploads/personal')));
app.use('/productos_img', express.static(path.join(__dirname, 'uploads/productos')));

app.use(session({
    secret: 'cc_instalaciones_erp_tech_ui_secret',
    resave: false,
    saveUninitialized: true
}));

const upload = multer({ dest: path.join(__dirname, 'uploads/temp/') });
const db = new sqlite3.Database(path.join(__dirname, 'micrm.db'));

const formatoDinero = (cantidad) => {
    return Number(cantidad || 0).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
};

const DATOS_MI_EMPRESA = {
    nombre: "C.C. INSTALACIONES INDUSTRIALES S.A. DE C.V.",
    rfc: "CII210924KU1",
    telefono: "722 308 8012",
    correo: "factura.cta@ccinstalacionesindustriales.com.mx",
    direccion: "Calle Faisan N 16 Tultitlan Estado de Mexico",
    banco: "BBVA México",
    titular: "C.C. INSTALACIONES INDUSTRIALES S.A. DE C.V.",
    cuenta: "0118002711",
    clabe: "012180001180027111"
};

const CATEGORIAS_VALIDAS = [
    "Tubería de Aluminio", "Compresores", "Consumibles", "Eléctrico",
    "Obra Civil", "HVAC", "Hidráulico", "Pailería", "Mano de obra especializada", "Equipos de renta"
];

db.serialize(() => {
    db.run(`CREATE TABLE IF NOT EXISTS usuarios (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT UNIQUE NOT NULL,
    password TEXT NOT NULL,
    departamento TEXT NOT NULL,
    rol TEXT NOT NULL
  )`, () => {
        db.get("SELECT COUNT(*) as total FROM usuarios", [], (err, row) => {
            if (row && row.total === 0) {
                db.run("INSERT INTO usuarios (username, password, departamento, rol) VALUES (?, ?, ?, ?)", ["AdminCCI", "Oficina-cci2026", "Administración", "Administrador"]);
                db.run("INSERT INTO usuarios (username, password, departamento, rol) VALUES (?, ?, ?, ?)", ["ControlAdm", "control2026", "Control Administrativo", "Administrador"]);
                db.run("INSERT INTO usuarios (username, password, departamento, rol) VALUES (?, ?, ?, ?)", ["Ventas1", "ventas2026", "Ventas", "Usuario"]);
                db.run("INSERT INTO usuarios (username, password, departamento, rol) VALUES (?, ?, ?, ?)", ["Contabilidad1", "conta2026", "Contabilidad", "Contabilidad"]);
                db.run("INSERT INTO usuarios (username, password, departamento, rol) VALUES (?, ?, ?, ?)", ["Almacen1", "almacen2026", "Almacén", "Usuario"]);
            }
        });
    });

    db.run(`CREATE TABLE IF NOT EXISTS personal (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    numero_trabajador TEXT UNIQUE NOT NULL,
    nombre_completo TEXT NOT NULL,
    puesto TEXT NOT NULL,
    telefono TEXT,
    rfc TEXT,
    nss TEXT,
    direccion TEXT,
    correo TEXT,
    contacto_emergencia TEXT,
    archivo_ine TEXT,
    archivo_curp TEXT,
    archivo_banco TEXT,
    archivo_cv TEXT,
    fecha_registro DATETIME DEFAULT CURRENT_TIMESTAMP
  )`);

    db.run(`CREATE TABLE IF NOT EXISTS clientes (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    numero_cliente TEXT,
    nombre_empresa TEXT NOT NULL,
    rfc TEXT,
    contacto_nombre TEXT,
    telefono TEXT,
    email TEXT,
    ubicacion TEXT,
    vendedor_asignado TEXT,
    archivo_csf TEXT,
    tiene_credito TEXT DEFAULT 'No',
    monto_credito REAL DEFAULT 0,
    regimen_fiscal TEXT,
    uso_cfdi TEXT,
    dias_credito INTEGER DEFAULT 30,
    fecha_registro DATETIME DEFAULT CURRENT_TIMESTAMP
  )`);

    db.run(`CREATE TABLE IF NOT EXISTS crm_agenda (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    cliente_id INTEGER,
    vendedor TEXT,
    tipo TEXT,
    fecha_programada TEXT,
    comentario TEXT,
    resultado_visita TEXT,
    estado TEXT DEFAULT 'Pendiente',
    FOREIGN KEY(cliente_id) REFERENCES clientes(id)
  )`);

    db.run(`CREATE TABLE IF NOT EXISTS calendario_notas (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    vendedor TEXT,
    mes TEXT NOT NULL,
    anio TEXT NOT NULL,
    nota TEXT NOT NULL,
    fecha_registro DATETIME DEFAULT CURRENT_TIMESTAMP
  )`);

    db.run(`CREATE TABLE IF NOT EXISTS soporte_tickets (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    cliente_id INTEGER,
    vendedor TEXT,
    tipo_solicitud TEXT,
    asunto TEXT NOT NULL,
    descripcion TEXT NOT NULL,
    estado TEXT DEFAULT 'Abierto',
    prioridad TEXT DEFAULT 'Media',
    fecha_registro DATETIME DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY(cliente_id) REFERENCES clientes(id)
  )`);

    db.run(`CREATE TABLE IF NOT EXISTS productos_servicios (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    codigo TEXT,
    categoria TEXT DEFAULT 'Tubería de Aluminio',
    nombre TEXT NOT NULL,
    tipo TEXT NOT NULL,
    precio_unitario REAL NOT NULL,
    stock INTEGER DEFAULT 0,
    tiempo_entrega TEXT DEFAULT 'Inmediato',
    dimensiones TEXT,
    foto TEXT,
    descripcion TEXT,
    fecha_registro DATETIME DEFAULT CURRENT_TIMESTAMP
  )`);

    db.run(`CREATE TABLE IF NOT EXISTS cotizaciones (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    cliente_id INTEGER,
    vendedor TEXT,
    moneda TEXT DEFAULT 'MXN',
    descuento REAL DEFAULT 0,
    vigencia_dias INTEGER DEFAULT 15,
    estado TEXT DEFAULT 'Enviada',
    motivo_estado TEXT,
    archivo_oc TEXT,
    archivo_pago TEXT,
    archivo_factura TEXT,
    porcentaje_pagado REAL DEFAULT 0,
    estatus_almacen TEXT DEFAULT 'Pendiente Logística',
    reporte_entrega TEXT,
    remision_almacen TEXT,
    fecha_registro DATETIME DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY(cliente_id) REFERENCES clientes(id)
  )`);

    db.run(`CREATE TABLE IF NOT EXISTS pagos_cotizacion (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    cotizacion_id INTEGER,
    monto_abono REAL NOT NULL,
    porcentaje_abono REAL NOT NULL,
    referencia TEXT,
    archivo_comprobante TEXT,
    fecha_pago DATETIME DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY(cotizacion_id) REFERENCES cotizaciones(id)
  )`);

    db.run(`CREATE TABLE IF NOT EXISTS cotizacion_detalles (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    cotizacion_id INTEGER,
    producto_id INTEGER,
    cantidad INTEGER,
    precio_unitario REAL,
    FOREIGN KEY(cotizacion_id) REFERENCES cotizaciones(id),
    FOREIGN KEY(producto_id) REFERENCES productos_servicios(id)
  )`);

    db.run(`CREATE TABLE IF NOT EXISTS prefacturas (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    cotizacion_id INTEGER,
    estatus_prefactura TEXT DEFAULT 'Generada',
    monto_total REAL,
    fecha_emision DATETIME DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY(cotizacion_id) REFERENCES cotizaciones(id)
  )`);

    db.run(`CREATE TABLE IF NOT EXISTS chat_interno (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    departamento TEXT,
    remitente TEXT,
    mensaje TEXT,
    fecha DATETIME DEFAULT CURRENT_TIMESTAMP
  )`);

    db.run(`CREATE TABLE IF NOT EXISTS gastos_operativos (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    vendedor TEXT,
    concepto TEXT,
    categoria TEXT,
    monto REAL,
    comprobante TEXT,
    fecha DATETIME DEFAULT CURRENT_TIMESTAMP
  )`);
});

const EstilosGlobales = `
  <link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap">
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
  <style>
    @keyframes pulsoTech {
      0% { background-position: 0 0, 0 0; }
      100% { background-position: 40px 40px, 80px 80px; }
    }
    body {
      background-color: #030712;
      background-image: radial-gradient(circle at 50% 20%, #0f172a 0%, #030712 80%),
        url("data:image/svg+xml,%3Csvg width='40' height='40' viewBox='0 0 40 40' xmlns='http://www.w3.org/2000/svg'%3E%3Cg fill='%2338bdf8' fill-opacity='0.03' fill-rule='evenodd'%3E%3Cpath d='M0 40L40 0H20L0 20M40 40V20L20 40'/%3E%3C/g%3E%3C/svg%3E");
      animation: pulsoTech 30s linear infinite;
      font-family: 'Inter', system-ui, -apple-system, sans-serif;
      margin: 0; padding: 20px; color: #e2e8f0; min-height: 100vh; box-sizing: border-box;
    }
    .contenedor-tarjeta {
      background: rgba(15, 23, 42, 0.75);
      backdrop-filter: blur(16px);
      -webkit-backdrop-filter: blur(16px);
      border: 1px solid rgba(56, 189, 248, 0.15);
      border-radius: 12px;
      box-shadow: 0 10px 30px -10px rgba(0,0,0,0.5), inset 0 1px 0 rgba(255,255,255,0.05);
      padding: 24px;
      margin-bottom: 24px;
      transition: all 0.3s ease;
    }
    .contenedor-tarjeta:hover {
      border-color: rgba(56, 189, 248, 0.3);
      box-shadow: 0 15px 35px -10px rgba(56, 189, 248, 0.1);
    }
    h1, h2, h3, h4 { color: #f8fafc; font-weight: 600; letter-spacing: -0.025em; }
    .btn-primario {
      background: linear-gradient(135deg, #0284c7 0%, #0369a1 100%);
      color: white; border: none; padding: 10px 18px; border-radius: 8px;
      font-weight: 600; cursor: pointer; text-transform: uppercase; font-size: 11px;
      letter-spacing: 0.05em; transition: all 0.2s ease; box-shadow: 0 4px 12px rgba(2, 132, 199, 0.3);
      display: inline-flex; align-items: center; gap: 8px; text-decoration: none;
    }
    .btn-primario:hover { transform: translateY(-1px); filter: brightness(1.1); box-shadow: 0 6px 16px rgba(2, 132, 199, 0.4); }
    .btn-peligro {
      background: linear-gradient(135deg, #ef4444 0%, #b91c1c 100%);
      color: white; border: none; padding: 6px 12px; border-radius: 6px;
      cursor: pointer; font-size: 11px; font-weight: 600; text-decoration: none;
      display: inline-flex; align-items: center; gap: 5px; transition: all 0.2s;
    }
    .btn-peligro:hover { filter: brightness(1.1); }
    .input-industrial {
      width: 100%; padding: 11px 14px; margin-bottom: 14px; box-sizing: border-box;
      border: 1px solid rgba(148, 163, 184, 0.2); border-radius: 8px;
      background: rgba(30, 41, 59, 0.6); color: #f8fafc; font-size: 13px;
      transition: all 0.2s; outline: none;
    }
    .input-industrial:focus {
      border-color: #38bdf8;
      box-shadow: 0 0 0 3px rgba(56, 189, 248, 0.15);
      background: rgba(30, 41, 59, 0.9);
    }
    select.input-industrial option { background: #0f172a; color: #f8fafc; }
    .nav-industrial {
      background: rgba(15, 23, 42, 0.85); backdrop-filter: blur(12px);
      padding: 14px 24px; border-radius: 12px; border: 1px solid rgba(56, 189, 248, 0.2);
      box-shadow: 0 8px 32px rgba(0,0,0,0.4); display: flex; justify-content: space-between;
      align-items: center; margin-bottom: 28px; flex-wrap: wrap; gap: 15px;
    }
    .logo-contenedor { background: #ffffff; padding: 6px 10px; border-radius: 8px; display: flex; align-items: center; box-shadow: 0 2px 8px rgba(0,0,0,0.2); }
    .nav-industrial a { text-decoration: none; color: #94a3b8; font-weight: 500; padding: 8px 12px; border-radius: 6px; font-size: 12px; text-transform: uppercase; letter-spacing: 0.05em; transition: all 0.2s; display: inline-flex; align-items: center; gap: 6px; }
    .nav-industrial a:hover { background: rgba(56, 189, 248, 0.1); color: #38bdf8; }
    .nav-industrial .logo-area { display: flex; align-items: center; gap: 15px; }
    .nav-industrial .user-info { color: #94a3b8; font-size: 11px; text-align: right; border-left: 1px solid rgba(255,255,255,0.1); padding-left: 15px; }
    .nav-industrial .user-info b { color: #f8fafc; font-size: 12px; }
    table { width: 100%; border-collapse: separate; border-spacing: 0; font-size: 13px; }
    th { background-color: rgba(30, 41, 59, 0.9); color: #38bdf8; padding: 12px; font-weight: 600; text-align: left; border-bottom: 2px solid rgba(56, 189, 248, 0.2); text-transform: uppercase; font-size: 11px; letter-spacing: 0.05em; }
    th:first-child { border-top-left-radius: 8px; }
    th:last-child { border-top-right-radius: 8px; }
    td { padding: 12px; border-bottom: 1px solid rgba(148, 163, 184, 0.1); color: #cbd5e1; }
    tr:hover td { background: rgba(30, 41, 59, 0.3); }
    .badge-tech {
      padding: 4px 10px; border-radius: 6px; font-weight: 600; font-size: 11px;
      display: inline-flex; align-items: center; gap: 5px; letter-spacing: 0.03em;
    }
  </style>
`;

function verificarSesion(req, res, next) {
    if (req.session && req.session.usuario) return next();
    return res.redirect('/login');
}

function generarEncabezado(titulo, req) {
    const u = req.session && req.session.usuario ? req.session.usuario : { username: 'Invitado', departamento: 'General', rol: 'Invitado' };
    const esAdminTotal = u.rol === 'Administrador' || u.departamento === 'Administración' || u.departamento === 'Control Administrativo';
    const esConta = u.departamento === 'Contabilidad';
    const permiteClientes = esAdminTotal || u.departamento === 'Ventas' || esConta;
    const permiteSoporte = esAdminTotal || u.departamento === 'Ventas' || esConta;
    const permiteCotiz = esAdminTotal || u.departamento === 'Ventas' || esConta;
    const permitePagos = esAdminTotal || u.departamento === 'Ventas' || esConta;
    const permiteMetricas = esAdminTotal || u.departamento === 'Ventas' || esConta;
    const permiteContaPanel = esAdminTotal || esConta;
    const permiteAlmacen = esAdminTotal || u.departamento === 'Almacén';
    const permiteAdminTotal = esAdminTotal || esConta; // Contabilidad ahora tiene acceso al panel de Control Administrativo

    return `
    ${EstilosGlobales}
    <div class="nav-industrial">
        <div class="logo-area">
          <div class="logo-contenedor"><img src="/logo.png" alt="Logo" style="height: 35px; display: block;"></div>
          <div><h2 style="color: white; margin: 0; font-size: 16px; font-weight: 600;"><i class="fa-solid fa-microchip" style="color: #38bdf8; margin-right: 8px;"></i>${titulo}</h2></div>
        </div>
        <div class="user-info"><i class="fa-solid fa-user-shield" style="color: #38bdf8;"></i> <b>${u.username}</b><br><span>Depto: ${u.departamento}</span></div>
        <nav style="display: flex; align-items: center; gap: 4px; flex-wrap: wrap;">
          ${permiteClientes ? `<a href="/"><i class="fa-solid fa-address-book"></i> Clientes</a>` : ''}
          ${permiteSoporte ? `<a href="/soporte"><i class="fa-solid fa-headset"></i> Soporte</a>` : ''}
          ${permiteCotiz ? `<a href="/cotizaciones"><i class="fa-solid fa-file-invoice-dollar"></i> Cotiz.</a>` : ''}
          ${permitePagos ? `<a href="/pagos-clientes"><i class="fa-solid fa-wallet" style="color:#34d399;"></i> Pagos</a>` : ''}
          ${permiteMetricas ? `<a href="/gastos-ventas"><i class="fa-solid fa-chart-line"></i> Métricas</a>` : ''}
          ${permiteContaPanel ? `<a href="/contabilidad" style="color:#fbbf24;"><i class="fa-solid fa-vault"></i> Conta</a>` : ''}
          ${permiteAlmacen ? `<a href="/almacen-logistica" style="color:#4ade80;"><i class="fa-solid fa-warehouse"></i> Almacén</a>` : ''}
          <a href="/chat" style="color:#38bdf8;"><i class="fa-solid fa-comments"></i> Chat</a>
          <a href="/productos"><i class="fa-solid fa-boxes-stacked"></i> Catálogos</a>
          ${permiteAdminTotal ? `<a href="/control-administrativo" style="color: #f43f5e;"><i class="fa-solid fa-shield-halved"></i> Admin</a>` : ''}
          ${u.rol === 'Administrador' || u.departamento === 'Administración' ? `<a href="/usuarios" style="color: #4ade80;"><i class="fa-solid fa-users-gear"></i> Accesos</a>` : ''}
          <a href="/logout" style="background: rgba(239, 68, 68, 0.2); color: #fca5a5; border: 1px solid rgba(239, 68, 68, 0.4);"><i class="fa-solid fa-power-off"></i> Salir</a>
        </nav>
    </div>
  `;
}

// LOGIN
app.get('/login', (req, res) => {
    res.send(`
    ${EstilosGlobales}
    <html>
      <body style="display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0;">
        <div class="contenedor-tarjeta" style="width: 360px; text-align: center; padding: 40px; border: 1px solid rgba(56, 189, 248, 0.3);">
          <div style="background: white; display: inline-block; padding: 12px; border-radius: 10px; margin-bottom: 20px; box-shadow: 0 4px 15px rgba(0,0,0,0.3);">
            <img src="/logo.png" style="height: 55px; display: block;">
          </div>
          <h3 style="margin-top: 0; margin-bottom: 5px; color: #f8fafc;">Portal ERP Industrial</h3>
          <p style="font-size: 11px; color: #94a3b8; margin-bottom: 25px; text-transform: uppercase; letter-spacing: 0.1em;">C.C. Instalaciones Industriales</p>
          <form action="/login" method="POST">
            <div style="position: relative;"><i class="fa-solid fa-user" style="position: absolute; left: 14px; top: 14px; color: #64748b;"></i><input type="text" name="username" required placeholder="Usuario" class="input-industrial" style="padding-left: 40px;"></div>
            <div style="position: relative;"><i class="fa-solid fa-lock" style="position: absolute; left: 14px; top: 14px; color: #64748b;"></i><input type="password" name="password" required placeholder="Contraseña" class="input-industrial" style="padding-left: 40px;"></div>
            <button type="submit" class="btn-primario" style="width: 100%; padding: 12px; margin-top: 10px; justify-content: center;"><i class="fa-solid fa-fingerprint"></i> Autenticar Acceso</button>
          </form>
        </div>
      </body>
    </html>
  `);
});

app.post('/login', (req, res) => {
    const { username, password } = req.body;
    db.get("SELECT * FROM usuarios WHERE username = ? AND password = ?", [username, password], (err, usuario) => {
        if (usuario) {
            req.session.usuario = usuario;
            res.redirect('/');
        } else {
            res.send("<script>alert('Credenciales inválidas'); window.location.href='/login';</script>");
        }
    });
});

app.get('/logout', (req, res) => {
    req.session.destroy(() => res.redirect('/login'));
});

// GESTIÓN DE USUARIOS
app.get('/usuarios', verificarSesion, (req, res) => {
    const u = req.session.usuario;
    if (u.rol !== 'Administrador' && u.departamento !== 'Administración' && u.departamento !== 'Control Administrativo') {
        return res.send("<script>alert('Acceso denegado.'); window.location.href='/';</script>");
    }

    db.all("SELECT * FROM usuarios ORDER BY id DESC", [], (err, filas) => {
        let tabla = '';
        filas.forEach(usr => {
            tabla += `<tr>
        <td style="padding: 12px;"><i class="fa-solid fa-user-gear" style="color: #38bdf8; margin-right: 6px;"></i><strong>${usr.username}</strong></td>
        <td style="padding: 12px;"><span class="badge-tech" style="background: rgba(56, 189, 248, 0.1); color: #38bdf8;">${usr.departamento}</span></td>
        <td style="padding: 12px;">${usr.rol}</td>
        <td style="padding: 12px;">
          <a href="/editar-usuario/${usr.id}" class="btn-primario" style="padding: 5px 10px; font-size: 10px; background: linear-gradient(135deg, #d97706, #b45309); text-decoration: none; display: inline-block; margin-bottom: 2px;"><i class="fa-solid fa-pen-to-square"></i> Editar</a>
          ${usr.username !== 'AdminCCI' ? `<a href="/eliminar-usuario/${usr.id}" onclick="return confirm('¿Borrar usuario?')" class="btn-peligro" style="padding: 5px 10px; font-size: 10px;"><i class="fa-solid fa-trash"></i> Borrar</a>` : '<span style="color:#64748b; font-size:11px;"><i class="fa-solid fa-shield"></i> Protegido</span>'}
        </td>
      </tr>`;
        });

        res.send(`
      <html>
        <body style="max-width: 1200px; margin: auto;">
          ${generarEncabezado('Control de Accesos por Departamento', req)}
          <div style="display: flex; gap: 24px; flex-wrap: wrap;">
            <div class="contenedor-tarjeta" style="width: 340px; height: fit-content;">
              <h3 style="margin-top:0;"><i class="fa-solid fa-user-plus" style="color: #38bdf8; margin-right: 8px;"></i>Crear Credencial</h3>
              <form action="/guardar-usuario" method="POST">
                <input type="text" name="username" required placeholder="Nombre de Usuario" class="input-industrial">
                <input type="password" name="password" required placeholder="Contraseña de Acceso" class="input-industrial">
                <label style="font-size:11px; font-weight:600; color: #94a3b8; text-transform: uppercase;">Departamento:</label>
                <select name="departamento" class="input-industrial" style="margin-top:5px;">
                  <option value="Administración">Administración</option>
                  <option value="Control Administrativo">Control Administrativo</option>
                  <option value="Ventas">Ventas</option>
                  <option value="Contabilidad">Contabilidad</option>
                  <option value="Almacén">Almacén</option>
                  <option value="Presupuestos">Presupuestos Generales</option>
                </select>
                <label style="font-size:11px; font-weight:600; color: #94a3b8; text-transform: uppercase;">Nivel de Rol:</label>
                <select name="rol" class="input-industrial" style="margin-top:5px; margin-bottom: 18px;">
                  <option value="Usuario">Operativo / Estándar</option>
                  <option value="Administrador">Administrador Total</option>
                </select>
                <button type="submit" class="btn-primario" style="width: 100%; justify-content: center;"><i class="fa-solid fa-floppy-disk"></i> Registrar Usuario</button>
              </form>
            </div>
            <div class="contenedor-tarjeta" style="flex-grow: 1; overflow-x: auto;">
              <h3 style="margin-top:0;"><i class="fa-solid fa-network-wired" style="color: #38bdf8; margin-right: 8px;"></i>Usuarios Activos</h3>
              <table>
                <tr>
                  <th>Usuario</th>
                  <th>Departamento</th>
                  <th>Rol</th>
                  <th>Acciones</th>
                </tr>
                ${tabla}
              </table>
            </div>
          </div>
        </body>
      </html>
    `);
    });
});

app.post('/guardar-usuario', verificarSesion, (req, res) => {
    const u = req.session.usuario;
    if (u.rol !== 'Administrador' && u.departamento !== 'Administración' && u.departamento !== 'Control Administrativo') return res.send("Acceso denegado");
    const { username, password, departamento, rol } = req.body;
    db.run("INSERT INTO usuarios (username, password, departamento, rol) VALUES (?, ?, ?, ?)", [username, password, departamento, rol], () => {
        res.redirect('/usuarios');
    });
});

app.get('/editar-usuario/:id', verificarSesion, (req, res) => {
    const u = req.session.usuario;
    if (u.rol !== 'Administrador' && u.departamento !== 'Administración' && u.departamento !== 'Control Administrativo') return res.send("Acceso denegado");

    db.get("SELECT * FROM usuarios WHERE id = ?", [req.params.id], (err, usr) => {
        if (!usr) return res.send("Usuario no encontrado");
        res.send(`
      <html>
        <body style="padding-top: 30px; max-width: 500px; margin: auto;">
          ${generarEncabezado('Modificar Credencial', req)}
          <div class="contenedor-tarjeta">
            <h3 style="color:#38bdf8; margin-top:0;"><i class="fa-solid fa-pen"></i> Editar: ${usr.username}</h3>
            <form action="/actualizar-usuario/${usr.id}" method="POST">
              <label style="font-size:11px; font-weight:600; color: #94a3b8; text-transform: uppercase;">Usuario:</label>
              <input type="text" name="username" value="${usr.username}" required class="input-industrial" style="margin-top:5px;">
              
              <label style="font-size:11px; font-weight:600; color: #94a3b8; text-transform: uppercase;">Nueva Contraseña (Opcional):</label>
              <input type="password" name="password" placeholder="Dejar en blanco para mantener actual" class="input-industrial" style="margin-top:5px;">
              
              <label style="font-size:11px; font-weight:600; color: #94a3b8; text-transform: uppercase;">Departamento:</label>
              <select name="departamento" class="input-industrial" style="margin-top:5px;">
                <option value="Administración" ${usr.departamento === 'Administración' ? 'selected' : ''}>Administración</option>
                <option value="Control Administrativo" ${usr.departamento === 'Control Administrativo' ? 'selected' : ''}>Control Administrativo</option>
                <option value="Ventas" ${usr.departamento === 'Ventas' ? 'selected' : ''}>Ventas</option>
                <option value="Contabilidad" ${usr.departamento === 'Contabilidad' ? 'selected' : ''}>Contabilidad</option>
                <option value="Almacén" ${usr.departamento === 'Almacén' ? 'selected' : ''}>Almacén</option>
                <option value="Presupuestos" ${usr.departamento === 'Presupuestos' ? 'selected' : ''}>Presupuestos Generales</option>
              </select>

              <label style="font-size:11px; font-weight:600; color: #94a3b8; text-transform: uppercase;">Rol:</label>
              <select name="rol" class="input-industrial" style="margin-top:5px; margin-bottom: 20px;">
                <option value="Usuario" ${usr.rol === 'Usuario' ? 'selected' : ''}>Operativo</option>
                <option value="Administrador" ${usr.rol === 'Administrador' ? 'selected' : ''}>Administrador</option>
              </select>

              <button type="submit" class="btn-primario" style="width:100%; background: linear-gradient(135deg, #10b981, #059669); justify-content: center;"><i class="fa-solid fa-check"></i> Guardar Cambios</button>
            </form>
            <br><a href="/usuarios" style="color:#38bdf8; font-weight:600; text-decoration:none; display:block; text-align:center; font-size: 12px;"><i class="fa-solid fa-arrow-left"></i> Regresar</a>
          </div>
        </body>
      </html>
    `);
    });
});

app.post('/actualizar-usuario/:id', verificarSesion, (req, res) => {
    const u = req.session.usuario;
    if (u.rol !== 'Administrador' && u.departamento !== 'Administración' && u.departamento !== 'Control Administrativo') return res.send("Acceso denegado");
    const { username, password, departamento, rol } = req.body;
    const userId = req.params.id;

    if (password && password.trim() !== '') {
        db.run("UPDATE usuarios SET username = ?, password = ?, departamento = ?, rol = ? WHERE id = ?",
            [username, password, departamento, rol, userId], () => res.redirect('/usuarios'));
    } else {
        db.run("UPDATE usuarios SET username = ?, departamento = ?, rol = ? WHERE id = ?",
            [username, departamento, rol, userId], () => res.redirect('/usuarios'));
    }
});

app.get('/eliminar-usuario/:id', verificarSesion, (req, res) => {
    db.run("DELETE FROM usuarios WHERE id = ? AND username != 'AdminCCI'", [req.params.id], () => res.redirect('/usuarios'));
});

// CONTROL ADMINISTRATIVO SUPREMO (Accesible por Admin y Contabilidad)
app.get('/control-administrativo', verificarSesion, (req, res) => {
    const u = req.session.usuario;
    if (u.departamento !== 'Control Administrativo' && u.rol !== 'Administrador' && u.departamento !== 'Administración' && u.departamento !== 'Contabilidad') {
        return res.send("<script>alert('Acceso exclusivo para Control Administrativo y Contabilidad'); window.location.href='/';</script>");
    }

    db.all("SELECT * FROM personal ORDER BY id DESC", [], (err, empleados) => {
        db.all("SELECT g.*, u.username FROM gastos_operativos g LEFT JOIN usuarios u ON g.vendedor = u.username ORDER BY g.id DESC", [], (err, gastos) => {
            db.all("SELECT crm.*, cl.nombre_empresa FROM crm_agenda crm JOIN clientes cl ON crm.cliente_id = cl.id ORDER BY crm.fecha_programada ASC", [], (err, agenda) => {
                db.all("SELECT c.*, cl.nombre_empresa, SUM(d.cantidad * d.precio_unitario) as subtotal FROM cotizaciones c JOIN clientes cl ON c.cliente_id = cl.id LEFT JOIN cotizacion_detalles d ON c.id = d.cotizacion_id GROUP BY c.id ORDER BY c.id DESC", [], (err, cotes) => {

                    let tablaPersonal = '';
                    empleados.forEach(emp => {
                        tablaPersonal += `<tr>
                            <td style="padding:10px;"><i class="fa-solid fa-id-badge" style="color:#38bdf8;"></i> <strong>${emp.numero_trabajador}</strong></td>
                            <td style="padding:10px;">${emp.nombre_completo}<br><small style="color:#94a3b8;">${emp.puesto}</small></td>
                            <td style="padding:10px;">${emp.telefono || 'S/N'}<br><small>${emp.correo || ''}</small></td>
                            <td style="padding:10px;">RFC: ${emp.rfc || 'S/N'}<br>NSS: ${emp.nss || 'S/N'}</td>
                            <td style="padding:10px;">
                                ${emp.archivo_ine ? `<a href="/personal/${emp.archivo_ine}" target="_blank" style="font-size:11px; display:block; color:#38bdf8;"><i class="fa-solid fa-file-pdf"></i> INE</a>` : ''}
                                ${emp.archivo_curp ? `<a href="/personal/${emp.archivo_curp}" target="_blank" style="font-size:11px; display:block; color:#38bdf8;"><i class="fa-solid fa-file-pdf"></i> CURP</a>` : ''}
                                ${emp.archivo_banco ? `<a href="/personal/${emp.archivo_banco}" target="_blank" style="font-size:11px; display:block; color:#38bdf8;"><i class="fa-solid fa-credit-card"></i> Banco</a>` : ''}
                                ${emp.archivo_cv ? `<a href="/personal/${emp.archivo_cv}" target="_blank" style="font-size:11px; display:block; color:#38bdf8;"><i class="fa-solid fa-briefcase"></i> CV</a>` : ''}
                            </td>
                            <td style="padding:10px;">
                                <a href="/editar-empleado/${emp.id}" class="btn-primario" style="padding:4px 8px; font-size:10px; background: linear-gradient(135deg, #d97706, #b45309); text-decoration:none;"><i class="fa-solid fa-pen"></i></a>
                            </td>
                        </tr>`;
                    });

                    let tablaGastos = '';
                    let gastoTotalGlobal = 0;
                    gastos.forEach(g => {
                        gastoTotalGlobal += (g.monto || 0);
                        tablaGastos += `<tr>
                            <td style="padding:10px;"><span class="badge-tech" style="background:rgba(56,189,248,0.1); color:#38bdf8;">${g.vendedor || 'General'}</span></td>
                            <td style="padding:10px;"><strong>${g.concepto}</strong><br><small style="color:#94a3b8;">${g.categoria}</small></td>
                            <td style="padding:10px; color:#f87171; font-weight:600;">$${formatoDinero(g.monto)}</td>
                            <td style="padding:10px;">${g.fecha}</td>
                            <td style="padding:10px;">${g.comprobante ? `<a href="/gastos/${g.comprobante}" target="_blank" style="color:#38bdf8;"><i class="fa-solid fa-receipt"></i> Ver</a>` : 'S/C'}</td>
                        </tr>`;
                    });

                    let tablaAgendaAdmin = '';
                    agenda.forEach(ag => {
                        tablaAgendaAdmin += `<tr>
                            <td style="padding:10px;"><strong>${ag.nombre_empresa}</strong></td>
                            <td style="padding:10px;"><span class="badge-tech" style="background:rgba(56,189,248,0.1); color:#38bdf8;">${ag.vendedor}</span></td>
                            <td style="padding:10px;">${ag.tipo} <small>(${ag.fecha_programada})</small></td>
                            <td style="padding:10px;">${ag.comentario}</td>
                            <td style="padding:10px;"><b>${ag.estado}</b></td>
                        </tr>`;
                    });

                    let tablaCotAdmin = '';
                    let totalVentasAceptadas = 0;
                    let totalVentasRechazadas = 0;
                    cotes.forEach(cot => {
                        const sub = (cot.subtotal || 0) * (1 - ((cot.descuento || 0) / 100));
                        const total = sub * 1.16;
                        if (cot.estado === 'Aceptada') totalVentasAceptadas += total;
                        if (cot.estado === 'Rechazada') totalVentasRechazadas += total;

                        tablaCotAdmin += `<tr>
                            <td style="padding:10px;">COT-${cot.id} <small>(${cot.vendedor})</small></td>
                            <td style="padding:10px;"><strong>${cot.nombre_empresa}</strong></td>
                            <td style="padding:10px;">${cot.estado}</td>
                            <td style="padding:10px;">$${formatoDinero(total)} ${cot.moneda}</td>
                            <td style="padding:10px;">
                                <a href="/gestionar-estatus/${cot.id}" class="btn-primario" style="padding:4px 8px; font-size:10px; background:linear-gradient(135deg, #7c3aed, #6d28d9); text-decoration:none;"><i class="fa-solid fa-gears"></i></a>
                            </td>
                        </tr>`;
                    });

                    res.send(`
                        <html>
                            <body style="max-width: 1450px; margin: auto;">
                                ${generarEncabezado('Control Administrativo (Supervisión Total)', req)}
                                
                                <div style="display:flex; gap:20px; margin-bottom:24px; flex-wrap:wrap;">
                                    <div class="contenedor-tarjeta" style="flex:1; text-align:center; background:rgba(3, 105, 161, 0.2); border-color: rgba(3, 105, 161, 0.4);">
                                        <h4 style="margin:0; color:#38bdf8; font-size:12px; text-transform:uppercase;"><i class="fa-solid fa-circle-check"></i> Ventas Aceptadas</h4>
                                        <p style="font-size:24px; font-weight:700; margin:8px 0; color:#f8fafc;">$${formatoDinero(totalVentasAceptadas)} <span style="font-size:12px;">MXN</span></p>
                                    </div>
                                    <div class="contenedor-tarjeta" style="flex:1; text-align:center; background:rgba(185, 28, 28, 0.2); border-color: rgba(185, 28, 28, 0.4);">
                                        <h4 style="margin:0; color:#f87171; font-size:12px; text-transform:uppercase;"><i class="fa-solid fa-circle-xmark"></i> Ventas Rechazadas</h4>
                                        <p style="font-size:24px; font-weight:700; margin:8px 0; color:#f8fafc;">$${formatoDinero(totalVentasRechazadas)} <span style="font-size:12px;">MXN</span></p>
                                    </div>
                                    <div class="contenedor-tarjeta" style="flex:1; text-align:center; background:rgba(217, 119, 6, 0.2); border-color: rgba(217, 119, 6, 0.4);">
                                        <h4 style="margin:0; color:#fbbf24; font-size:12px; text-transform:uppercase;"><i class="fa-solid fa-wallet"></i> Gastos Registrados</h4>
                                        <p style="font-size:24px; font-weight:700; margin:8px 0; color:#f8fafc;">$${formatoDinero(gastoTotalGlobal)} <span style="font-size:12px;">MXN</span></p>
                                    </div>
                                </div>

                                <div style="display: flex; gap: 24px; margin-bottom: 24px; flex-wrap: wrap;">
                                    <div class="contenedor-tarjeta" style="width: 380px; height: fit-content;">
                                        <h3 style="margin-top:0;"><i class="fa-solid fa-user-plus" style="color:#38bdf8; margin-right:8px;"></i>Expediente de Personal</h3>
                                        <form action="/guardar-empleado" method="POST" enctype="multipart/form-data">
                                            <input type="text" name="numero_trabajador" required placeholder="No. Trabajador (Ej. TR-001)" class="input-industrial">
                                            <input type="text" name="nombre_completo" required placeholder="Nombre Completo" class="input-industrial">
                                            <input type="text" name="puesto" required placeholder="Puesto / Cargo" class="input-industrial">
                                            <input type="text" name="telefono" placeholder="Teléfono" class="input-industrial">
                                            <input type="text" name="rfc" placeholder="RFC" class="input-industrial">
                                            <input type="text" name="nss" placeholder="NSS" class="input-industrial">
                                            <input type="text" name="direccion" placeholder="Dirección" class="input-industrial">
                                            <input type="email" name="correo" placeholder="Correo Electrónico" class="input-industrial">
                                            <input type="text" name="contacto_emergencia" placeholder="Contacto de Emergencia y Tel." class="input-industrial">
                                            
                                            <label style="font-size:11px; font-weight:600; color:#94a3b8;">INE Adjunto:</label>
                                            <input type="file" name="archivo_ine" accept=".pdf,.jpg,.png" class="input-industrial" style="font-size:11px; padding:6px;">
                                            <label style="font-size:11px; font-weight:600; color:#94a3b8;">CURP Adjunto:</label>
                                            <input type="file" name="archivo_curp" accept=".pdf,.jpg,.png" class="input-industrial" style="font-size:11px; padding:6px;">
                                            <label style="font-size:11px; font-weight:600; color:#94a3b8;">Estado de Cuenta:</label>
                                            <input type="file" name="archivo_banco" accept=".pdf,.jpg,.png" class="input-industrial" style="font-size:11px; padding:6px;">
                                            <label style="font-size:11px; font-weight:600; color:#94a3b8;">CV Profesional:</label>
                                            <input type="file" name="archivo_cv" accept=".pdf,.jpg,.png" class="input-industrial" style="font-size:11px; padding:6px; margin-bottom:15px;">

                                            <button type="submit" class="btn-primario" style="width:100%; background:linear-gradient(135deg, #10b981, #059669); justify-content:center;"><i class="fa-solid fa-floppy-disk"></i> Guardar Expediente</button>
                                        </form>
                                    </div>
                                    <div class="contenedor-tarjeta" style="flex-grow: 1; overflow-x: auto;">
                                        <h3 style="margin-top:0;"><i class="fa-solid fa-id-card" style="color:#38bdf8; margin-right:8px;"></i>Directorio y Documentación de Personal</h3>
                                        <table>
                                            <tr>
                                                <th>No.</th>
                                                <th>Nombre & Puesto</th>
                                                <th>Contacto</th>
                                                <th>Fiscales</th>
                                                <th>Documentos</th>
                                                <th>Acción</th>
                                            </tr>
                                            ${tablaPersonal}
                                        </table>
                                    </div>
                                </div>

                                <div style="display:flex; gap:24px; flex-wrap:wrap;">
                                    <div class="contenedor-tarjeta" style="flex:1; overflow-x:auto;">
                                        <h3 style="margin-top:0;"><i class="fa-solid fa-receipt" style="color:#38bdf8; margin-right:8px;"></i>Gastos Operativos Registrados</h3>
                                        <table>
                                            <tr>
                                                <th>Usuario</th>
                                                <th>Concepto</th>
                                                <th>Monto</th>
                                                <th>Fecha</th>
                                                <th>Comprobante</th>
                                            </tr>
                                            ${tablaGastos}
                                        </table>
                                    </div>
                                    <div class="contenedor-tarjeta" style="flex:1; overflow-x:auto;">
                                        <h3 style="margin-top:0;"><i class="fa-solid fa-clipboard-list" style="color:#38bdf8; margin-right:8px;"></i>Revisión de Cotizaciones</h3>
                                        <table>
                                            <tr>
                                                <th>Folio</th>
                                                <th>Cliente</th>
                                                <th>Estatus</th>
                                                <th>Total</th>
                                                <th>Gestión</th>
                                            </tr>
                                            ${tablaCotAdmin}
                                        </table>
                                    </div>
                                </div>

                                <div class="contenedor-tarjeta" style="overflow-x: auto; margin-top: 24px;">
                                    <h3 style="margin-top:0;"><i class="fa-solid fa-calendar-days" style="color:#38bdf8; margin-right:8px;"></i>Cronograma General de CRM</h3>
                                    <table>
                                        <tr>
                                            <th>Cliente</th>
                                            <th>Vendedor</th>
                                            <th>Tipo & Fecha</th>
                                            <th>Comentario</th>
                                            <th>Estado</th>
                                        </tr>
                                        ${tablaAgendaAdmin}
                                    </table>
                                </div>
                            </body>
                        </html>
                    `);
                });
            });
        });
    });
});

const uploadPersonalFiles = upload.fields([
    { name: 'archivo_ine', maxCount: 1 },
    { name: 'archivo_curp', maxCount: 1 },
    { name: 'archivo_banco', maxCount: 1 },
    { name: 'archivo_cv', maxCount: 1 }
]);

app.post('/guardar-empleado', verificarSesion, uploadPersonalFiles, (req, res) => {
    const u = req.session.usuario;
    if (u.departamento !== 'Control Administrativo' && u.rol !== 'Administrador' && u.departamento !== 'Administración' && u.departamento !== 'Contabilidad') return res.send("Acceso denegado");

    const { numero_trabajador, nombre_completo, puesto, telefono, rfc, nss, direccion, correo, contacto_emergencia } = req.body;

    let ine = null, curp = null, banco = null, cv = null;
    if (req.files) {
        if (req.files['archivo_ine']) { ine = req.files['archivo_ine'][0].filename + path.extname(req.files['archivo_ine'][0].originalname); fs.renameSync(req.files['archivo_ine'][0].path, path.join(__dirname, 'uploads/personal', ine)); }
        if (req.files['archivo_curp']) { curp = req.files['archivo_curp'][0].filename + path.extname(req.files['archivo_curp'][0].originalname); fs.renameSync(req.files['archivo_curp'][0].path, path.join(__dirname, 'uploads/personal', curp)); }
        if (req.files['archivo_banco']) { banco = req.files['archivo_banco'][0].filename + path.extname(req.files['archivo_banco'][0].originalname); fs.renameSync(req.files['archivo_banco'][0].path, path.join(__dirname, 'uploads/personal', banco)); }
        if (req.files['archivo_cv']) { cv = req.files['archivo_cv'][0].filename + path.extname(req.files['archivo_cv'][0].originalname); fs.renameSync(req.files['archivo_cv'][0].path, path.join(__dirname, 'uploads/personal', cv)); }
    }

    db.run("INSERT INTO personal (numero_trabajador, nombre_completo, puesto, telefono, rfc, nss, direccion, correo, contacto_emergencia, archivo_ine, archivo_curp, archivo_banco, archivo_cv) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
        [numero_trabajador, nombre_completo, puesto, telefono, rfc, nss, direccion, correo, contacto_emergencia, ine, curp, banco, cv], () => {
            res.redirect('/control-administrativo');
        });
});

app.get('/editar-empleado/:id', verificarSesion, (req, res) => {
    db.get("SELECT * FROM personal WHERE id = ?", [req.params.id], (err, emp) => {
        if (!emp) return res.send("Empleado no encontrado");
        res.send(`
            <html>
                <body style="padding-top:30px; max-width:600px; margin:auto;">
                    ${generarEncabezado('Modificar Trabajador', req)}
                    <div class="contenedor-tarjeta">
                        <h3 style="color:#38bdf8; margin-top:0;"><i class="fa-solid fa-user-pen"></i> Editar: ${emp.nombre_completo}</h3>
                        <form action="/actualizar-empleado/${emp.id}" method="POST">
                            <label style="font-size:11px; font-weight:600; color:#94a3b8;">Puesto:</label>
                            <input type="text" name="puesto" value="${emp.puesto}" required class="input-industrial">
                            <label style="font-size:11px; font-weight:600; color:#94a3b8;">Teléfono:</label>
                            <input type="text" name="telefono" value="${emp.telefono || ''}" class="input-industrial">
                            <label style="font-size:11px; font-weight:600; color:#94a3b8;">Correo:</label>
                            <input type="email" name="correo" value="${emp.correo || ''}" class="input-industrial">
                            <label style="font-size:11px; font-weight:600; color:#94a3b8;">Contacto de Emergencia:</label>
                            <input type="text" name="contacto_emergencia" value="${emp.contacto_emergencia || ''}" class="input-industrial" style="margin-bottom:20px;">
                            <button type="submit" class="btn-primario" style="width:100%; background:linear-gradient(135deg, #10b981, #059669); justify-content:center;"><i class="fa-solid fa-check"></i> Guardar Cambios</button>
                        </form>
                        <br><a href="/control-administrativo" style="color:#38bdf8; font-weight:600; text-decoration:none; font-size:12px;"><i class="fa-solid fa-arrow-left"></i> Regresar</a>
                    </div>
                </body>
            </html>
        `);
    });
});

app.post('/actualizar-empleado/:id', verificarSesion, (req, res) => {
    const { puesto, telefono, correo, contacto_emergencia } = req.body;
    db.run("UPDATE personal SET puesto = ?, telefono = ?, correo = ?, contacto_emergencia = ? WHERE id = ?", [puesto, telefono, correo, contacto_emergencia, req.params.id], () => {
        res.redirect('/control-administrativo');
    });
});

// CLIENTES & CRM
app.get('/', verificarSesion, (req, res) => {
    const u = req.session.usuario;
    const esAdmin = u.rol === 'Administrador' || u.departamento === 'Administración' || u.departamento === 'Control Administrativo' || u.departamento === 'Contabilidad';
    const mesSeleccionado = req.query.mes || new Date().toISOString().slice(0, 7);

    let queryClientes = esAdmin ? "SELECT * FROM clientes ORDER BY id DESC" : "SELECT * FROM clientes WHERE vendedor_asignado = ? OR vendedor_asignado = ? ORDER BY id DESC";
    let paramsClientes = esAdmin ? [] : [u.username, u.departamento];

    db.all(queryClientes, paramsClientes, (err, clientes) => {
        let queryAgenda = esAdmin ?
            "SELECT crm.*, cl.nombre_empresa FROM crm_agenda crm JOIN clientes cl ON crm.cliente_id = cl.id ORDER BY crm.fecha_programada ASC" :
            "SELECT crm.*, cl.nombre_empresa FROM crm_agenda crm JOIN clientes cl ON crm.cliente_id = cl.id WHERE crm.vendedor = ? OR crm.vendedor = ? ORDER BY crm.fecha_programada ASC";
        let paramsAgenda = esAdmin ? [] : [u.username, u.departamento];

        db.all(queryAgenda, paramsAgenda, (err, agendas) => {
            db.all("SELECT * FROM calendario_notas WHERE mes = ? ORDER BY id DESC", [mesSeleccionado], (err, notasMes) => {

                const ahora = new Date();
                let alertasRojasCount = 0;
                let alertasAmarillasCount = 0;

                let tablaClientes = '';
                clientes.forEach(c => {
                    let telLimpio = (c.telefono || '').replace(/\D/g, '');
                    let infoCredito = c.tiene_credito === 'Sí' ? `<span style="color:#38bdf8; font-weight:600;"><i class="fa-solid fa-circle-check"></i> Sí ($${formatoDinero(c.monto_credito)})</span>` : `<span style="color:#94a3b8;">Sin Crédito</span>`;

                    tablaClientes += `<tr>
              <td style="padding: 12px;">${c.numero_cliente || '-'}</td>
              <td style="padding: 12px;"><strong>${c.nombre_empresa}</strong><br><small style="color:#94a3b8;">Vend: ${c.vendedor_asignado || 'General'}</small></td>
              <td style="padding: 12px;">
                ${infoCredito}<br>
                <small style="color:#a78bfa; font-size:10px;"><i class="fa-solid fa-shield-halved"></i> Contabilidad / Ventas</small>
              </td>
              <td style="padding: 12px;">
                ${c.telefono ? `<a href="https://wa.me/${telLimpio}" target="_blank" rel="noopener noreferrer" style="color:#4ade80; font-weight:600; text-decoration:none;"><i class="fa-brands fa-whatsapp"></i> ${c.telefono}</a>` : 'S/Tel'}
              </td>
              <td style="padding: 12px;">${c.archivo_csf ? `<a href="/csf/${c.archivo_csf}" target="_blank" rel="noopener noreferrer" style="color:#a78bfa; font-weight:600;"><i class="fa-solid fa-file-pdf"></i> Ver CSF</a>` : '<span style="color:#fbbf24; font-size:11px;">Pendiente</span>'}</td>
              <td style="padding: 12px;">
                <a href="/cliente-detalle/${c.id}" class="btn-primario" style="padding:4px 8px; font-size:10px; background:linear-gradient(135deg, #0ea5e9, #0284c7); text-decoration:none; display:inline-block; margin-bottom:2px;"><i class="fa-solid fa-eye"></i></a>
                <a href="/editar-cliente/${c.id}" class="btn-primario" style="padding:4px 8px; font-size:10px; background:linear-gradient(135deg, #d97706, #b45309); text-decoration:none; display:inline-block; margin-bottom:2px;"><i class="fa-solid fa-pen"></i></a>
                <a href="/eliminar-cliente/${c.id}" onclick="return confirm('¿Estás seguro de eliminar este cliente? Se borrará de la cartera.')" class="btn-peligro" style="padding:4px 8px; font-size:10px;"><i class="fa-solid fa-trash"></i></a>
              </td>
            </tr>`;
                });

                let tablaAgenda = '';
                agendas.forEach(ag => {
                    let fechaCita = new Date(ag.fecha_programada);
                    let estadoColorBg = 'rgba(148, 163, 184, 0.1)';
                    let estadoColorText = '#94a3b8';
                    let estadoTexto = ag.estado || 'Pendiente';

                    if (ag.estado === 'Completado') {
                        estadoColorBg = 'rgba(16, 185, 129, 0.15)';
                        estadoColorText = '#34d399';
                        estadoTexto = 'Completado';
                    } else if (fechaCita < ahora) {
                        estadoColorBg = 'rgba(239, 68, 68, 0.15)';
                        estadoColorText = '#f87171';
                        estadoTexto = 'Vencido (Reagendar)';
                        alertasRojasCount++;
                    } else {
                        estadoColorBg = 'rgba(251, 191, 36, 0.15)';
                        estadoColorText = '#fbbf24';
                        estadoTexto = 'En Proceso';
                        alertasAmarillasCount++;
                    }

                    tablaAgenda += `<tr>
              <td style="padding: 12px;"><strong>${ag.nombre_empresa}</strong></td>
              <td style="padding: 12px;"><span class="badge-tech" style="background:rgba(56,189,248,0.1); color:#38bdf8;"><i class="fa-solid fa-user"></i> ${ag.vendedor || 'General'}</span></td>
              <td style="padding: 12px;">${ag.tipo}</td>
              <td style="padding: 12px;">${ag.fecha_programada}</td>
              <td style="padding: 12px;">${ag.comentario}</td>
              <td style="padding: 12px;"><span style="background:${estadoColorBg}; color:${estadoColorText}; padding:4px 10px; border-radius:6px; font-weight:600; font-size:11px; display:inline-block;">${estadoTexto}</span></td>
              <td style="padding: 12px;">
                <a href="/concluir-agenda/${ag.id}" class="btn-primario" style="padding:4px 8px; font-size:10px; background:linear-gradient(135deg, #10b981, #059669); text-decoration:none; display:inline-block; margin-bottom:2px;"><i class="fa-solid fa-check"></i></a>
                <a href="/reagendar-agenda/${ag.id}" class="btn-primario" style="padding:4px 8px; font-size:10px; background:linear-gradient(135deg, #d97706, #b45309); text-decoration:none; display:inline-block; margin-bottom:2px;"><i class="fa-solid fa-rotate"></i></a>
                <a href="/eliminar-agenda/${ag.id}" onclick="return confirm('¿Eliminar este registro del cronograma?')" class="btn-peligro" style="padding:4px 8px; font-size:10px;"><i class="fa-solid fa-trash"></i></a>
              </td>
            </tr>`;
                });

                let listaNotasHtml = '';
                notasMes.forEach(n => {
                    listaNotasHtml += `<div style="background:rgba(30,41,59,0.5); border:1px solid rgba(56,189,248,0.2); padding:10px 14px; border-radius:8px; margin-bottom:10px; display:flex; justify-content:space-between; align-items:center;">
                        <div>
                            <span style="font-size:11px; color:#38bdf8; font-weight:600;"><i class="fa-solid fa-user"></i> ${n.vendedor}</span>
                            <p style="margin:4px 0 0 0; font-size:13px; color:#f8fafc;">${n.nota}</p>
                        </div>
                        <a href="/eliminar-nota/${n.id}" onclick="return confirm('¿Borrar esta nota?')" style="color:#f87171; text-decoration:none; font-size:12px;"><i class="fa-solid fa-trash"></i></a>
                    </div>`;
                });

                let optsClientes = clientes.map(c => `<option value="${c.id}">${c.nombre_empresa}</option>`).join('');

                res.send(`
            <html>
              <body style="max-width: 1450px; margin: auto;">
                ${generarEncabezado('Directorio de Clientes & CRM Avanzado', req)}

                ${(alertasRojasCount > 0 || alertasAmarillasCount > 0) ? `
                  <div style="background: ${alertasRojasCount > 0 ? 'rgba(239, 68, 68, 0.15)' : 'rgba(251, 191, 36, 0.15)'}; border-left: 5px solid ${alertasRojasCount > 0 ? '#ef4444' : '#fbbf24'}; padding: 16px; border-radius: 8px; margin-bottom: 24px;">
                    <h4 style="margin: 0; color: ${alertasRojasCount > 0 ? '#f87171' : '#fbbf24'};"><i class="fa-solid fa-bell"></i> Alerta del Sistema de Citas y Pendientes</h4>
                    <p style="margin: 6px 0 0 0; font-size: 13px; color: #cbd5e1;">
                      ${alertasRojasCount > 0 ? `<b>¡Atención!</b> Tienes <b>${alertasRojasCount}</b> cita(s) o pendiente(s) vencidas en color <b>ROJO</b> que requieren reagendación.` : ''}
                      ${alertasAmarillasCount > 0 ? ` Tienes <b>${alertasAmarillasCount}</b> pendiente(s) en proceso.` : ''}
                    </p>
                  </div>
                ` : ''}
                
                <div style="display: flex; gap: 24px; margin-bottom: 24px; flex-wrap: wrap;">
                  <div class="contenedor-tarjeta" style="width: 380px; height: fit-content;">
                    <h3 style="margin-top:0;"><i class="fa-solid fa-user-plus" style="color:#38bdf8; margin-right:8px;"></i>Registro de Cliente</h3>
                    <form action="/guardar-cliente" method="POST" enctype="multipart/form-data">
                      <input type="text" name="numero_cliente" placeholder="No. Cliente (Ej. CLI-001)" class="input-industrial">
                      <input type="text" name="nombre_empresa" required placeholder="Nombre de la Empresa" class="input-industrial">
                      <input type="text" name="rfc" placeholder="RFC Fiscal" class="input-industrial">
                      <input type="text" name="contacto_nombre" placeholder="Contacto Principal" class="input-industrial">
                      <input type="text" name="telefono" placeholder="Teléfono / WhatsApp" class="input-industrial">
                      <input type="email" name="email" placeholder="Correo Electrónico" class="input-industrial">
                      <input type="text" name="ubicacion" placeholder="Ubicación / Dirección" class="input-industrial">
                      
                      <label style="font-size:11px; font-weight:600; color:#38bdf8; text-transform:uppercase;">¿Cuenta con Crédito?:</label>
                      <select name="tiene_credito" class="input-industrial" style="margin-top:5px;">
                        <option value="No">No (Contado / Anticipo)</option>
                        <option value="Sí">Sí Cuenta con Crédito</option>
                      </select>

                      <label style="font-size:11px; font-weight:600; color:#94a3b8; text-transform:uppercase;">Monto de Crédito Autorizado ($):</label>
                      <input type="number" step="0.01" name="monto_credito" value="0" placeholder="Ej. 50000" class="input-industrial" style="margin-top:5px;">
                      <small style="color:#a78bfa; display:block; margin-bottom:12px;"><i class="fa-solid fa-shield-halved"></i> Otorgado por Contabilidad o Ventas.</small>

                      <label style="font-size:11px; font-weight:600; color:#38bdf8; text-transform:uppercase;">Constancia Fiscal (CSF):</label>
                      <input type="file" name="archivo_csf" accept=".pdf,.jpg,.png" required class="input-industrial" style="font-size:11px; margin-bottom:15px; padding:6px;">
                      
                      <input type="hidden" name="vendedor_asignado" value="${u.username}">
                      <button type="submit" class="btn-primario" style="width: 100%; justify-content: center;"><i class="fa-solid fa-cloud-arrow-up"></i> Guardar Cliente</button>
                    </form>
                  </div>
                  <div class="contenedor-tarjeta" style="flex-grow: 1; overflow-x: auto;">
                    <h3 style="margin-top:0;"><i class="fa-solid fa-address-book" style="color:#38bdf8; margin-right:8px;"></i>Cartera de Clientes & Créditos</h3>
                    <table>
                      <tr>
                        <th>No.</th>
                        <th>Empresa & Vendedor</th>
                        <th>Crédito</th>
                        <th>WhatsApp</th>
                        <th>CSF Fiscal</th>
                        <th>Acciones</th>
                      </tr>
                      ${tablaClientes}
                    </table>
                  </div>
                </div>

                <div class="contenedor-tarjeta" style="margin-bottom: 24px;">
                  <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:15px; flex-wrap:wrap; gap:10px;">
                    <h3 style="margin:0;"><i class="fa-solid fa-calendar-days" style="color:#38bdf8; margin-right:8px;"></i>Calendario de Notas por Mes</h3>
                    <form action="/" method="GET" style="display:flex; gap:10px; align-items:center;">
                      <label style="font-size:11px; font-weight:600; color:#94a3b8; text-transform:uppercase;">Seleccionar Mes:</label>
                      <input type="month" name="mes" value="${mesSeleccionado}" class="input-industrial" style="margin-bottom:0; width:160px; padding:6px 10px;">
                      <button type="submit" class="btn-primario" style="padding:6px 14px;"><i class="fa-solid fa-filter"></i> Ver Mes</button>
                    </form>
                  </div>

                  <div style="display:flex; gap:24px; flex-wrap:wrap;">
                    <div style="width: 350px;">
                      <form action="/guardar-nota-mes" method="POST">
                        <input type="hidden" name="mes" value="${mesSeleccionado}">
                        <label style="font-size:11px; font-weight:600; color:#38bdf8; text-transform:uppercase;">Añadir Nota para (${mesSeleccionado}):</label>
                        <textarea name="nota" required placeholder="Escribe aquí la nota o pendiente general del mes..." class="input-industrial" style="height:90px; margin-top:5px; margin-bottom:12px;"></textarea>
                        <button type="submit" class="btn-primario" style="width:100%; background:linear-gradient(135deg, #0ea5e9, #0284c7); justify-content:center;"><i class="fa-solid fa-plus"></i> Guardar Nota Mensual</button>
                      </form>
                    </div>
                    <div style="flex-grow:1; max-height:220px; overflow-y:auto; padding-right:5px;">
                      ${listaNotasHtml || '<p style="color:#94a3b8; font-size:13px; text-align:center; padding-top:40px;">No hay notas registradas para este mes.</p>'}
                    </div>
                  </div>
                </div>

                <div style="display: flex; gap: 24px; flex-wrap: wrap;">
                  <div class="contenedor-tarjeta" style="width: 380px; height: fit-content;">
                    <h3 style="margin-top:0;"><i class="fa-solid fa-calendar-check" style="color:#38bdf8; margin-right:8px;"></i>Agendar Cita / Visita</h3>
                    <form action="/guardar-agenda" method="POST">
                      <label style="font-size:11px; font-weight:600; color:#94a3b8; text-transform:uppercase;">Cliente:</label>
                      <select name="cliente_id" required class="input-industrial" style="margin-top:5px;">
                        <option value="">-- Seleccionar --</option>
                        ${optsClientes}
                      </select>
                      <label style="font-size:11px; font-weight:600; color:#94a3b8; text-transform:uppercase;">Tipo de Actividad:</label>
                      <select name="tipo" class="input-industrial" style="margin-top:5px;">
                        <option value="Visita Presencial">🤝 Visita Presencial</option>
                        <option value="Llamada Telefónica">📞 Llamada Telefónica</option>
                        <option value="Correo Electrónico">✉️ Correo Electrónico</option>
                      </select>
                      <label style="font-size:11px; font-weight:600; color:#94a3b8; text-transform:uppercase;">Fecha y Hora:</label>
                      <input type="datetime-local" name="fecha_programada" required class="input-industrial" style="margin-top:5px;">
                      <input type="text" name="comentario" required placeholder="Notas de la cita..." class="input-industrial">
                      <button type="submit" class="btn-primario" style="width: 100%; background:linear-gradient(135deg, #0ea5e9, #0284c7); justify-content: center;"><i class="fa-solid fa-calendar-plus"></i> Agendar</button>
                    </form>
                  </div>
                  <div class="contenedor-tarjeta" style="flex-grow: 1; overflow-x: auto;">
                    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px;">
                      <h3 style="margin:0;"><i class="fa-solid fa-clock-rotate-left" style="color:#38bdf8; margin-right:8px;"></i>Cronograma y Pendientes con Vendedor</h3>
                    </div>
                    <table>
                      <tr>
                        <th>Cliente</th>
                        <th>Vendedor</th>
                        <th>Tipo</th>
                        <th>Fecha & Hora</th>
                        <th>Detalle</th>
                        <th>Estado</th>
                        <th>Acción</th>
                      </tr>
                      ${tablaAgenda}
                    </table>
                  </div>
                </div>
              </body>
            </html>
          `);
            });
        });
    });
});

app.get('/eliminar-cliente/:id', verificarSesion, (req, res) => {
    db.run("DELETE FROM clientes WHERE id = ?", [req.params.id], () => res.redirect('/'));
});

app.get('/eliminar-agenda/:id', verificarSesion, (req, res) => {
    db.run("DELETE FROM crm_agenda WHERE id = ?", [req.params.id], () => res.redirect('/'));
});

app.post('/guardar-nota-mes', verificarSesion, (req, res) => {
    const u = req.session.usuario;
    const { mes, nota } = req.body;
    const anio = mes ? mes.slice(0, 4) : '2026';
    db.run("INSERT INTO calendario_notas (vendedor, mes, anio, nota) VALUES (?, ?, ?, ?)", [u.username, mes, anio, nota], () => {
        res.redirect(`/?mes=${mes}`);
    });
});

app.get('/eliminar-nota/:id', verificarSesion, (req, res) => {
    db.get("SELECT mes FROM calendario_notas WHERE id = ?", [req.params.id], (err, row) => {
        let mesRef = row ? row.mes : '';
        db.run("DELETE FROM calendario_notas WHERE id = ?", [req.params.id], () => {
            res.redirect(`/?mes=${mesRef}`);
        });
    });
});

app.get('/reagendar-agenda/:id', verificarSesion, (req, res) => {
    db.get("SELECT ag.*, cl.nombre_empresa FROM crm_agenda ag JOIN clientes cl ON ag.cliente_id = cl.id WHERE ag.id = ?", [req.params.id], (err, ag) => {
        if (!ag) return res.send("Registro no encontrado");
        res.send(`
      <html>
        <body style="padding-top: 30px; max-width: 500px; margin: auto;">
          ${generarEncabezado('Reagendar Cita', req)}
          <div class="contenedor-tarjeta">
            <h3 style="color:#fbbf24; margin-top:0;"><i class="fa-solid fa-rotate"></i> Reagendar: ${ag.nombre_empresa}</h3>
            <p style="font-size: 13px; color: #94a3b8;">Anterior: ${ag.tipo} (${ag.fecha_programada})</p>
            <form action="/guardar-reagendar-agenda/${ag.id}" method="POST">
              <label style="font-size:11px; font-weight:600; color:#94a3b8; text-transform: uppercase;">Nueva Fecha y Hora:</label>
              <input type="datetime-local" name="fecha_programada" required class="input-industrial" style="margin-top:5px;">
              <label style="font-size:11px; font-weight:600; color:#94a3b8; text-transform: uppercase;">Motivo / Notas:</label>
              <input type="text" name="comentario" value="${ag.comentario}" required class="input-industrial" style="margin-top:5px; margin-bottom: 20px;">
              <button type="submit" class="btn-primario" style="width:100%; background:linear-gradient(135deg, #d97706, #b45309); justify-content: center;"><i class="fa-solid fa-check"></i> Guardar Nueva Fecha</button>
            </form>
            <br><a href="/" style="color:#38bdf8; font-weight:600; text-decoration:none; display:block; text-align:center; font-size:12px;"><i class="fa-solid fa-arrow-left"></i> Cancelar</a>
          </div>
        </body>
      </html>
    `);
    });
});

app.post('/guardar-reagendar-agenda/:id', verificarSesion, (req, res) => {
    const { fecha_programada, comentario } = req.body;
    db.run("UPDATE crm_agenda SET fecha_programada = ?, comentario = ?, estado = 'Pendiente' WHERE id = ?", [fecha_programada, comentario, req.params.id], () => {
        res.redirect('/');
    });
});

app.post('/guardar-agenda', verificarSesion, (req, res) => {
    const u = req.session.usuario;
    const { cliente_id, tipo, fecha_programada, comentario } = req.body;
    db.run("INSERT INTO crm_agenda (cliente_id, vendedor, tipo, fecha_programada, comentario) VALUES (?, ?, ?, ?, ?)",
        [cliente_id, u.username, tipo, fecha_programada, comentario], () => res.redirect('/'));
});

app.get('/concluir-agenda/:id', verificarSesion, (req, res) => {
    db.get("SELECT ag.*, cl.nombre_empresa FROM crm_agenda ag JOIN clientes cl ON ag.cliente_id = cl.id WHERE ag.id = ?", [req.params.id], (err, ag) => {
        if (!ag) return res.send("Registro no encontrado");
        res.send(`
      <html>
        <body style="padding-top: 30px; max-width: 600px; margin: auto;">
          ${generarEncabezado('Conclusión de Cita', req)}
          <div class="contenedor-tarjeta">
            <h2 style="color:#34d399; margin-top:0;"><i class="fa-solid fa-circle-check"></i> Concluir Visita</h2>
            <p style="font-size: 13px;"><b>Cliente:</b> ${ag.nombre_empresa}<br><b>Actividad:</b> ${ag.tipo} (${ag.fecha_programada})</p>
            <form action="/guardar-concluir-agenda/${ag.id}" method="POST">
              <label style="font-size:11px; font-weight:600; color:#94a3b8; text-transform: uppercase;">Comentarios / Resultados:</label>
              <textarea name="resultado_visita" required placeholder="Acuerdos y requerimientos de la visita..." class="input-industrial" style="height:100px; margin-top:5px; margin-bottom:20px;"></textarea>
              <button type="submit" class="btn-primario" style="width:100%; background:linear-gradient(135deg, #10b981, #059669); justify-content: center;"><i class="fa-solid fa-file-invoice"></i> Guardar y Proceder a Cotizar</button>
            </form>
          </div>
        </body>
      </html>
    `);
    });
});

app.post('/guardar-concluir-agenda/:id', verificarSesion, (req, res) => {
    const { resultado_visita } = req.body;
    db.get("SELECT cliente_id FROM crm_agenda WHERE id = ?", [req.params.id], (err, row) => {
        db.run("UPDATE crm_agenda SET estado = 'Completado', resultado_visita = ? WHERE id = ?", [resultado_visita, req.params.id], () => {
            if (row) {
                res.redirect(`/cotizaciones?cliente_id=${row.cliente_id}`);
            } else {
                res.redirect('/');
            }
        });
    });
});

// SOPORTE Y TICKETS
app.get('/soporte', verificarSesion, (req, res) => {
    const u = req.session.usuario;
    const esAdmin = u.rol === 'Administrador' || u.departamento === 'Administración' || u.departamento === 'Control Administrativo' || u.departamento === 'Contabilidad';

    let queryClientes = esAdmin ? "SELECT id, nombre_empresa FROM clientes ORDER BY nombre_empresa" : "SELECT id, nombre_empresa FROM clientes WHERE vendedor_asignado = ? OR vendedor_asignado = ? ORDER BY nombre_empresa";
    let paramsClientes = esAdmin ? [] : [u.username, u.departamento];

    db.all(queryClientes, paramsClientes, (err, clientes) => {
        let queryTickets = esAdmin ?
            "SELECT t.*, cl.nombre_empresa FROM soporte_tickets t JOIN clientes cl ON t.cliente_id = cl.id ORDER BY t.id DESC" :
            "SELECT t.*, cl.nombre_empresa FROM soporte_tickets t JOIN clientes cl ON t.cliente_id = cl.id WHERE t.vendedor = ? OR t.vendedor = ? ORDER BY t.id DESC";
        let paramsTickets = esAdmin ? [] : [u.username, u.departamento];

        db.all(queryTickets, paramsTickets, (err, tickets) => {
            let optsClientes = clientes.map(c => `<option value="${c.id}">${c.nombre_empresa}</option>`).join('');
            let tablaTickets = '';

            tickets.forEach(tk => {
                let colorEstado = tk.estado === 'Resuelto' ? '#34d399' : tk.estado === 'En Proceso' ? '#fbbf24' : '#38bdf8';
                let colorPrioridad = tk.prioridad === 'Alta' ? '#f87171' : tk.prioridad === 'Urgente' ? '#a78bfa' : '#94a3b8';

                tablaTickets += `<tr>
                    <td style="padding: 12px;"><strong>TCK-${tk.id}</strong><br><small style="color:#94a3b8;">${tk.vendedor}</small></td>
                    <td style="padding: 12px;"><strong>${tk.nombre_empresa}</strong></td>
                    <td style="padding: 12px;"><span class="badge-tech" style="background:rgba(56,189,248,0.1); color:#38bdf8;">${tk.tipo_solicitud}</span></td>
                    <td style="padding: 12px;"><strong>${tk.asunto}</strong><br><small style="color:#94a3b8;">${tk.descripcion}</small></td>
                    <td style="padding: 12px;"><span style="color:${colorPrioridad}; font-weight:600;">${tk.prioridad || 'Media'}</span></td>
                    <td style="padding: 12px;"><span style="color:${colorEstado}; font-weight:600;">${tk.estado}</span></td>
                    <td style="padding: 12px;">
                        <a href="/actualizar-ticket/${tk.id}" class="btn-primario" style="padding:4px 8px; font-size:10px; background:linear-gradient(135deg, #0ea5e9, #0284c7); text-decoration:none;"><i class="fa-solid fa-pen"></i></a>
                    </td>
                </tr>`;
            });

            res.send(`
                <html>
                    <body style="max-width: 1450px; margin: auto;">
                        ${generarEncabezado('Servicio y Soporte al Cliente', req)}
                        <div style="display: flex; gap: 24px; flex-wrap: wrap;">
                            <div class="contenedor-tarjeta" style="width: 380px; height: fit-content;">
                                <h3 style="margin-top:0;"><i class="fa-solid fa-ticket" style="color:#38bdf8; margin-right:8px;"></i>Nuevo Ticket</h3>
                                <form action="/guardar-ticket" method="POST">
                                    <label style="font-size:11px; font-weight:600; color:#94a3b8; text-transform:uppercase;">Cliente:</label>
                                    <select name="cliente_id" required class="input-industrial" style="margin-top:5px;">
                                        <option value="">-- Seleccionar --</option>
                                        ${optsClientes}
                                    </select>

                                    <label style="font-size:11px; font-weight:600; color:#94a3b8; text-transform:uppercase;">Tipo:</label>
                                    <select name="tipo_solicitud" class="input-industrial" style="margin-top:5px;">
                                        <option value="Soporte Técnico">🛠️ Soporte Técnico</option>
                                        <option value="Queja o Reclamo">⚠️ Queja o Reclamo</option>
                                        <option value="Duda Operativa">❓ Duda Operativa</option>
                                        <option value="Solicitud de Asistencia">📋 Asistencia</option>
                                    </select>

                                    <label style="font-size:11px; font-weight:600; color:#94a3b8; text-transform:uppercase;">Prioridad:</label>
                                    <select name="prioridad" class="input-industrial" style="margin-top:5px;">
                                        <option value="Baja">Baja</option>
                                        <option value="Media" selected>Media</option>
                                        <option value="Alta">Alta</option>
                                        <option value="Urgente">Urgente</option>
                                    </select>

                                    <input type="text" name="asunto" required placeholder="Asunto breve..." class="input-industrial">
                                    <textarea name="descripcion" required placeholder="Detalles de la incidencia..." class="input-industrial" style="height:80px; margin-bottom:15px;"></textarea>

                                    <button type="submit" class="btn-primario" style="width: 100%; background:linear-gradient(135deg, #10b981, #059669); justify-content: center;"><i class="fa-solid fa-floppy-disk"></i> Registrar Ticket</button>
                                </form>
                            </div>
                            <div class="contenedor-tarjeta" style="flex-grow: 1; overflow-x: auto;">
                                <h3 style="margin-top:0;"><i class="fa-solid fa-headset" style="color:#38bdf8; margin-right:8px;"></i>Historial de Tickets de Soporte</h3>
                                <table>
                                    <tr>
                                        <th>Folio</th>
                                        <th>Empresa</th>
                                        <th>Tipo</th>
                                        <th>Asunto & Descripción</th>
                                        <th>Prioridad</th>
                                        <th>Estatus</th>
                                        <th>Acción</th>
                                    </tr>
                                    ${tablaTickets}
                                </table>
                            </div>
                        </div>
                    </body>
                </html>
            `);
        });
    });
});

app.post('/guardar-ticket', verificarSesion, (req, res) => {
    const u = req.session.usuario;
    const { cliente_id, tipo_solicitud, prioridad, asunto, descripcion } = req.body;
    db.run("INSERT INTO soporte_tickets (cliente_id, vendedor, tipo_solicitud, prioridad, asunto, descripcion) VALUES (?, ?, ?, ?, ?, ?)",
        [cliente_id, u.username, tipo_solicitud, prioridad, asunto, descripcion], () => {
            res.redirect('/soporte');
        });
});

app.get('/actualizar-ticket/:id', verificarSesion, (req, res) => {
    db.get("SELECT t.*, cl.nombre_empresa FROM soporte_tickets t JOIN clientes cl ON t.cliente_id = cl.id WHERE t.id = ?", [req.params.id], (err, tk) => {
        if (!tk) return res.send("Ticket no encontrado");
        res.send(`
            <html>
                <body style="padding-top: 30px; max-width: 500px; margin: auto;">
                    ${generarEncabezado('Actualizar Estatus de Ticket', req)}
                    <div class="contenedor-tarjeta">
                        <h3 style="color:#38bdf8; margin-top:0;"><i class="fa-solid fa-ticket"></i> TCK-${tk.id} — ${tk.nombre_empresa}</h3>
                        <p style="font-size:13px; color:#94a3b8;"><b>Asunto:</b> ${tk.asunto}<br><b>Descripción:</b> ${tk.descripcion}</p>
                        <form action="/guardar-estatus-ticket/${tk.id}" method="POST">
                            <label style="font-size:11px; font-weight:600; color:#94a3b8; text-transform:uppercase;">Estatus:</label>
                            <select name="estado" class="input-industrial" style="margin-top:5px; margin-bottom:20px;">
                                <option value="Abierto" ${tk.estado === 'Abierto' ? 'selected' : ''}>Abierto</option>
                                <option value="En Proceso" ${tk.estado === 'En Proceso' ? 'selected' : ''}>En Proceso</option>
                                <option value="Resuelto" ${tk.estado === 'Resuelto' ? 'selected' : ''}>Resuelto ✅</option>
                            </select>
                            <button type="submit" class="btn-primario" style="width:100%; background:linear-gradient(135deg, #10b981, #059669); justify-content: center;"><i class="fa-solid fa-check"></i> Actualizar Estatus</button>
                        </form>
                        <br><a href="/soporte" style="color:#38bdf8; font-weight:600; text-decoration:none; display:block; text-align:center; font-size:12px;"><i class="fa-solid fa-arrow-left"></i> Regresar</a>
                    </div>
                </body>
            </html>
        `);
    });
});

app.post('/guardar-estatus-ticket/:id', verificarSesion, (req, res) => {
    const { estado } = req.body;
    db.run("UPDATE soporte_tickets SET estado = ? WHERE id = ?", [estado, req.params.id], () => {
        res.redirect('/soporte');
    });
});

const uploadCsf = upload.single('archivo_csf');
app.post('/guardar-cliente', verificarSesion, uploadCsf, (req, res) => {
    const { numero_cliente, nombre_empresa, rfc, contacto_nombre, telefono, email, ubicacion, vendedor_asignado, dias_credito, tiene_credito, monto_credito } = req.body;
    let archivo_csf = null;
    if (req.file) {
        const ext = path.extname(req.file.originalname);
        archivo_csf = req.file.filename + ext;
        fs.renameSync(req.file.path, path.join(__dirname, 'uploads/csf', archivo_csf));
    }
    db.run("INSERT INTO clientes (numero_cliente, nombre_empresa, rfc, contacto_nombre, telefono, email, ubicacion, vendedor_asignado, archivo_csf, tiene_credito, monto_credito, dias_credito) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
        [numero_cliente, nombre_empresa, rfc, contacto_nombre, telefono, email, ubicacion, vendedor_asignado, archivo_csf, tiene_credito || 'No', monto_credito || 0, dias_credito || 30], () => res.redirect('/'));
});

app.get('/cliente-detalle/:id', verificarSesion, (req, res) => {
    db.get("SELECT * FROM clientes WHERE id = ?", [req.params.id], (err, c) => {
        if (!c) return res.send("Cliente no encontrado");
        let infoCredito = c.tiene_credito === 'Sí' ? `Sí ($${formatoDinero(c.monto_credito)})` : 'No cuenta con crédito';
        res.send(`
      <html>
        <body style="padding-top: 30px; max-width: 600px; margin: auto;">
          ${generarEncabezado('Detalle de Cliente', req)}
          <div class="contenedor-tarjeta">
            <h2 style="color:#38bdf8; margin-top:0;"><i class="fa-solid fa-building"></i> ${c.nombre_empresa}</h2>
            <p style="font-size:13px; line-height: 1.6; color: #cbd5e1;">
              <b>No. Cliente:</b> ${c.numero_cliente || 'N/A'}<br>
              <b>RFC:</b> ${c.rfc || 'N/A'}<br>
              <b>Crédito Autorizado:</b> <span style="color:#a78bfa; font-weight:600;">${infoCredito}</span><br>
              <b>Contacto Principal:</b> ${c.contacto_nombre || 'N/A'}<br>
              <b>Teléfono (WhatsApp):</b> ${c.telefono || 'N/A'} ${c.telefono ? `<a href="https://wa.me/${c.telefono.replace(/\D/g, '')}" target="_blank" style="color:#4ade80;"><i class="fa-brands fa-whatsapp"></i> Chat</a>` : ''}<br>
              <b>Correo:</b> ${c.email || 'N/A'}<br>
              <b>Ubicación:</b> ${c.ubicacion || 'N/A'}<br>
              <b>Vendedor Asignado:</b> ${c.vendedor_asignado || 'General'}<br>
              <b>Constancia Fiscal (CSF):</b> ${c.archivo_csf ? `<a href="/csf/${c.archivo_csf}" target="_blank" style="color:#38bdf8;"><i class="fa-solid fa-file-pdf"></i> Ver Archivo</a>` : 'No adjunta'}
            </p>
            <br><a href="/" class="btn-primario" style="text-decoration:none;"><i class="fa-solid fa-arrow-left"></i> Volver al Directorio</a>
          </div>
        </body>
      </html>
    `);
    });
});

app.get('/editar-cliente/:id', verificarSesion, (req, res) => {
    db.get("SELECT * FROM clientes WHERE id = ?", [req.params.id], (err, c) => {
        if (!c) return res.send("Cliente no encontrado");
        res.send(`
      <html>
        <body style="padding-top: 30px; max-width: 600px; margin: auto;">
          ${generarEncabezado('Modificar Cliente y Crédito', req)}
          <div class="contenedor-tarjeta">
            <h2 style="color:#38bdf8; margin-top:0;"><i class="fa-solid fa-pen-to-square"></i> Modificar: ${c.nombre_empresa}</h2>
            <form action="/actualizar-cliente/${c.id}" method="POST">
              <label style="font-size:11px; font-weight:600; color:#94a3b8;">Nombre Empresa:</label>
              <input type="text" name="nombre_empresa" value="${c.nombre_empresa}" required class="input-industrial">
              <label style="font-size:11px; font-weight:600; color:#94a3b8;">RFC:</label>
              <input type="text" name="rfc" value="${c.rfc || ''}" class="input-industrial">
              
              <label style="font-size:11px; font-weight:600; color:#38bdf8;">¿Cuenta con Crédito?:</label>
              <select name="tiene_credito" class="input-industrial" style="margin-top:5px;">
                <option value="No" ${c.tiene_credito === 'No' ? 'selected' : ''}>No</option>
                <option value="Sí" ${c.tiene_credito === 'Sí' ? 'selected' : ''}>Sí Cuenta con Crédito</option>
              </select>

              <label style="font-size:11px; font-weight:600; color:#94a3b8;">Monto de Crédito Autorizado ($):</label>
              <input type="number" step="0.01" name="monto_credito" value="${c.monto_credito || 0}" class="input-industrial">
              
              <label style="font-size:11px; font-weight:600; color:#94a3b8;">Contacto:</label>
              <input type="text" name="contacto_nombre" value="${c.contacto_nombre || ''}" class="input-industrial">
              <label style="font-size:11px; font-weight:600; color:#94a3b8;">Teléfono:</label>
              <input type="text" name="telefono" value="${c.telefono || ''}" class="input-industrial">
              <label style="font-size:11px; font-weight:600; color:#94a3b8;">Correo:</label>
              <input type="email" name="email" value="${c.email || ''}" class="input-industrial">
              <label style="font-size:11px; font-weight:600; color:#94a3b8;">Ubicación:</label>
              <input type="text" name="ubicacion" value="${c.ubicacion || ''}" class="input-industrial" style="margin-bottom:20px;">
              <button type="submit" class="btn-primario" style="width:100%; background:linear-gradient(135deg, #10b981, #059669); justify-content: center;"><i class="fa-solid fa-check"></i> Guardar Cambios</button>
            </form>
            <br><a href="/" style="color:#38bdf8; font-weight:600; text-decoration:none; font-size:12px;"><i class="fa-solid fa-arrow-left"></i> Cancelar</a>
          </div>
        </body>
      </html>
    `);
    });
});

app.post('/actualizar-cliente/:id', verificarSesion, (req, res) => {
    const { nombre_empresa, rfc, tiene_credito, monto_credito, contacto_nombre, telefono, email, ubicacion } = req.body;
    db.run("UPDATE clientes SET nombre_empresa = ?, rfc = ?, tiene_credito = ?, monto_credito = ?, contacto_nombre = ?, telefono = ?, email = ?, ubicacion = ? WHERE id = ?",
        [nombre_empresa, rfc, tiene_credito, monto_credito || 0, contacto_nombre, telefono, email, ubicacion, req.params.id], () => res.redirect('/'));
});

// CHAT INTERDEPARTAMENTAL
app.get('/chat', verificarSesion, (req, res) => {
    const u = req.session.usuario;
    db.all("SELECT * FROM chat_interno ORDER BY id DESC LIMIT 50", [], (err, mensajes) => {
        let listaMsg = '';
        mensajes.reverse().forEach(m => {
            listaMsg += `<div style="background:rgba(30, 41, 59, 0.4); border: 1px solid rgba(56,189,248,0.1); padding:12px; border-radius:8px; margin-bottom:10px; font-size:13px;">
        <span style="color:#38bdf8; font-weight:600;"><i class="fa-solid fa-user-circle"></i> ${m.remitente} (${m.departamento})</span> <span style="font-size:11px; color:#94a3b8;">- ${m.fecha}</span><br>
        <div style="margin-top:6px; color:#f8fafc;">${m.mensaje}</div>
      </div>`;
        });

        res.send(`
      <html>
        <body style="max-width: 1000px; margin: auto;">
          ${generarEncabezado('Chat Interdepartamental', req)}
          <div class="contenedor-tarjeta">
            <h3 style="margin-top:0;"><i class="fa-solid fa-comments" style="color:#38bdf8; margin-right:8px;"></i>Coordinación en Tiempo Real</h3>
            <div style="height:380px; overflow-y:auto; border:1px solid rgba(56, 189, 248, 0.2); padding:15px; border-radius:8px; background:rgba(3, 7, 18, 0.5); margin-bottom:15px;">
              ${listaMsg || '<p style="color:#94a3b8; text-align:center; padding-top:140px;">No hay mensajes recientes en el canal.</p>'}
            </div>
            <form action="/enviar-mensaje" method="POST" style="display:flex; gap:12px;">
              <input type="text" name="mensaje" required placeholder="Escribe un mensaje para los departamentos..." class="input-industrial" style="margin-bottom:0; flex-grow:1;">
              <button type="submit" class="btn-primario" style="height: 44px;"><i class="fa-solid fa-paper-plane"></i> Enviar</button>
            </form>
          </div>
        </body>
      </html>
    `);
    });
});

app.post('/enviar-mensaje', verificarSesion, (req, res) => {
    const u = req.session.usuario;
    const { mensaje } = req.body;
    db.run("INSERT INTO chat_interno (departamento, remitente, mensaje) VALUES (?, ?, ?)", [u.departamento, u.username, mensaje], () => {
        res.redirect('/chat');
    });
});

// CATÁLOGOS & STOCK
app.get('/productos', verificarSesion, (req, res) => {
    const u = req.session.usuario;
    const esAdmin = u.rol === 'Administrador' || u.departamento === 'Administración' || u.departamento === 'Control Administrativo';
    const categoriaFiltro = req.query.cat || '';
    const busqueda = req.query.busqueda || '';

    let query = "SELECT * FROM productos_servicios WHERE 1=1";
    let params = [];
    if (categoriaFiltro) { query += " AND categoria = ?"; params.push(categoriaFiltro); }
    if (busqueda) { query += " AND (codigo LIKE ? OR nombre LIKE ? OR descripcion LIKE ?)"; params.push(`%${busqueda}%`, `%${busqueda}%`, `%${busqueda}%`); }
    query += " ORDER BY categoria, id DESC";

    db.all(query, params, (err, filas) => {
        let tabla = '';
        filas.forEach(p => {
            let fotoHtml = p.foto ? `<img src="/productos_img/${p.foto}" style="width:40px; height:40px; object-fit:cover; border-radius:6px; border:1px solid rgba(56,189,248,0.2);">` : '<span style="color:#64748b; font-size:11px;">S/Foto</span>';

            tabla += `<tr>
        <td style="padding: 12px; text-align:center;">${fotoHtml}</td>
        <td style="padding: 12px; font-weight: 600; color: #38bdf8;">${p.codigo || 'S/N'}</td>
        <td style="padding: 12px;"><span class="badge-tech" style="background: rgba(56, 189, 248, 0.1); color: #38bdf8;">${p.categoria}</span></td>
        <td style="padding: 12px;"><strong>${p.nombre}</strong><br><small style="color:#94a3b8;">${p.descripcion || ''}</small></td>
        <td style="padding: 12px; color: #cbd5e1;">${p.dimensiones || 'N/A'}</td>
        <td style="padding: 12px; color: #34d399; font-weight: 600;">$${formatoDinero(p.precio_unitario)}</td>
        <td style="padding: 12px; color: #cbd5e1;">Stock: ${p.stock || 0}<br><small style="color:#a78bfa;">Entrega: ${p.tiempo_entrega || 'Inmediato'}</small></td>
        <td style="padding: 12px;">
          ${esAdmin ? `
            <a href="/editar-producto/${p.id}" class="btn-primario" style="padding: 4px 8px; font-size:10px; background:linear-gradient(135deg, #d97706, #b45309); text-decoration:none; display:inline-block; margin-bottom:2px;"><i class="fa-solid fa-pen"></i></a>
            <a href="/eliminar-producto/${p.id}" onclick="return confirm('¿Eliminar ítem del inventario?')" class="btn-peligro" style="padding: 4px 8px;"><i class="fa-solid fa-trash"></i></a>
          ` : '<span style="color: #64748b; font-size: 11px;">Lectura</span>'}
        </td>
      </tr>`;
        });

        let navCategorias = CATEGORIAS_VALIDAS.map(cat =>
            `<a href="/productos?cat=${encodeURIComponent(cat)}" style="text-decoration: none; color: ${cat === categoriaFiltro ? '#ffffff' : '#94a3b8'}; font-size: 11px; background: ${cat === categoriaFiltro ? '#0284c7' : 'rgba(30,41,59,0.5)'}; padding: 6px 10px; border-radius: 6px; margin-right: 5px; margin-bottom: 6px; display: inline-flex; align-items: center; gap: 5px; border: 1px solid rgba(56,189,248,0.15);"><i class="fa-solid fa-folder"></i> ${cat}</a>`
        ).join('');

        let optionsSelectCat = CATEGORIAS_VALIDAS.map(cat => `<option value="${cat}">${cat}</option>`).join('');

        res.send(`
      <html>
        <body style="max-width: 1450px; margin: auto;">
          ${generarEncabezado('Catálogos Especializados & Inventario Industrial', req)}
          
          <div class="contenedor-tarjeta">
            <div style="margin-bottom: 15px; display: flex; flex-wrap: wrap; gap: 6px; align-items: center; justify-content: space-between;">
              <div>
                <strong style="font-size: 12px; color: #94a3b8; text-transform: uppercase; margin-right: 8px;"><i class="fa-solid fa-filter"></i> Filtrar Catálogo:</strong>
                <a href="/productos" style="text-decoration: none; color: ${!categoriaFiltro ? '#fff' : '#94a3b8'}; font-size: 11px; background: ${!categoriaFiltro ? '#0284c7' : 'rgba(30,41,59,0.5)'}; padding: 6px 10px; border-radius: 6px; margin-right: 5px; border: 1px solid rgba(56,189,248,0.15);"><i class="fa-solid fa-list"></i> Todos</a>
                ${navCategorias}
              </div>
              <div>
                <a href="/descargar-excel" class="btn-primario" style="background: linear-gradient(135deg, #10b981, #059669);"><i class="fa-solid fa-file-excel"></i> Descargar Excel Total</a>
              </div>
            </div>
            <form action="/productos" method="GET" style="display: flex; gap: 12px; align-items: center; border-top: 1px solid rgba(148,163,184,0.1); padding-top: 15px;">
              <input type="hidden" name="cat" value="${categoriaFiltro}">
              <input type="text" name="busqueda" value="${busqueda}" placeholder="🔍 Buscar por código, nombre o descripción..." class="input-industrial" style="margin-bottom: 0;">
              <button type="submit" class="btn-primario" style="height: 44px;"><i class="fa-solid fa-magnifying-glass"></i> Buscar</button>
              <a href="/productos${categoriaFiltro ? '?cat=' + encodeURIComponent(categoriaFiltro) : ''}" style="background: rgba(100,116,139,0.3); color: #cbd5e1; padding: 11px 16px; text-decoration: none; border-radius: 8px; font-size: 11px; font-weight: 600; text-transform: uppercase; border: 1px solid rgba(100,116,139,0.3);">Limpiar</a>
            </form>
          </div>

          <div style="display: flex; gap: 24px; flex-wrap: wrap;">
            ${esAdmin ? `
            <div style="width: 360px; display: flex; flex-direction: column; gap: 24px;">
              <div class="contenedor-tarjeta">
                <h3 style="margin-top:0; color:#38bdf8;"><i class="fa-solid fa-file-excel" style="margin-right:8px;"></i>Importar Excel</h3>
                <p style="font-size:11px; color:#f87171; margin-bottom:10px; font-weight:600;"><i class="fa-solid fa-triangle-exclamation"></i> Al subir el archivo, se REEMPLAZARÁ todo el inventario de la categoría seleccionada por los datos del nuevo Excel.</p>
                <form action="/cargar-excel" method="POST" enctype="multipart/form-data">
                  <label style="font-size: 11px; font-weight: 600; color: #94a3b8; text-transform: uppercase;">Seleccionar Sección:</label>
                  <select name="categoria" required class="input-industrial" style="margin-top: 5px;">${optionsSelectCat}</select>
                  <input type="file" name="archivo_excel" accept=".xlsx, .xls, .csv" required style="width: 100%; margin-bottom: 15px; font-size: 11px; color:#cbd5e1;"><br>
                  <button type="submit" class="btn-primario" style="background: linear-gradient(135deg, #0891b2, #0e7490); width: 100%; justify-content: center;"><i class="fa-solid fa-upload"></i> Subir e Importar</button>
                </form>
              </div>

              <div class="contenedor-tarjeta">
                <h3 style="margin-top:0;"><i class="fa-solid fa-circle-plus" style="color:#38bdf8; margin-right:8px;"></i>Agregar Ítem Manual</h3>
                <form action="/guardar-producto" method="POST" enctype="multipart/form-data">
                  <label style="font-size: 11px; font-weight: 600; color: #94a3b8; text-transform: uppercase;">Categoría:</label>
                  <select name="categoria" class="input-industrial" style="margin-top: 5px;">${optionsSelectCat}</select>
                  <input type="text" name="codigo" placeholder="Código o No. Parte" class="input-industrial">
                  <input type="text" name="nombre" required placeholder="Nombre del Equipo o Insumo" class="input-industrial">
                  <input type="number" step="0.01" name="precio_unitario" required placeholder="Precio Unitario ($)" class="input-industrial">
                  <input type="number" name="stock" value="10" placeholder="Stock" class="input-industrial">
                  <input type="text" name="tiempo_entrega" placeholder="Tiempo de Entrega (Ej. 3 días / Inmediato)" class="input-industrial">
                  <input type="text" name="dimensiones" placeholder="Dimensiones (Ej. 10x20 cm)" class="input-industrial">
                  <textarea name="descripcion" placeholder="Descripción detallada..." class="input-industrial" style="height:70px;"></textarea>
                  <label style="font-size: 11px; font-weight: 600; color: #94a3b8; text-transform: uppercase;">Foto del Producto:</label>
                  <input type="file" name="foto" accept=".jpg,.png,.jpeg" class="input-industrial" style="font-size:11px; padding:6px; margin-bottom: 15px;">
                  <button type="submit" class="btn-primario" style="background: linear-gradient(135deg, #10b981, #059669); width: 100%; justify-content: center;"><i class="fa-solid fa-floppy-disk"></i> Guardar Ítem</button>
                </form>
              </div>
            </div>
            ` : ''}

            <div class="contenedor-tarjeta" style="flex-grow: 1; overflow-x: auto;">
              <h3 style="margin-top:0;"><i class="fa-solid fa-boxes-stacked" style="color:#38bdf8; margin-right:8px;"></i>Inventario ${categoriaFiltro ? '(' + categoriaFiltro + ')' : '(General)'}</h3>
              <table>
                <tr>
                  <th>Foto</th>
                  <th>Código</th>
                  <th>Categoría</th>
                  <th>Nombre & Descripción</th>
                  <th>Dimensiones</th>
                  <th>Precio</th>
                  <th>Stock / Entrega</th>
                  <th>Acción</th>
                </tr>
                ${tabla}
              </table>
            </div>
          </div>
        </body>
      </html>
    `);
    });
});

const uploadProductoFiles = upload.single('foto');

app.post('/guardar-producto', verificarSesion, uploadProductoFiles, (req, res) => {
    const u = req.session.usuario;
    if (u.rol !== 'Administrador' && u.departamento !== 'Administración' && u.departamento !== 'Control Administrativo') return res.send("Acceso denegado");

    const { codigo, categoria, nombre, precio_unitario, stock, tiempo_entrega, dimensiones, descripcion } = req.body;
    let fotoNombre = null;
    if (req.file) {
        fotoNombre = req.file.filename + path.extname(req.file.originalname);
        fs.renameSync(req.file.path, path.join(__dirname, 'uploads/productos', fotoNombre));
    }

    db.run("INSERT INTO productos_servicios (codigo, categoria, nombre, tipo, precio_unitario, stock, tiempo_entrega, dimensiones, foto, descripcion) VALUES (?, ?, ?, 'Producto', ?, ?, ?, ?, ?, ?)",
        [codigo, categoria, nombre, precio_unitario, stock || 0, tiempo_entrega || 'Inmediato', dimensiones || '', fotoNombre || '', descripcion || ''], () => res.redirect('/productos'));
});

app.get('/editar-producto/:id', verificarSesion, (req, res) => {
    db.get("SELECT * FROM productos_servicios WHERE id = ?", [req.params.id], (err, p) => {
        if (!p) return res.send("Producto no encontrado");
        let optionsSelectCat = CATEGORIAS_VALIDAS.map(cat => `<option value="${cat}" ${cat === p.categoria ? 'selected' : ''}>${cat}</option>`).join('');

        res.send(`
            <html>
                <body style="padding-top:30px; max-width:600px; margin:auto;">
                    ${generarEncabezado('Modificar Ítem de Inventario', req)}
                    <div class="contenedor-tarjeta">
                        <h3 style="color:#38bdf8; margin-top:0;"><i class="fa-solid fa-pen"></i> Editar: ${p.nombre}</h3>
                        <form action="/actualizar-producto/${p.id}" method="POST" enctype="multipart/form-data">
                            <label style="font-size:11px; font-weight:600; color:#94a3b8;">Categoría:</label>
                            <select name="categoria" class="input-industrial" style="margin-top:5px;">${optionsSelectCat}</select>
                            <label style="font-size:11px; font-weight:600; color:#94a3b8;">Código:</label>
                            <input type="text" name="codigo" value="${p.codigo || ''}" class="input-industrial">
                            <label style="font-size:11px; font-weight:600; color:#94a3b8;">Nombre del Ítem:</label>
                            <input type="text" name="nombre" value="${p.nombre}" required class="input-industrial">
                            <label style="font-size:11px; font-weight:600; color:#94a3b8;">Precio Unitario ($):</label>
                            <input type="number" step="0.01" name="precio_unitario" value="${p.precio_unitario}" required class="input-industrial">
                            <label style="font-size:11px; font-weight:600; color:#94a3b8;">Stock:</label>
                            <input type="number" name="stock" value="${p.stock || 0}" class="input-industrial">
                            <label style="font-size:11px; font-weight:600; color:#94a3b8;">Tiempo de Entrega:</label>
                            <input type="text" name="tiempo_entrega" value="${p.tiempo_entrega || 'Inmediato'}" class="input-industrial">
                            <label style="font-size:11px; font-weight:600; color:#94a3b8;">Dimensiones:</label>
                            <input type="text" name="dimensiones" value="${p.dimensiones || ''}" class="input-industrial">
                            <label style="font-size:11px; font-weight:600; color:#94a3b8;">Descripción:</label>
                            <textarea name="descripcion" class="input-industrial" style="height:75px;">${p.descripcion || ''}</textarea>
                            <label style="font-size:11px; font-weight:600; color:#94a3b8;">Actualizar Foto:</label>
                            <input type="file" name="foto" accept=".jpg,.png,.jpeg" class="input-industrial" style="font-size:11px; padding:6px; margin-bottom:20px;">
                            <button type="submit" class="btn-primario" style="width:100%; background:linear-gradient(135deg, #10b981, #059669); justify-content:center;"><i class="fa-solid fa-check"></i> Guardar Cambios</button>
                        </form>
                        <br><a href="/productos" style="color:#38bdf8; font-weight:600; text-decoration:none; font-size:12px;"><i class="fa-solid fa-arrow-left"></i> Cancelar</a>
                    </div>
                </body>
            </html>
        `);
    });
});

app.post('/actualizar-producto/:id', verificarSesion, uploadProductoFiles, (req, res) => {
    const { categoria, codigo, nombre, precio_unitario, stock, tiempo_entrega, dimensiones, descripcion } = req.body;
    const prodId = req.params.id;

    if (req.file) {
        let fotoNombre = req.file.filename + path.extname(req.file.originalname);
        fs.renameSync(req.file.path, path.join(__dirname, 'uploads/productos', fotoNombre));
        db.run("UPDATE productos_servicios SET categoria = ?, codigo = ?, nombre = ?, precio_unitario = ?, stock = ?, tiempo_entrega = ?, dimensiones = ?, foto = ?, descripcion = ? WHERE id = ?",
            [categoria, codigo, nombre, precio_unitario, stock || 0, tiempo_entrega, dimensiones, fotoNombre, descripcion, prodId], () => res.redirect('/productos'));
    } else {
        db.run("UPDATE productos_servicios SET categoria = ?, codigo = ?, nombre = ?, precio_unitario = ?, stock = ?, tiempo_entrega = ?, dimensiones = ?, descripcion = ? WHERE id = ?",
            [categoria, codigo, nombre, precio_unitario, stock || 0, tiempo_entrega, dimensiones, descripcion, prodId], () => res.redirect('/productos'));
    }
});

const uploadExcel = upload.single('archivo_excel');
app.post('/cargar-excel', verificarSesion, uploadExcel, (req, res) => {
    const u = req.session.usuario;
    if (u.rol !== 'Administrador' && u.departamento !== 'Administración' && u.departamento !== 'Control Administrativo') return res.send("Acceso denegado");
    if (!req.file) return res.send("No se subió ningún archivo.");

    try {
        const workbook = XLSX.readFile(req.file.path);
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const filas = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        const categoriaSeleccionada = req.body.categoria || 'Tubería de Aluminio';

        db.run("DELETE FROM productos_servicios WHERE categoria = ?", [categoriaSeleccionada], () => {
            let stmt = db.prepare("INSERT INTO productos_servicios (codigo, categoria, nombre, tipo, precio_unitario, stock, tiempo_entrega, dimensiones, descripcion) VALUES (?, ?, ?, 'Producto', ?, ?, ?, ?, ?)");

            for (let i = 0; i < filas.length; i++) {
                const cols = filas[i];
                if (cols && cols.length >= 4) {
                    let codigoRaw = cols[2];
                    let precioRaw = cols[3];

                    if (codigoRaw && String(codigoRaw).toUpperCase() !== 'CÓDIGO' && !isNaN(parseFloat(precioRaw))) {
                        let codigo = String(codigoRaw).trim();
                        let precio = parseFloat(precioRaw);
                        let dimensiones = cols[1] ? String(cols[1]).trim() : '';
                        let stockOTiempo = cols[5] !== undefined ? String(cols[5]).trim() : '';
                        let stock = 0;
                        let tiempo_entrega = 'Inmediato';

                        if (stockOTiempo !== '') {
                            if (!isNaN(Number(stockOTiempo))) {
                                stock = parseInt(stockOTiempo);
                            } else {
                                tiempo_entrega = stockOTiempo;
                            }
                        }

                        let descripcionCompleta = cols[13] ? String(cols[13]).trim() : '';
                        let nombre = descripcionCompleta !== '' ? descripcionCompleta : 'Item ' + codigo;
                        if (nombre.length > 60) {
                            nombre = nombre.substring(0, 60) + '...';
                        }

                        stmt.run([codigo, categoriaSeleccionada, nombre, precio, stock, tiempo_entrega, dimensiones, descripcionCompleta]);
                    }
                }
            }

            stmt.finalize(() => {
                try { fs.unlinkSync(req.file.path); } catch (e) { }
                res.redirect('/productos');
            });
        });

    } catch (error) {
        console.error(error);
        res.send("Error al procesar el archivo Excel.");
    }
});

app.get('/descargar-excel', verificarSesion, (req, res) => {
    db.all("SELECT codigo AS 'Código', categoria AS 'Categoría', nombre AS 'Nombre', precio_unitario AS 'Precio Unitario (MXN)', stock AS 'Stock', tiempo_entrega AS 'Tiempo de Entrega', dimensiones AS 'Dimensiones', descripcion AS 'Descripción' FROM productos_servicios ORDER BY categoria, codigo", [], (err, filas) => {
        if (err) return res.send("Error al consultar la base de datos.");

        const ws = XLSX.utils.json_to_sheet(filas);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Inventario General");

        const tempFilePath = path.join(__dirname, 'uploads/temp/Inventario_Actual.xlsx');
        XLSX.writeFile(wb, tempFilePath);

        res.download(tempFilePath, 'Inventario_General.xlsx', (err) => {
            if (err) console.error(err);
            try { fs.unlinkSync(tempFilePath); } catch (e) { }
        });
    });
});

app.get('/eliminar-producto/:id', verificarSesion, (req, res) => {
    db.run("DELETE FROM productos_servicios WHERE id = ?", [req.params.id], () => res.redirect('/productos'));
});

// COTIZACIONES & HISTORIAL
app.get('/cotizaciones', verificarSesion, (req, res) => {
    const u = req.session.usuario;
    const catSeleccionada = req.query.filtro_cat || '';
    const buscarProd = req.query.buscar_prod || '';
    const filtroClienteId = req.query.filtro_cliente_id || '';
    const clientePred = req.query.cliente_id || '';
    const esAdmin = u.rol === 'Administrador' || u.departamento === 'Administración' || u.departamento === 'Control Administrativo' || u.departamento === 'Contabilidad';

    let queryClientes = esAdmin ? "SELECT id, nombre_empresa FROM clientes ORDER BY nombre_empresa" : "SELECT id, nombre_empresa FROM clientes WHERE vendedor_asignado = ? OR vendedor_asignado = ? ORDER BY nombre_empresa";
    let paramsClientes = esAdmin ? [] : [u.username, u.departamento];

    db.all(queryClientes, paramsClientes, (err, clientes) => {

        let queryProd = "SELECT id, codigo, categoria, nombre, precio_unitario FROM productos_servicios WHERE 1=1";
        let paramsProd = [];
        if (catSeleccionada) {
            queryProd += " AND categoria = ?";
            paramsProd.push(catSeleccionada);
        }
        if (buscarProd) {
            queryProd += " AND (codigo LIKE ? OR nombre LIKE ?)";
            paramsProd.push(`%${buscarProd}%`, `%${buscarProd}%`);
        }
        queryProd += " ORDER BY categoria, nombre";

        db.all(queryProd, paramsProd, (err, productos) => {

            let queryCot = `
        SELECT c.*, cl.nombre_empresa, pf.estatus_prefactura, 
               SUM(d.cantidad * d.precio_unitario) as subtotal
        FROM cotizaciones c 
        JOIN clientes cl ON c.cliente_id = cl.id 
        LEFT JOIN cotizacion_detalles d ON c.id = d.cotizacion_id
        LEFT JOIN prefacturas pf ON c.id = pf.cotizacion_id
        WHERE 1=1
      `;
            let paramsCot = [];

            if (!esAdmin) {
                queryCot += " AND (c.vendedor = ? OR c.vendedor = ?)";
                paramsCot.push(u.username, u.departamento);
            }

            if (filtroClienteId) {
                queryCot += " AND c.cliente_id = ?";
                paramsCot.push(filtroClienteId);
            }

            queryCot += " GROUP BY c.id ORDER BY c.id DESC";

            db.all(queryCot, paramsCot, (err, cotizaciones) => {
                let opcionesClientes = clientes.map(c => `<option value="${c.id}" ${c.id == clientePred ? 'selected' : ''}>${c.nombre_empresa}</option>`).join('');
                let opcionesFiltroClientes = clientes.map(c => `<option value="${c.id}" ${c.id == filtroClienteId ? 'selected' : ''}>${c.nombre_empresa}</option>`).join('');
                let opcionesProductos = productos.map(p => `<option value="${p.id}">[${p.categoria}] [${p.codigo || 'S/N'}] ${p.nombre} - $${formatoDinero(p.precio_unitario)}</option>`).join('');
                let optionsSelectCat = CATEGORIAS_VALIDAS.map(cat => `<option value="${cat}" ${cat === catSeleccionada ? 'selected' : ''}>${cat}</option>`).join('');

                let tablaCotizaciones = '';
                if (cotizaciones) {
                    cotizaciones.forEach(cot => {
                        const subtotalBruto = cot.subtotal || 0;
                        const montoDescuento = subtotalBruto * ((cot.descuento || 0) / 100);
                        const subtotal = subtotalBruto - montoDescuento;
                        const iva = subtotal * 0.16;
                        const total = subtotal + iva;
                        const porcPagado = cot.porcentaje_pagado || 0;
                        const saldoPendiente = total * (1 - (porcPagado / 100));

                        let badgeEstado = cot.estado === 'Aceptada' ? '<span style="background: rgba(16,185,129,0.2); color: #34d399; padding: 4px 8px; border-radius: 6px; font-weight: 600; font-size: 11px;"><i class="fa-solid fa-circle-check"></i> Aceptada</span>' :
                            cot.estado === 'Rechazada' ? '<span style="background: rgba(239,68,68,0.2); color: #f87171; padding: 4px 8px; border-radius: 6px; font-weight: 600; font-size: 11px;"><i class="fa-solid fa-circle-xmark"></i> Rechazada</span>' :
                                '<span style="background: rgba(56,189,248,0.2); color: #38bdf8; padding: 4px 8px; border-radius: 6px; font-weight: 600; font-size: 11px;"><i class="fa-solid fa-paper-plane"></i> Enviada</span>';

                        let facturaLink = cot.archivo_factura ? `<a href="/facturas/${cot.archivo_factura}" target="_blank" rel="noopener noreferrer" style="color:#34d399; font-weight:600; background:rgba(16,185,129,0.1); padding:4px 8px; border-radius:6px; display:inline-block;"><i class="fa-solid fa-file-invoice-dollar"></i> Factura Timbrada</a>` : '<span style="color:#fbbf24; font-size:11px;"><i class="fa-solid fa-clock"></i> Pendiente Timbrado</span>';

                        let almacenBadge = cot.estatus_almacen === 'Concluido' ? '<span style="color:#34d399; font-weight:600;"><i class="fa-solid fa-check"></i> Logística Concluida</span>' : '<span style="color:#a78bfa;">En Proceso</span>';

                        let botonEliminarRechazada = cot.estado === 'Rechazada' ? `<a href="/eliminar-cotizacion/${cot.id}" onclick="return confirm('¿Eliminar esta cotización rechazada?')" class="btn-peligro" style="padding: 4px 8px; font-size:10px; margin-top:2px; display:inline-block;"><i class="fa-solid fa-trash"></i> Borrar</a>` : '';

                        tablaCotizaciones += `
              <tr>
                <td style="padding: 12px;"><strong>COT-${cot.id}</strong><br><small style="color:#94a3b8;">${cot.vendedor}</small></td>
                <td style="padding: 12px;"><strong>${cot.nombre_empresa || 'Cliente'}</strong></td>
                <td style="padding: 12px;">${badgeEstado}</td>
                <td style="padding: 12px;">
                  Total: <b style="color:#f8fafc;">$${formatoDinero(total)}</b> ${cot.moneda}<br>
                  <small style="color:#38bdf8;">Abonado: ${porcPagado}% ($${formatoDinero(total * (porcPagado / 100))})</small><br>
                  <small style="color:${saldoPendiente > 0 ? '#f87171' : '#34d399'};">Saldo: $${formatoDinero(saldoPendiente)}</small>
                </td>
                <td style="padding: 12px;">${cot.archivo_oc ? `<a href="/ocs/${cot.archivo_oc}" target="_blank" rel="noopener noreferrer" style="color:#38bdf8; font-weight:600;"><i class="fa-solid fa-file-pdf"></i> Ver OC</a>` : 'Sin OC'}</td>
                <td style="padding: 12px;">${facturaLink}</td>
                <td style="padding: 12px;">${almacenBadge}<br>${cot.reporte_entrega ? `<a href="/almacen/${cot.reporte_entrega}" target="_blank" style="color:#a78bfa; font-size:11px; font-weight:600;"><i class="fa-solid fa-file-lines"></i> Reporte PDF</a>` : ''}</td>
                <td style="padding: 12px;">
                  <a href="/imprimir-cotizacion/${cot.id}" target="_blank" rel="noopener noreferrer" style="background: linear-gradient(135deg, #0ea5e9, #0284c7); color: white; padding: 5px 8px; text-decoration: none; border-radius: 6px; font-size: 10px; font-weight: 600; display: inline-block; margin-bottom: 3px;"><i class="fa-solid fa-print"></i> PDF</a>
                  <a href="/modificar-cotizacion/${cot.id}" style="background: linear-gradient(135deg, #d97706, #b45309); color: white; padding: 5px 8px; text-decoration: none; border-radius: 6px; font-size: 10px; font-weight: 600; display: inline-block; margin-bottom: 3px;"><i class="fa-solid fa-pen"></i></a>
                  <a href="/gestionar-estatus/${cot.id}" style="background: linear-gradient(135deg, #7c3aed, #6d28d9); color: white; padding: 5px 8px; text-decoration: none; border-radius: 6px; font-size: 10px; font-weight: 600; display: inline-block; margin-bottom: 3px;"><i class="fa-solid fa-gears"></i></a>
                  ${botonEliminarRechazada}
                </td>
              </tr>
            `;
                    });
                }

                res.send(`
          <html>
            <head>
              <script>
                function agregarLineaProducto() {
                  const contenedor = document.getElementById('lista-productos');
                  const nuevaFila = document.createElement('div');
                  nuevaFila.style.display = 'flex';
                  nuevaFila.style.gap = '10px';
                  nuevaFila.style.marginBottom = '10px';
                  nuevaFila.innerHTML = \`
                    <select name="productos[]" required class="input-industrial" style="margin-bottom:0;">
                      <option value="">-- Seleccionar Producto --</option>
                      ${opcionesProductos}
                    </select>
                    <input type="number" name="cantidades[]" required min="1" value="1" placeholder="Cant." class="input-industrial" style="width: 70px; margin-bottom:0;">
                    <button type="button" onclick="this.parentElement.remove()" style="background: #ef4444; color: white; border: none; padding: 8px 12px; cursor: pointer; border-radius: 6px;"><i class="fa-solid fa-xmark"></i></button>
                  \`;
                  contenedor.appendChild(nuevaFila);
                }
              </script>
            </head>
            <body style="max-width: 1450px; margin: auto;">
              ${generarEncabezado('Módulo de Cotizaciones & Facturación', req)}

              <div class="contenedor-tarjeta" style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:15px;">
                <form action="/cotizaciones" method="GET" style="display: flex; gap: 15px; align-items: center; flex-wrap: wrap; margin-bottom:0;">
                  <strong style="font-size: 12px; color: #94a3b8; text-transform: uppercase;"><i class="fa-solid fa-filter"></i> Filtrar por Cliente:</strong>
                  <select name="filtro_cliente_id" class="input-industrial" style="width: 250px; margin-bottom: 0;">
                    <option value="">-- Todos los Clientes --</option>
                    ${opcionesFiltroClientes}
                  </select>
                  <button type="submit" class="btn-primario"><i class="fa-solid fa-magnifying-glass"></i> Filtrar Cotizaciones</button>
                  <a href="/cotizaciones" style="background: rgba(100,116,139,0.3); color: #cbd5e1; padding: 10px 15px; text-decoration: none; border-radius: 8px; font-size: 11px; font-weight: 600; text-transform: uppercase; border: 1px solid rgba(100,116,139,0.3);">Ver Todos</a>
                </form>
                <div>
                  <a href="/pagos-clientes" class="btn-primario" style="background: linear-gradient(135deg, #10b981, #059669);"><i class="fa-solid fa-wallet"></i> Ventana de Control de Pagos</a>
                </div>
              </div>

              <div style="display: flex; gap: 24px; flex-wrap: wrap;">
                <div class="contenedor-tarjeta" style="width: 420px; height: fit-content;">
                  <h3 style="margin-top:0;"><i class="fa-solid fa-file-circle-plus" style="color:#38bdf8; margin-right:8px;"></i>Nueva Cotización</h3>
                  
                  <form action="/cotizaciones" method="GET" style="background: rgba(30,41,59,0.5); padding: 12px; border-radius: 8px; margin-bottom: 15px; border: 1px solid rgba(56,189,248,0.15);">
                    <label style="font-size: 11px; font-weight: 600; color: #38bdf8; text-transform: uppercase;">1. Filtrar inventario por catálogo:</label>
                    <select name="filtro_cat" class="input-industrial" style="margin-top: 5px; margin-bottom: 8px;">
                      <option value="">-- Seleccionar Catálogo --</option>
                      ${optionsSelectCat}
                    </select>
                    <input type="text" name="buscar_prod" value="${buscarProd}" placeholder="Buscar código o nombre..." class="input-industrial" style="margin-bottom: 8px;">
                    <button type="submit" class="btn-primario" style="background: linear-gradient(135deg, #475569, #334155); width: 100%; justify-content: center;"><i class="fa-solid fa-boxes-stacked"></i> Cargar Catálogo</button>
                  </form>

                  <form action="/guardar-cotizacion" method="POST">
                    <label style="font-size: 11px; font-weight: 600; color: #94a3b8; text-transform: uppercase;">Seleccionar Cliente:</label>
                    <select name="cliente_id" required class="input-industrial" style="margin-top: 5px;">
                      <option value="">-- Seleccionar --</option>
                      ${opcionesClientes}
                    </select>

                    <label style="font-size: 11px; font-weight: 600; color: #94a3b8; text-transform: uppercase;">Vigencia:</label>
                    <select name="vigencia_dias" class="input-industrial" style="margin-top: 5px;">
                      <option value="7">7 Días</option>
                      <option value="10">10 Días</option>
                      <option value="15" selected>15 Días (Estándar)</option>
                    </select>

                    <label style="font-size: 11px; font-weight: 600; color: #94a3b8; text-transform: uppercase;">Moneda:</label>
                    <select name="moneda" required class="input-industrial" style="margin-top: 5px;">
                      <option value="MXN">Pesos (MXN)</option>
                      <option value="USD">Dólares (USD)</option>
                    </select>

                    <label style="font-size: 11px; font-weight: 600; color: #94a3b8; text-transform: uppercase;">Descuento Comercial:</label>
                    <select name="descuento" class="input-industrial" style="margin-top: 5px; margin-bottom: 15px;">
                      <option value="0">Sin Descuento (0%)</option>
                      <option value="3">3%</option>
                      <option value="5">5%</option>
                      <option value="8">8%</option>
                      <option value="10">10%</option>
                      <option value="15">15%</option>
                      <option value="20">20%</option>
                    </select>

                    <input type="hidden" name="vendedor" value="${u.username}">

                    <label style="font-size: 11px; font-weight: 600; color: #38bdf8; text-transform: uppercase;">Partidas (${productos.length} ítems cargados):</label>
                    <div id="lista-productos" style="margin-top: 8px;">
                      <div style="display: flex; gap: 10px; margin-bottom: 8px;">
                        <select name="productos[]" required class="input-industrial" style="margin-bottom:0;">
                          <option value="">-- Seleccionar Producto --</option>
                          ${opcionesProductos}
                        </select>
                        <input type="number" name="cantidades[]" required min="1" value="1" placeholder="Cant." class="input-industrial" style="width: 70px; margin-bottom:0;">
                      </div>
                    </div>

                    <button type="button" onclick="agregarLineaProducto()" style="background: rgba(100,116,139,0.3); color: #cbd5e1; border: 1px solid rgba(100,116,139,0.4); padding: 8px 12px; cursor: pointer; border-radius: 6px; font-size: 11px; margin-bottom: 15px; font-weight: 600;"><i class="fa-solid fa-plus"></i> Añadir otra partida</button><br>

                    <button type="submit" class="btn-primario" style="background: linear-gradient(135deg, #10b981, #059669); width: 100%; justify-content: center;"><i class="fa-solid fa-file-shield"></i> Generar Cotización</button>
                  </form>
                </div>

                <div class="contenedor-tarjeta" style="flex-grow: 1; overflow-x: auto;">
                  <h3 style="margin-top:0;"><i class="fa-solid fa-clipboard-user" style="color:#38bdf8; margin-right:8px;"></i>Historial de Cotizaciones & Órdenes</h3>
                  <table>
                    <tr>
                      <th>Folio</th>
                      <th>Cliente</th>
                      <th>Estatus</th>
                      <th>Finanzas / Pagos</th>
                      <th>Orden Compra</th>
                      <th>Factura PDF</th>
                      <th>Almacén</th>
                      <th>Acción</th>
                    </tr>
                    ${tablaCotizaciones}
                  </table>
                </div>
              </div>
            </body>
          </html>
        `);
            });
        });
    });
});

app.post('/guardar-cotizacion', verificarSesion, (req, res) => {
    const { cliente_id, vendedor, moneda, descuento, vigencia_dias } = req.body;
    let productos = req.body['productos[]'] || req.body.productos;
    let cantidades = req.body['cantidades[]'] || req.body.cantidades;

    if (!Array.isArray(productos)) productos = [productos];
    if (!Array.isArray(cantidades)) cantidades = [cantidades];

    db.run("INSERT INTO cotizaciones (cliente_id, vendedor, moneda, descuento, vigencia_dias) VALUES (?, ?, ?, ?, ?)", [cliente_id, vendedor, moneda, descuento || 0, vigencia_dias || 15], function (err) {
        if (err) return res.send("Error al guardar cotización");
        const cotizacion_id = this.lastID;

        let stmt = db.prepare("INSERT INTO cotizacion_detalles (cotizacion_id, producto_id, cantidad, precio_unitario) VALUES (?, ?, ?, (SELECT precio_unitario FROM productos_servicios WHERE id = ?))");
        for (let i = 0; i < productos.length; i++) {
            if (productos[i] && cantidades[i]) {
                stmt.run([cotizacion_id, productos[i], cantidades[i], productos[i]]);
            }
        }
        stmt.finalize(() => {
            res.redirect('/cotizaciones');
        });
    });
});

app.get('/eliminar-cotizacion/:id', verificarSesion, (req, res) => {
    db.run("DELETE FROM cotizaciones WHERE id = ? AND estado = 'Rechazada'", [req.params.id], () => {
        res.redirect('/cotizaciones');
    });
});

// CONTROL DE PAGOS Y ABONOS POR CLIENTE
app.get('/pagos-clientes', verificarSesion, (req, res) => {
    const queryCotizacionesAceptadas = `
        SELECT c.*, cl.nombre_empresa, 
               SUM(d.cantidad * d.precio_unitario) as subtotal
        FROM cotizaciones c 
        JOIN clientes cl ON c.cliente_id = cl.id 
        LEFT JOIN cotizacion_detalles d ON c.id = d.cotizacion_id
        WHERE c.estado = 'Aceptada'
        GROUP BY c.id ORDER BY c.id DESC
    `;

    db.all(queryCotizacionesAceptadas, [], (err, cotizaciones) => {
        let opcionesCotizaciones = cotizaciones.map(cot => {
            const sub = (cot.subtotal || 0) * (1 - ((cot.descuento || 0) / 100));
            const total = sub * 1.16;
            return `<option value="${cot.id}">COT-${cot.id} — ${cot.nombre_empresa} (Total: $${formatoDinero(total)} ${cot.moneda})</option>`;
        }).join('');

        db.all("SELECT p.*, c.id as cot_id, cl.nombre_empresa FROM pagos_cotizacion p JOIN cotizaciones c ON p.cotizacion_id = c.id JOIN clientes cl ON c.cliente_id = cl.id ORDER BY p.id DESC", [], (err, pagosList) => {
            let tablaPagos = '';
            pagosList.forEach(pag => {
                tablaPagos += `<tr>
                    <td style="padding:10px;"><strong>COT-${pag.cot_id}</strong></td>
                    <td style="padding:10px;">${pag.nombre_empresa}</td>
                    <td style="padding:10px; color:#34d399; font-weight:600;">$${formatoDinero(pag.monto_abono)}</td>
                    <td style="padding:10px;"><span class="badge-tech" style="background:rgba(56,189,248,0.1); color:#38bdf8;">${pag.porcentaje_abono}%</span></td>
                    <td style="padding:10px;">${pag.referencia || 'S/Ref'}</td>
                    <td style="padding:10px;">${pag.fecha_pago}</td>
                    <td style="padding:10px;">${pag.archivo_comprobante ? `<a href="/pagos/${pag.archivo_comprobante}" target="_blank" style="color:#38bdf8; font-weight:600;"><i class="fa-solid fa-receipt"></i> Ver Comprobante</a>` : 'Sin archivo'}</td>
                </tr>`;
            });

            let tablaResumen = '';
            cotizaciones.forEach(cot => {
                const sub = (cot.subtotal || 0) * (1 - ((cot.descuento || 0) / 100));
                const total = sub * 1.16;
                const porc = cot.porcentaje_pagado || 0;
                const abonado = total * (porc / 100);
                const saldo = total - abonado;

                let estadoFinanciero = porc >= 100 ? '<span style="color:#34d399; font-weight:600;"><i class="fa-solid fa-circle-check"></i> Factura Pagada (100%) ✅</span>' : `<span style="color:#f87171; font-weight:600;"><i class="fa-solid fa-clock"></i> Con Deuda ($${formatoDinero(saldo)})</span>`;

                tablaResumen += `<tr>
                    <td style="padding:10px;"><strong>COT-${cot.id}</strong></td>
                    <td style="padding:10px;">${cot.nombre_empresa}</td>
                    <td style="padding:10px;">$${formatoDinero(total)} ${cot.moneda}</td>
                    <td style="padding:10px; color:#38bdf8;">${porc}% ($${formatoDinero(abonado)})</td>
                    <td style="padding:10px;">${estadoFinanciero}</td>
                </tr>`;
            });

            res.send(`
                <html>
                    <body style="max-width: 1450px; margin: auto;">
                        ${generarEncabezado('Módulo de Control de Pagos, Abonos & Deudas', req)}
                        
                        <div style="display:flex; gap:24px; flex-wrap:wrap; margin-bottom:24px;">
                            <div class="contenedor-tarjeta" style="width:380px; height:fit-content;">
                                <h3 style="margin-top:0;"><i class="fa-solid fa-hand-holding-dollar" style="color:#38bdf8; margin-right:8px;"></i>Registrar Nuevo Abono / Anticipo</h3>
                                <form action="/guardar-pago-cliente" method="POST" enctype="multipart/form-data">
                                    <label style="font-size:11px; font-weight:600; color:#94a3b8; text-transform:uppercase;">Seleccionar Cotización Aceptada:</label>
                                    <select name="cotizacion_id" required class="input-industrial" style="margin-top:5px;">
                                        <option value="">-- Seleccionar --</option>
                                        ${opcionesCotizaciones}
                                    </select>
                                    
                                    <label style="font-size:11px; font-weight:600; color:#94a3b8; text-transform:uppercase;">Monto del Abono ($):</label>
                                    <input type="number" step="0.01" name="monto_abono" required placeholder="Ej. 25000" class="input-industrial" style="margin-top:5px;">
                                    
                                    <label style="font-size:11px; font-weight:600; color:#94a3b8; text-transform:uppercase;">Porcentaje que representa este abono (%):</label>
                                    <select name="porcentaje_abono" class="input-industrial" style="margin-top:5px;">
                                        <option value="25">25% (Anticipo estándar)</option>
                                        <option value="50">50% (Avance medio)</option>
                                        <option value="75">75% (Avanzado)</option>
                                        <option value="100">100% (Pago Total / Liquidación)</option>
                                    </select>

                                    <input type="text" name="referencia" placeholder="Referencia bancaria o nota..." class="input-industrial">
                                    
                                    <label style="font-size:11px; font-weight:600; color:#94a3b8; text-transform:uppercase;">Comprobante de Pago (PDF / Imagen):</label>
                                    <input type="file" name="archivo_comprobante" accept=".pdf,.jpg,.png" required class="input-industrial" style="font-size:11px; margin-bottom:15px; padding:6px;">

                                    <button type="submit" class="btn-primario" style="width:100%; background:linear-gradient(135deg, #10b981, #059669); justify-content:center;"><i class="fa-solid fa-floppy-disk"></i> Registrar y Actualizar Saldo</button>
                                </form>
                            </div>

                            <div class="contenedor-tarjeta" style="flex-grow:1; overflow-x:auto;">
                                <h3 style="margin-top:0;"><i class="fa-solid fa-scale-balanced" style="color:#38bdf8; margin-right:8px;"></i>Estatus de Facturas: Pagadas vs Deudas (Comunicado a Contabilidad)</h3>
                                <table>
                                    <tr>
                                        <th>Folio</th>
                                        <th>Empresa</th>
                                        <th>Monto Total</th>
                                        <th>Abonado</th>
                                        <th>Estatus Financiero</th>
                                    </tr>
                                    ${tablaResumen}
                                </table>
                            </div>
                        </div>

                        <div class="contenedor-tarjeta" style="overflow-x:auto;">
                            <h3 style="margin-top:0;"><i class="fa-solid fa-receipt" style="color:#38bdf8; margin-right:8px;"></i>Historial de Comprobantes de Pago Adjuntos</h3>
                            <table>
                                <tr>
                                    <th>Folio</th>
                                    <th>Cliente</th>
                                    <th>Monto Abonado</th>
                                    <th>Porcentaje</th>
                                    <th>Referencia</th>
                                    <th>Fecha</th>
                                    <th>Comprobante</th>
                                </tr>
                                ${tablaPagos}
                            </table>
                        </div>
                    </body>
                </html>
            `);
        });
    });
});

const uploadComprobantePago = upload.single('archivo_comprobante');
app.post('/guardar-pago-cliente', verificarSesion, uploadComprobantePago, (req, res) => {
    const { cotizacion_id, monto_abono, porcentaje_abono, referencia } = req.body;
    let compFile = null;
    if (req.file) {
        const ext = path.extname(req.file.originalname);
        compFile = req.file.filename + ext;
        fs.renameSync(req.file.path, path.join(__dirname, 'uploads/pagos', compFile));
    }

    db.run("INSERT INTO pagos_cotizacion (cotizacion_id, monto_abono, porcentaje_abono, referencia, archivo_comprobante) VALUES (?, ?, ?, ?, ?)",
        [cotizacion_id, monto_abono, porcentaje_abono, referencia, compFile], () => {
            db.run("UPDATE cotizaciones SET porcentaje_pagado = ?, archivo_pago = ? WHERE id = ?", [porcentaje_abono, compFile, cotizacion_id], () => {
                res.redirect('/pagos-clientes');
            });
        });
});

// MODIFICAR COTIZACIÓN
app.get('/modificar-cotizacion/:id', verificarSesion, (req, res) => {
    db.get("SELECT c.*, cl.nombre_empresa FROM cotizaciones c JOIN clientes cl ON c.cliente_id = cl.id WHERE c.id = ?", [req.params.id], (err, cot) => {
        if (!cot) return res.send("Cotización no encontrada");
        db.all("SELECT d.*, p.nombre FROM cotizacion_detalles d JOIN productos_servicios p ON d.producto_id = p.id WHERE d.cotizacion_id = ?", [cot.id], (err, items) => {
            let listaItems = '';
            items.forEach(it => {
                listaItems += `<div style="display:flex; gap:12px; margin-bottom:10px; align-items:center;">
          <input type="text" value="${it.nombre}" readonly class="input-industrial" style="margin-bottom:0; background:rgba(15,23,42,0.8);">
          <input type="number" name="cantidades_${it.id}" value="${it.cantidad}" min="1" class="input-industrial" style="width:90px; margin-bottom:0;">
          <a href="/eliminar-linea-cotizacion/${it.id}/${cot.id}" class="btn-peligro"><i class="fa-solid fa-trash"></i></a>
        </div>`;
            });

            res.send(`
        <html>
          <body style="padding-top: 30px; max-width: 700px; margin: auto;">
            ${generarEncabezado('Modificar Partidas de Cotización', req)}
            <div class="contenedor-tarjeta">
              <h2 style="color:#38bdf8; margin-top:0;"><i class="fa-solid fa-pen-to-square"></i> COT-${cot.id} — ${cot.nombre_empresa}</h2>
              <form action="/actualizar-cotizacion-cantidades/${cot.id}" method="POST">
                <h4 style="color:#94a3b8; font-size:12px; text-transform:uppercase; margin-bottom:12px;">Partidas Actuales:</h4>
                ${listaItems}
                <button type="submit" class="btn-primario" style="width:100%; margin-top:20px; background:linear-gradient(135deg, #10b981, #059669); justify-content:center;"><i class="fa-solid fa-check"></i> Guardar Modificaciones</button>
              </form>
              <br><a href="/cotizaciones" style="color:#38bdf8; font-weight:600; text-decoration:none; display:block; text-align:center; font-size:12px;"><i class="fa-solid fa-arrow-left"></i> Regresar</a>
            </div>
          </body>
        </html>
      `);
        });
    });
});

app.get('/eliminar-linea-cotizacion/:detalle_id/:cot_id', verificarSesion, (req, res) => {
    db.run("DELETE FROM cotizacion_detalles WHERE id = ?", [req.params.detalle_id], () => {
        res.redirect(`/modificar-cotizacion/${req.params.cot_id}`);
    });
});

app.post('/actualizar-cotizacion-cantidades/:id', verificarSesion, (req, res) => {
    const cotId = req.params.id;
    db.all("SELECT id FROM cotizacion_detalles WHERE cotizacion_id = ?", [cotId], (err, items) => {
        if (items) {
            items.forEach(it => {
                let nuevaCant = req.body[`cantidades_${it.id}`];
                if (nuevaCant) {
                    db.run("UPDATE cotizacion_detalles SET cantidad = ? WHERE id = ?", [nuevaCant, it.id]);
                }
            });
        }
        res.redirect('/cotizaciones');
    });
});

// GESTIÓN DE ESTATUS, OC Y PAGOS
app.get('/gestionar-estatus/:id', verificarSesion, (req, res) => {
    db.get("SELECT c.*, cl.nombre_empresa FROM cotizaciones c JOIN clientes cl ON c.cliente_id = cl.id WHERE c.id = ?", [req.params.id], (err, cot) => {
        if (err || !cot) return res.send("Cotización no encontrada");

        res.send(`
      <html>
        <body style="padding-top: 30px; max-width: 650px; margin: auto;">
          ${generarEncabezado('Gestión Financiera & Estatus', req)}
          <div class="contenedor-tarjeta">
            <h3 style="color:#38bdf8; margin-top:0;"><i class="fa-solid fa-gears"></i> Folio: COT-${cot.id} — ${cot.nombre_empresa}</h3>
            
            <form action="/guardar-estatus/${cot.id}" method="POST" enctype="multipart/form-data">
              <label style="font-size: 11px; font-weight: 600; color: #94a3b8; text-transform: uppercase;">Estatus de la Cotización:</label>
              <select name="estado" class="input-industrial" style="margin-top: 5px;">
                <option value="Enviada" ${cot.estado === 'Enviada' ? 'selected' : ''}>Enviada (Pendiente)</option>
                <option value="Aceptada" ${cot.estado === 'Aceptada' ? 'selected' : ''}>Aceptada ✅</option>
                <option value="Rechazada" ${cot.estado === 'Rechazada' ? 'selected' : ''}>Rechazada ❌</option>
              </select>

              <label style="font-size: 11px; font-weight: 600; color: #94a3b8; text-transform: uppercase;">Porcentaje Acumulado de Pagos del Cliente (%):</label>
              <select name="porcentaje_pagado" class="input-industrial" style="margin-top: 5px;">
                <option value="0" ${cot.porcentaje_pagado == 0 ? 'selected' : ''}>0% (Sin abonos)</option>
                <option value="25" ${cot.porcentaje_pagado == 25 ? 'selected' : ''}>25% (Anticipo)</option>
                <option value="50" ${cot.porcentaje_pagado == 50 ? 'selected' : ''}>50% (Avance medio)</option>
                <option value="75" ${cot.porcentaje_pagado == 75 ? 'selected' : ''}>75% (Avanzado)</option>
                <option value="100" ${cot.porcentaje_pagado == 100 ? 'selected' : ''}>100% (Pagado Totalmente ✅)</option>
              </select>

              <label style="font-size: 11px; font-weight: 600; color: #94a3b8; text-transform: uppercase;">Motivo / Razón:</label>
              <input type="text" name="motivo_estado" value="${cot.motivo_estado || ''}" placeholder="Notas de cobro..." class="input-industrial" style="margin-top: 5px;">

              <label style="font-size: 11px; font-weight: 600; color: #94a3b8; text-transform: uppercase;">Orden de Compra (OC):</label>
              ${cot.archivo_oc ? `<p style="font-size: 11px; color: #38bdf8;"><i class="fa-solid fa-file-pdf"></i> Actual: <a href="/ocs/${cot.archivo_oc}" target="_blank" style="color:#38bdf8;">Ver OC</a></p>` : ''}
              <input type="file" name="archivo_oc" accept=".pdf, .jpg, .png" class="input-industrial" style="margin-top: 5px; font-size: 11px; padding: 6px;">

              <label style="font-size: 11px; font-weight: 600; color: #94a3b8; text-transform: uppercase;">Comprobante de Depósito / Pago:</label>
              ${cot.archivo_pago ? `<p style="font-size: 11px; color: #34d399;"><i class="fa-solid fa-receipt"></i> Actual: <a href="/pagos/${cot.archivo_pago}" target="_blank" style="color:#34d399;">Ver Depósito</a></p>` : ''}
              <input type="file" name="archivo_pago" accept=".pdf, .jpg, .png" class="input-industrial" style="margin-top: 5px; font-size: 11px; padding: 6px; margin-bottom: 20px;">

              <button type="submit" class="btn-primario" style="width: 100%; background: linear-gradient(135deg, #10b981, #059669); justify-content: center; padding: 12px;"><i class="fa-solid fa-check"></i> Guardar Actualización</button>
            </form>
            <br>
            <a href="/cotizaciones" style="display: block; text-align: center; color: #38bdf8; font-weight: 600; text-decoration: none; font-size: 12px;"><i class="fa-solid fa-arrow-left"></i> Volver al listado</a>
          </div>
        </body>
      </html>
    `);
    });
});

const cpUpload = upload.fields([{ name: 'archivo_oc', maxCount: 1 }, { name: 'archivo_pago', maxCount: 1 }]);
app.post('/guardar-estatus/:id', verificarSesion, cpUpload, (req, res) => {
    const cotId = req.params.id;
    const { estado, motivo_estado, porcentaje_pagado } = req.body;

    let ocFile = null;
    if (req.files && req.files['archivo_oc'] && req.files['archivo_oc'][0]) {
        const f = req.files['archivo_oc'][0];
        const ext = path.extname(f.originalname);
        ocFile = f.filename + ext;
        fs.renameSync(f.path, path.join(__dirname, 'uploads/ocs', ocFile));
    }

    let pagoFile = null;
    if (req.files && req.files['archivo_pago'] && req.files['archivo_pago'][0]) {
        const f = req.files['archivo_pago'][0];
        const ext = path.extname(f.originalname);
        pagoFile = f.filename + ext;
        fs.renameSync(f.path, path.join(__dirname, 'uploads/pagos', pagoFile));
    }

    let q = "UPDATE cotizaciones SET estado = ?, motivo_estado = ?, porcentaje_pagado = ?";
    let p = [estado, motivo_estado, porcentaje_pagado || 0];
    if (ocFile) { q += ", archivo_oc = ?"; p.push(ocFile); }
    if (pagoFile) { q += ", archivo_pago = ?"; p.push(pagoFile); }
    q += " WHERE id = ?"; p.push(cotId);

    db.run(q, p, () => res.redirect('/cotizaciones'));
});

// PDF / REIMPRESIÓN COTIZACIÓN
app.get('/imprimir-cotizacion/:id', verificarSesion, (req, res) => {
    const queryEncabezado = `
    SELECT c.id, c.vendedor, c.moneda, c.descuento, c.vigencia_dias, c.fecha_registro, cl.nombre_empresa, cl.rfc, cl.telefono, cl.email, cl.ubicacion, cl.contacto_nombre 
    FROM cotizaciones c 
    JOIN clientes cl ON c.cliente_id = cl.id 
    WHERE c.id = ?
  `;

    db.get(queryEncabezado, [req.params.id], (err, cot) => {
        if (err || !cot) return res.send("Cotización no encontrada");

        db.all("SELECT d.cantidad, d.precio_unitario, p.codigo, p.categoria, p.nombre as producto, p.descripcion FROM cotizacion_detalles d JOIN productos_servicios p ON d.producto_id = p.id WHERE d.cotizacion_id = ?", [req.params.id], (err, detalles) => {
            let filasTabla = '';
            let subtotalBruto = 0;

            if (detalles) {
                detalles.forEach(item => {
                    const importe = item.cantidad * item.precio_unitario;
                    subtotalBruto += importe;
                    filasTabla += `
            <tr>
              <td style="padding: 10px; border-bottom: 1px solid #ddd;">
                <span style="font-size: 11px; background: #e9ecef; padding: 2px 5px; border-radius: 3px;">${item.categoria}</span>
                <strong style="color: #0056b3;">[${item.codigo || 'S/N'}]</strong> <strong>${item.producto}</strong><br>
                <small style="color: #666;">${item.descripcion || ''}</small>
              </td>
              <td style="padding: 10px; border-bottom: 1px solid #ddd; text-align: center;">${item.cantidad}</td>
              <td style="padding: 10px; border-bottom: 1px solid #ddd; text-align: right;">$${formatoDinero(item.precio_unitario)}</td>
              <td style="padding: 10px; border-bottom: 1px solid #ddd; text-align: right;">$${formatoDinero(importe)}</td>
            </tr>
          `;
                });
            }

            const porcDescuento = cot.descuento || 0;
            const montoDescuento = subtotalBruto * (porcDescuento / 100);
            const subtotalNeto = subtotalBruto - montoDescuento;
            const iva = subtotalNeto * 0.16;
            const total = subtotalNeto + iva;

            res.send(`
        <html>
          <head>
            <title>Cotización COT-${cot.id}</title>
            <style>
              body { font-family: Arial, sans-serif; padding: 30px; color: #333; max-width: 850px; margin: auto; }
              .header { display: flex; justify-content: space-between; align-items: flex-start; border-bottom: 2px solid #0056b3; padding-bottom: 15px; }
              .logo { height: 70px; margin-bottom: 10px; }
              .empresa-info { font-size: 13px; color: #555; line-height: 1.4; }
              .grid-info { display: flex; justify-content: space-between; margin-top: 20px; gap: 20px; }
              .box { background: #f8f9fa; padding: 15px; border-radius: 5px; flex: 1; font-size: 13px; line-height: 1.5; border: 1px solid #e9ecef; }
              table { width: 100%; border-collapse: collapse; margin-top: 25px; font-size: 14px; }
              th { background-color: #0056b3; color: white; padding: 10px; text-align: left; }
              .totales { margin-top: 20px; text-align: right; font-size: 15px; }
              .banco-box { margin-top: 30px; background: #eef6ff; padding: 15px; border-radius: 5px; border-left: 4px solid #0056b3; font-size: 13px; }
              .btn-imprimir { background: #28a745; color: white; padding: 10px 20px; border: none; font-size: 15px; cursor: pointer; border-radius: 5px; margin-bottom: 20px; font-weight: bold; }
              @media print { .btn-imprimir { display: none; } }
            </style>
          </head>
          <body>
            <button class="btn-imprimir" onclick="window.print()">🖨️ Guardar como PDF / Imprimir</button>

            <div class="header">
              <div>
                <img src="/logo.png" class="logo" alt="Logo Empresa"><br>
                <strong style="font-size: 15px; color: #333;">${DATOS_MI_EMPRESA.nombre}</strong>
                <div class="empresa-info">
                  RFC: ${DATOS_MI_EMPRESA.rfc}<br>
                  Teléfono: ${DATOS_MI_EMPRESA.telefono}<br>
                  Correo: ${DATOS_MI_EMPRESA.correo}<br>
                  ${DATOS_MI_EMPRESA.direccion}
                </div>
              </div>
              <div style="text-align: right;">
                <h2 style="color: #0056b3; margin: 0;">COTIZACIÓN</h2>
                <h3 style="margin: 5px 0; color: #555;">FOLIO: COT-${cot.id}</h3>
                <p style="margin: 0; font-size: 13px; color: #666;">Fecha: ${new Date(cot.fecha_registro).toLocaleDateString()}</p>
                <p style="margin: 0; font-size: 13px; color: #d97706;"><strong>Vigencia:</strong> ${cot.vigencia_dias || 15} días</p>
                <p style="margin: 5px 0 0 0; font-size: 13px; color: #333;"><strong>Vendedor:</strong> ${cot.vendedor || 'N/A'}</p>
                <p style="margin: 5px 0 0 0; font-size: 14px; color: #0056b3;"><strong>Moneda:</strong> ${cot.moneda}</p>
              </div>
            </div>

            <div class="grid-info">
              <div class="box">
                <strong style="color: #0056b3;">DATOS DEL CLIENTE</strong><br>
                <strong>Empresa:</strong> ${cot.nombre_empresa}<br>
                <strong>RFC:</strong> ${cot.rfc || 'N/A'}<br>
                <strong>Contacto:</strong> ${cot.contacto_nombre || 'N/A'}<br>
                <strong>Teléfono:</strong> ${cot.telefono || 'N/A'}<br>
                <strong>Correo:</strong> ${cot.email || 'N/A'}<br>
                <strong>Ubicación:</strong> ${cot.ubicacion || 'N/A'}
              </div>
            </div>

            <table>
              <thead>
                <tr>
                  <th>Categoría, Código y Concepto / Ítem</th>
                  <th style="text-align: center;">Cant.</th>
                  <th style="text-align: right;">Precio Unitario</th>
                  <th style="text-align: right;">Importe</th>
                </tr>
              </thead>
              <tbody>
                ${filasTabla}
              </tbody>
            </table>

            <div class="totales">
              <p style="margin: 4px 0;">Subtotal Bruto: $${formatoDinero(subtotalBruto)}</p>
              ${porcDescuento > 0 ? `<p style="margin: 4px 0; color: #28a745;">Descuento (${porcDescuento}%): -$${formatoDinero(montoDescuento)}</p>` : ''}
              <p style="margin: 4px 0;">Subtotal Neto: $${formatoDinero(subtotalNeto)}</p>
              <p style="margin: 4px 0; color: #d9534f;">IVA (16%): $${formatoDinero(iva)}</p>
              <h2 style="color: #0056b3; margin: 8px 0 0 0;">Total: $${formatoDinero(total)} ${cot.moneda}</h2>
            </div>

            <div class="banco-box">
              <strong style="color: #0056b3;">INFORMACIÓN DE PAGO / DEPÓSITO BANCARIO</strong><br>
              <strong>Banco:</strong> ${DATOS_MI_EMPRESA.banco} | 
              <strong>Titular:</strong> ${DATOS_MI_EMPRESA.titular}<br>
              <strong>No. de Cuenta:</strong> ${DATOS_MI_EMPRESA.cuenta} | 
              <strong>CLABE Interbancaria:</strong> ${DATOS_MI_EMPRESA.clabe}
            </div>

          </body>
        </html>
      `);
        });
    });
});

// MÓDULO DE CONTABILIDAD
app.get('/contabilidad', verificarSesion, (req, res) => {
    const u = req.session.usuario;
    if (u.departamento !== 'Contabilidad' && u.rol !== 'Administrador' && u.departamento !== 'Control Administrativo') {
        return res.send("<script>alert('Acceso exclusivo para Contabilidad'); window.location.href='/';</script>");
    }

    const queryAll = `
    SELECT c.*, cl.nombre_empresa, cl.rfc, cl.archivo_csf, pf.estatus_prefactura,
           SUM(d.cantidad * d.precio_unitario) as subtotal
    FROM cotizaciones c 
    JOIN clientes cl ON c.cliente_id = cl.id 
    LEFT JOIN cotizacion_detalles d ON c.id = d.cotizacion_id
    LEFT JOIN prefacturas pf ON c.id = pf.cotizacion_id
    GROUP BY c.id ORDER BY c.id DESC
  `;

    db.all(queryAll, [], (err, cotis) => {
        let tablaCotis = '';
        cotis.forEach(cot => {
            let estadoBadge = cot.estado === 'Aceptada' ? '<span style="color:#34d399; font-weight:600;"><i class="fa-solid fa-circle-check"></i> Aceptada</span>' :
                cot.estado === 'Rechazada' ? '<span style="color:#f87171; font-weight:600;"><i class="fa-solid fa-circle-xmark"></i> Rechazada</span>' : '<span style="color:#38bdf8;">Enviada</span>';

            let finanzasBadge = (cot.porcentaje_pagado || 0) >= 100 ? '<span style="color:#34d399; font-weight:600;"><i class="fa-solid fa-circle-check"></i> Pagado (100%) ✅</span>' : `<span style="color:#f87171; font-weight:600;"><i class="fa-solid fa-clock"></i> Deuda (${cot.porcentaje_pagado || 0}%)</span>`;

            tablaCotis += `<tr>
        <td style="padding:12px;"><strong>COT-${cot.id}</strong><br><small style="color:#94a3b8;">${cot.vendedor || 'General'}</small></td>
        <td style="padding:12px;"><strong>${cot.nombre_empresa}</strong><br><small style="color:#94a3b8;">RFC: ${cot.rfc || 'S/N'}</small><br>
            ${cot.archivo_csf ? `<a href="/csf/${cot.archivo_csf}" target="_blank" style="color:#a78bfa; font-weight:600; font-size:11px;"><i class="fa-solid fa-file-pdf"></i> Ver CSF</a>` : '<span style="color:#f87171; font-size:11px;">Sin CSF</span>'}
        </td>
        <td style="padding:12px;">${estadoBadge}<br><b style="font-size:11px;">${finanzasBadge}</b></td>
        <td style="padding:12px;">
          ${cot.archivo_oc ? `<a href="/ocs/${cot.archivo_oc}" target="_blank" rel="noopener noreferrer" style="color:#38bdf8; font-weight:600; display:block;"><i class="fa-solid fa-file-pdf"></i> Ver OC</a>` : '<span style="color:#64748b;">Sin OC</span>'}
          ${cot.archivo_pago ? `<a href="/pagos/${cot.archivo_pago}" target="_blank" rel="noopener noreferrer" style="color:#34d399; font-weight:600; display:block; margin-top:4px;"><i class="fa-solid fa-receipt"></i> Ver Comprobante Abono</a>` : '<span style="color:#64748b; display:block; margin-top:4px;">Sin Abono</span>'}
        </td>
        <td style="padding:12px;">
          <a href="/prefactura/${cot.id}" class="btn-primario" style="padding:5px 10px; font-size:10px; background:linear-gradient(135deg, #7c3aed, #6d28d9); text-decoration:none; display:inline-block; margin-bottom:4px;"><i class="fa-solid fa-file-invoice"></i> Prefactura SAT</a>
          <form action="/subir-factura/${cot.id}" method="POST" enctype="multipart/form-data" style="margin-top:4px;">
            <label style="font-size:10px; font-weight:600; color:#38bdf8;">Factura PDF Timbrada:</label>
            <input type="file" name="archivo_factura" accept=".pdf" required style="font-size:10px; width:150px; display:block; margin:3px 0; color:#cbd5e1;">
            <button type="submit" class="btn-primario" style="padding:4px 8px; font-size:10px; background:linear-gradient(135deg, #10b981, #059669);"><i class="fa-solid fa-stamp"></i> Timbrar & Comisión</button>
          </form>
          ${cot.archivo_factura ? '<br><a href="/facturas/' + cot.archivo_factura + '" target="_blank" style="color:#34d399; font-weight:600; font-size:11px;"><i class="fa-solid fa-download"></i> Descargar Factura</a>' : ''}
        </td>
      </tr>`;
        });

        res.send(`
      <html>
        <body style="max-width: 1450px; margin: auto;">
          ${generarEncabezado('Panel de Contabilidad & Facturación', req)}
          <div class="contenedor-tarjeta">
            <h3 style="margin-top:0;"><i class="fa-solid fa-vault" style="color:#38bdf8; margin-right:8px;"></i>Control Financiero, Pagos, Deudas y Timbrado de Facturas</h3>
            <table>
              <tr>
                <th>Folio / Vendedor</th>
                <th>Cliente & CSF</th>
                <th>Estatus & Finanzas</th>
                <th>Documentos OC / Comprobantes de Pago</th>
                <th>Timbrado de Factura PDF</th>
              </tr>
              ${tablaCotis}
            </table>
          </div>
        </body>
      </html>
    `);
    });
});

const uploadFactura = upload.single('archivo_factura');
app.post('/subir-factura/:id', verificarSesion, uploadFactura, (req, res) => {
    const cotId = req.params.id;
    let facturaFile = null;
    if (req.file) {
        const ext = path.extname(req.file.originalname);
        facturaFile = req.file.filename + ext;
        fs.renameSync(req.file.path, path.join(__dirname, 'uploads/facturas', facturaFile));
    }
    db.run("UPDATE cotizaciones SET archivo_factura = ? WHERE id = ?", [facturaFile, cotId], () => {
        db.run("INSERT INTO prefacturas (cotizacion_id, estatus_prefactura) VALUES (?, 'Timbrada')", [cotId], () => {
            res.redirect('/contabilidad');
        });
    });
});

// MÓDULO DE ALBACÉN Y LOGÍSTICA (PEDIDO Y REMISIÓN)
app.get('/almacen-logistica', verificarSesion, (req, res) => {
    const u = req.session.usuario;
    if (u.departamento !== 'Almacén' && u.rol !== 'Administrador' && u.departamento !== 'Administración' && u.departamento !== 'Control Administrativo') {
        return res.send("<script>alert('Acceso exclusivo para Almacén y Logística'); window.location.href='/';</script>");
    }

    const queryAlmacen = `
    SELECT c.*, cl.nombre_empresa, cl.ubicacion, SUM(d.cantidad * d.precio_unitario) as subtotal
    FROM cotizaciones c 
    JOIN clientes cl ON c.cliente_id = cl.id 
    LEFT JOIN cotizacion_detalles d ON c.id = d.cotizacion_id
    WHERE c.estado = 'Aceptada'
    GROUP BY c.id ORDER BY c.id DESC
  `;

    db.all(queryAlmacen, [], (err, lista) => {
        let tablaLogistica = '';
        lista.forEach(item => {
            tablaLogistica += `<tr>
        <td style="padding:12px;"><strong>COT-${item.id}</strong><br><small style="color:#94a3b8;">${item.vendedor}</small></td>
        <td style="padding:12px;"><strong>${item.nombre_empresa}</strong><br><small style="color:#94a3b8;">${item.ubicacion || 'Sin dirección'}</small></td>
        <td style="padding:12px;">
          <a href="/pedido-almacen/${item.id}" target="_blank" class="btn-primario" style="padding:4px 8px; font-size:10px; background:linear-gradient(135deg, #0284c7, #0369a1); text-decoration:none;"><i class="fa-solid fa-file-lines"></i> Ver Pedido Suministro</a>
        </td>
        <td style="padding:12px;">
          <span style="color:${item.estatus_almacen === 'Concluido' ? '#34d399' : '#fbbf24'}; font-weight:600;"><i class="fa-solid fa-circle-dot"></i> ${item.estatus_almacen || 'Pendiente'}</span>
          ${item.remision_almacen ? `<br><a href="/almacen/${item.remision_almacen}" target="_blank" style="color:#34d399; font-weight:600; font-size:11px;"><i class="fa-solid fa-file-invoice"></i> Ver Remisión PDF</a>` : ''}
        </td>
        <td style="padding:12px;">
          <form action="/concluir-almacen/${item.id}" method="POST" enctype="multipart/form-data">
            <label style="font-size:10px; font-weight:600; color:#38bdf8;">Subir Remisión / Reporte PDF:</label>
            <input type="file" name="remision_almacen" accept=".pdf,.jpg,.png" required style="font-size:10px; width:150px; display:block; margin:3px 0; color:#cbd5e1;">
            <button type="submit" class="btn-primario" style="padding:4px 8px; font-size:10px; background:linear-gradient(135deg, #10b981, #059669);"><i class="fa-solid fa-check"></i> Generar Remisión & Concluir</button>
          </form>
        </td>
      </tr>`;
        });

        res.send(`
      <html>
        <body style="max-width: 1450px; margin: auto;">
          ${generarEncabezado('Módulo de Almacén, Suministros y Remisiones', req)}
          <div class="contenedor-tarjeta">
            <h3 style="margin-top:0;"><i class="fa-solid fa-warehouse" style="color:#38bdf8; margin-right:8px;"></i>Pedidos de Materiales a Suministrar & Remisiones</h3>
            <table>
              <tr>
                <th>Folio</th>
                <th>Cliente & Destino</th>
                <th>Suministro de Materiales</th>
                <th>Estatus Logístico</th>
                <th>Remisión de Almacén</th>
              </tr>
              ${tablaLogistica}
            </table>
          </div>
        </body>
      </html>
    `);
    });
});

app.get('/pedido-almacen/:id', verificarSesion, (req, res) => {
    const cotId = req.params.id;
    db.get("SELECT c.*, cl.nombre_empresa, cl.ubicacion, cl.telefono, cl.contacto_nombre FROM cotizaciones c JOIN clientes cl ON c.cliente_id = cl.id WHERE c.id = ?", [cotId], (err, cot) => {
        db.all("SELECT d.cantidad, p.codigo, p.categoria, p.nombre, p.dimensiones FROM cotizacion_detalles d JOIN productos_servicios p ON d.producto_id = p.id WHERE d.cotizacion_id = ?", [cotId], (err, items) => {
            res.send(`
                <html>
                    <body style="font-family:'Inter',sans-serif; padding:30px; max-width:850px; margin:auto; background:#0f172a; color:#f8fafc; border-radius:12px; border:1px solid rgba(56,189,248,0.2);">
                        <div style="display:flex; justify-content:space-between; border-bottom:2px solid rgba(56,189,248,0.3); padding-bottom:15px;">
                            <div>
                                <h2 style="margin:0; color:#38bdf8;"><i class="fa-solid fa-boxes-packing"></i> ORDEN DE SURTIDO / PEDIDO A ALMACÉN</h2>
                                <p style="margin:4px 0 0 0; font-size:12px; color:#94a3b8;">C.C. INSTALACIONES INDUSTRIALES S.A. DE C.V.</p>
                            </div>
                            <div style="text-align:right;">
                                <h3 style="margin:0; color:#f8fafc;">COTIZACIÓN: COT-${cot.id}</h3>
                                <p style="margin:4px 0 0 0; font-size:11px; color:#94a3b8;">Fecha: ${new Date().toLocaleDateString()}</p>
                            </div>
                        </div>

                        <div style="background:rgba(30,41,59,0.5); padding:15px; border-radius:8px; font-size:13px; margin:20px 0; border:1px solid rgba(56,189,248,0.15);">
                            <strong style="color:#38bdf8;">DESTINO Y DATOS DE ENTREGA:</strong><br>
                            <b>Empresa:</b> ${cot.nombre_empresa}<br>
                            <b>Contacto:</b> ${cot.contacto_nombre || 'N/A'} (${cot.telefono || 'S/Tel'})<br>
                            <b>Ubicación de Obra / Almacén:</b> ${cot.ubicacion || 'N/A'}<br>
                        </div>

                        <table style="width:100%; border-collapse:collapse; font-size:13px;">
                            <tr style="background:rgba(30,41,59,0.9); color:#38bdf8;">
                                <th style="padding:10px; text-align:left;">Código / Categoría</th>
                                <th style="padding:10px; text-align:left;">Descripción del Material</th>
                                <th style="padding:10px; text-align:center;">Dimensiones</th>
                                <th style="padding:10px; text-align:center;">Cantidad a Surtir</th>
                            </tr>
                            ${items.map(it => `
                                <tr>
                                    <td style="padding:10px; border-bottom:1px solid rgba(148,163,184,0.1);"><strong>[${it.codigo || 'S/N'}]</strong><br><small style="color:#94a3b8;">${it.categoria}</small></td>
                                    <td style="padding:10px; border-bottom:1px solid rgba(148,163,184,0.1);">${it.nombre}</td>
                                    <td style="padding:10px; border-bottom:1px solid rgba(148,163,184,0.1); text-align:center;">${it.dimensiones || 'N/A'}</td>
                                    <td style="padding:10px; border-bottom:1px solid rgba(148,163,184,0.1); text-align:center; font-weight:700; color:#34d399; font-size:15px;">${it.cantidad}</td>
                                </tr>
                            `).join('')}
                        </table>

                        <div style="margin-top:30px; display:flex; gap:15px; justify-content:flex-end;">
                            <button onclick="window.print()" style="background:linear-gradient(135deg, #0ea5e9, #0284c7); color:white; border:none; padding:12px 20px; font-weight:600; border-radius:8px; cursor:pointer;"><i class="fa-solid fa-print"></i> Imprimir Orden de Surtido</button>
                            <a href="/almacen-logistica" style="background:rgba(100,116,139,0.3); color:#cbd5e1; padding:12px 20px; text-decoration:none; border-radius:8px; font-weight:600; text-align:center; display:inline-flex; align-items:center; gap:8px;"><i class="fa-solid fa-arrow-left"></i> Volver</a>
                        </div>
                    </body>
                </html>
            `);
        });
    });
});

const uploadRemision = upload.single('remision_almacen');
app.post('/concluir-almacen/:id', verificarSesion, uploadRemision, (req, res) => {
    const cotId = req.params.id;
    let remFile = null;
    if (req.file) {
        const ext = path.extname(req.file.originalname);
        remFile = req.file.filename + ext;
        fs.renameSync(req.file.path, path.join(__dirname, 'uploads/almacen', remFile));
    }
    db.run("UPDATE cotizaciones SET estatus_almacen = 'Concluido', remision_almacen = ?, reporte_entrega = ? WHERE id = ?", [remFile, remFile, cotId], () => {
        res.redirect('/almacen-logistica');
    });
});

app.get('/prefactura/:id', verificarSesion, (req, res) => {
    const cotId = req.params.id;
    db.get("SELECT c.*, cl.*, SUM(d.cantidad * d.precio_unitario) as subtotal FROM cotizaciones c JOIN clientes cl ON c.cliente_id = cl.id LEFT JOIN cotizacion_detalles d ON c.id = d.cotizacion_id WHERE c.id = ?", [cotId], (err, cot) => {
        db.all("SELECT d.cantidad, d.precio_unitario, p.nombre FROM cotizacion_detalles d JOIN productos_servicios p ON d.producto_id = p.id WHERE d.cotizacion_id = ?", [cotId], (err, items) => {
            let sub = (cot.subtotal || 0) * (1 - ((cot.descuento || 0) / 100));
            let iva = sub * 0.16;
            let total = sub + iva;

            res.send(`
        <html>
          <body style="font-family:'Inter',sans-serif; padding:30px; max-width:850px; margin:auto; background:#0f172a; color:#f8fafc; border-radius:12px; border:1px solid rgba(56,189,248,0.2);">
            <div style="display:flex; justify-content:space-between; border-bottom:2px solid rgba(56,189,248,0.3); padding-bottom:15px;">
              <div>
                <h2 style="margin:0; color:#38bdf8;">PREFACTURA FISCAL SAT</h2>
                <p style="margin:4px 0 0 0; font-size:12px; color:#94a3b8;">C.C. INSTALACIONES INDUSTRIALES S.A. DE C.V.</p>
              </div>
              <div style="text-align:right;">
                <h3 style="margin:0; color:#f8fafc;">COTIZACIÓN: COT-${cot.id}</h3>
                <p style="margin:4px 0 0 0; font-size:11px; color:#94a3b8;">Fecha: ${new Date().toLocaleDateString()}</p>
              </div>
            </div>

            <form action="/guardar-y-timbrar/${cot.id}" method="POST" style="margin-top:20px;">
              <div style="background:rgba(30,41,59,0.5); padding:15px; border-radius:8px; font-size:13px; margin-bottom:20px; border:1px solid rgba(56,189,248,0.15);">
                <strong style="color:#38bdf8;">DATOS FISCALES DEL CLIENTE:</strong><br>
                <b>Empresa:</b> ${cot.nombre_empresa}<br>
                <b>RFC:</b> ${cot.rfc || 'N/A'}<br>
                <b>Dirección:</b> ${cot.ubicacion || 'N/A'}<br>
              </div>

              <table style="width:100%; border-collapse:collapse; font-size:13px;">
                <tr style="background:rgba(30,41,59,0.9); color:#38bdf8;">
                  <th style="padding:10px; text-align:left;">Concepto</th>
                  <th style="padding:10px; text-align:center;">Cant.</th>
                  <th style="padding:10px; text-align:right;">P. Unitario</th>
                  <th style="padding:10px; text-align:right;">Importe</th>
                </tr>
                ${items.map(it => `
                  <tr>
                    <td style="padding:10px; border-bottom:1px solid rgba(148,163,184,0.1);">${it.nombre}</td>
                    <td style="padding:10px; border-bottom:1px solid rgba(148,163,184,0.1); text-align:center;">${it.cantidad}</td>
                    <td style="padding:10px; border-bottom:1px solid rgba(148,163,184,0.1); text-align:right;">$${formatoDinero(it.precio_unitario)}</td>
                    <td style="padding:10px; border-bottom:1px solid rgba(148,163,184,0.1); text-align:right;">$${formatoDinero(it.cantidad * it.precio_unitario)}</td>
                  </tr>
                `).join('')}
              </table>

              <div style="text-align:right; margin-top:20px; font-size:14px;">
                <p style="margin:4px 0;">Subtotal: $${formatoDinero(sub)}</p>
                <p style="margin:4px 0; color:#f87171;">IVA (16%): $${formatoDinero(iva)}</p>
                <h3 style="color:#34d399; margin:8px 0;">Total Fiscal: $${formatoDinero(total)} ${cot.moneda}</h3>
              </div>

              <div style="margin-top:30px; display:flex; gap:15px; flex-wrap:wrap;">
                <button type="submit" style="background:linear-gradient(135deg, #10b981, #059669); color:white; border:none; padding:12px 20px; font-weight:600; flex:1; border-radius:8px; cursor:pointer;"><i class="fa-solid fa-floppy-disk"></i> Guardar Prefactura</button>
                <button type="button" onclick="window.print()" style="background:linear-gradient(135deg, #0ea5e9, #0284c7); color:white; border:none; padding:12px 20px; font-weight:600; border-radius:8px; cursor:pointer;"><i class="fa-solid fa-print"></i> Imprimir</button>
                <a href="https://sistema.ecofactura.mx/ingresar/" target="_blank" rel="noopener noreferrer" style="background:linear-gradient(135deg, #7c3aed, #6d28d9); color:white; padding:12px 20px; text-decoration:none; border-radius:8px; font-weight:600; text-align:center; display:flex; align-items:center; justify-content:center; gap:8px; flex:1;"><i class="fa-solid fa-bolt"></i> Abrir Eco Factura</a>
              </div>
            </form>
            <br><a href="/contabilidad" style="color:#38bdf8; font-weight:600; text-decoration:none; font-size:12px;"><i class="fa-solid fa-arrow-left"></i> Volver al Panel Contable</a>
          </body>
        </html>
      `);
        });
    });
});

app.post('/guardar-y-timbrar/:id', verificarSesion, (req, res) => {
    const cotId = req.params.id;
    db.run("INSERT INTO prefacturas (cotizacion_id, estatus_prefactura) VALUES (?, 'Timbrada')", [cotId], () => {
        res.send("<script>alert('¡Prefactura guardada con éxito!'); window.location.href='/contabilidad';</script>");
    });
});

// RESÚMENES DE VENTAS, COMISIONES, GASTOS & GRÁFICA DE RENDIMIENTO CON ALERTAS INTELIGENTES
app.get('/gastos-ventas', verificarSesion, (req, res) => {
    const u = req.session.usuario;
    const esAdminTotal = u.rol === 'Administrador' || u.departamento === 'Administración' || u.departamento === 'Control Administrativo' || u.departamento === 'Contabilidad';

    let queryCot = `
        SELECT c.*, cl.nombre_empresa, 
               SUM(d.cantidad * d.precio_unitario) as subtotal, 
               SUM(d.cantidad * (d.precio_unitario * 0.7)) as costovendido 
        FROM cotizaciones c 
        JOIN clientes cl ON c.cliente_id = cl.id 
        LEFT JOIN cotizacion_detalles d ON c.id = d.cotizacion_id 
    `;
    let paramsCot = [];

    if (!esAdminTotal) {
        queryCot += " WHERE c.vendedor = ? OR c.vendedor = ?";
        paramsCot.push(u.username, u.departamento);
    }
    queryCot += " GROUP BY c.id ORDER BY c.id DESC";

    db.all(queryCot, paramsCot, (err, cotizacionesList) => {
        let queryGastos = esAdminTotal ? "SELECT * FROM gastos_operativos ORDER BY fecha DESC" : "SELECT * FROM gastos_operativos WHERE vendedor = ? ORDER BY fecha DESC";
        let paramsGastos = esAdminTotal ? [] : [u.username];

        db.all(queryGastos, paramsGastos, (err, gastosList) => {

            let totalVendidoAceptado = 0;
            let totalComisionGanadaMensual = 0;
            let numAceptadas = 0;
            let numRechazadas = 0;
            let numEnviadas = 0;
            let tablaDesgloseClientes = '';

            let mesActualStr = new Date().toISOString().slice(0, 7);

            cotizacionesList.forEach(cot => {
                const sub = (cot.subtotal || 0) * (1 - ((cot.descuento || 0) / 100));
                const total = sub * 1.16;

                if (cot.estado === 'Aceptada') {
                    totalVendidoAceptado += total;
                    numAceptadas++;
                } else if (cot.estado === 'Rechazada') {
                    numRechazadas++;
                } else {
                    numEnviadas++;
                }

                let comisionFactura = 0;
                let estadoFacturaHtml = '<span style="color:#fbbf24;"><i class="fa-solid fa-clock"></i> Sin timbrar</span>';

                if (cot.archivo_factura) {
                    let porcentajeCom = total >= 500000 ? 0.05 : 0.03;
                    comisionFactura = total * porcentajeCom;

                    let fechaCot = (cot.fecha_registro || '').slice(0, 7);
                    if (fechaCot === mesActualStr) {
                        totalComisionGanadaMensual += comisionFactura;
                    }

                    estadoFacturaHtml = `<span style="color:#34d399; font-weight:600;"><i class="fa-solid fa-circle-check"></i> Timbrada (${porcentajeCom * 100}%)</span>`;
                }

                tablaDesgloseClientes += `<tr>
                    <td style="padding:12px;">COT-${cot.id} <small style="color:#94a3b8;">(${cot.vendedor})</small></td>
                    <td style="padding:12px;"><strong>${cot.nombre_empresa}</strong></td>
                    <td style="padding:12px;">${cot.estado}</td>
                    <td style="padding:12px;">$${formatoDinero(total)} ${cot.moneda}</td>
                    <td style="padding:12px;">${estadoFacturaHtml}</td>
                    <td style="padding:12px; color:#34d399; font-weight:600;">$${formatoDinero(comisionFactura)}</td>
                </tr>`;
            });

            let gastoTotalDiario = 0;
            let gastoTotalMensual = 0;
            let hoyStr = new Date().toISOString().slice(0, 10);

            let tablaGastosIndividuales = '';
            gastosList.forEach(g => {
                let fechaGasto = (g.fecha || '').slice(0, 10);
                let mesGasto = (g.fecha || '').slice(0, 7);
                let montoG = g.monto || 0;

                if (mesGasto === mesActualStr) {
                    gastoTotalMensual += montoG;
                }
                if (fechaGasto === hoyStr) {
                    gastoTotalDiario += montoG;
                }

                tablaGastosIndividuales += `<tr>
                    <td style="padding:12px;"><strong>${g.concepto}</strong><br><small style="color:#94a3b8;">${g.categoria}</small></td>
                    <td style="padding:12px; color:#f87171; font-weight:600;">$${formatoDinero(montoG)}</td>
                    <td style="padding:12px;">${g.fecha}</td>
                    <td style="padding:12px;">${g.comprobante ? `<a href="/gastos/${g.comprobante}" target="_blank" style="color:#38bdf8;"><i class="fa-solid fa-receipt"></i> Ver</a>` : 'S/C'}</td>
                </tr>`;
            });

            let mensajeAlertaGrafica = "";
            let colorAlertaBg = "";
            let colorAlertaBorde = "";

            if (gastoTotalMensual > totalComisionGanadaMensual) {
                mensajeAlertaGrafica = "🔴 <b>Alerta Roja:</b> Gastas más de lo que comisionas por mes. <b>¡Estás mal en ventas!</b>";
                colorAlertaBg = "rgba(239, 68, 68, 0.15)";
                colorAlertaBorde = "#ef4444";
            } else if (totalComisionGanadaMensual > 80000) {
                mensajeAlertaGrafica = "🟢 <b>¡Excelente Desempeño!</b> Tus comisiones rebasan los $80,000 mensuales. ¡Tienes unas muy buenas ventas y el sistema te felicita! 🏆";
                colorAlertaBg = "rgba(16, 185, 129, 0.15)";
                colorAlertaBorde = "#10b981";
            } else {
                mensajeAlertaGrafica = "🟡 <b>Alerta Normal:</b> Tus comisiones están por encima de tus gastos. <b>¡Estás bien en ventas!</b>";
                colorAlertaBg = "rgba(251, 191, 36, 0.15)";
                colorAlertaBorde = "#fbbf24";
            }

            let maxEscala = Math.max(totalComisionGanadaMensual, gastoTotalMensual, 80000, 1);
            let anchoComision = Math.round((totalComisionGanadaMensual / maxEscala) * 100);
            let anchoGasto = Math.round((gastoTotalMensual / maxEscala) * 100);

            res.send(`
        <html>
          <body style="max-width: 1450px; margin: auto;">
            ${generarEncabezado('Métrica de Ventas, Comisiones & Gastos', req)}
            
            <div class="contenedor-tarjeta" style="background: ${colorAlertaBg}; border-left: 6px solid ${colorAlertaBorde};">
              <h3 style="margin-top:0; color:#f8fafc;"><i class="fa-solid fa-chart-pie" style="color:#38bdf8; margin-right:8px;"></i>Gráfica y Alerta de Rendimiento Mensual (${mesActualStr})</h3>
              <p style="font-size: 14px; margin-bottom: 15px; color:#cbd5e1;">${mensajeAlertaGrafica}</p>
              
              <div style="background: rgba(15,23,42,0.8); padding: 16px; border-radius: 8px; margin-bottom: 15px; border: 1px solid rgba(56,189,248,0.15);">
                <div style="margin-bottom: 14px;">
                  <span style="font-size: 11px; font-weight: 600; color: #34d399; text-transform: uppercase;"><i class="fa-solid fa-coins"></i> Comisiones Mensuales: $${formatoDinero(totalComisionGanadaMensual)} MXN</span>
                  <div style="background: rgba(30,41,59,0.8); border-radius: 6px; height: 14px; width: 100%; margin-top: 6px; overflow: hidden; border: 1px solid rgba(16,185,129,0.3);">
                    <div style="background: linear-gradient(90deg, #10b981, #34d399); height: 100%; width: ${anchoComision}%; transition: width 0.5s ease;"></div>
                  </div>
                </div>
                <div>
                  <span style="font-size: 11px; font-weight: 600; color: #f87171; text-transform: uppercase;"><i class="fa-solid fa-wallet"></i> Gastos Mensuales: $${formatoDinero(gastoTotalMensual)} MXN</span>
                  <div style="background: rgba(30,41,59,0.8); border-radius: 6px; height: 14px; width: 100%; margin-top: 6px; overflow: hidden; border: 1px solid rgba(239,68,68,0.3);">
                    <div style="background: linear-gradient(90deg, #dc2626, #ef4444); height: 100%; width: ${anchoGasto}%; transition: width 0.5s ease;"></div>
                  </div>
                </div>
              </div>
            </div>

            <div style="display: flex; gap: 24px; margin-bottom: 24px; flex-wrap: wrap;">
              <div class="contenedor-tarjeta" style="flex:1;">
                <h3 style="margin-top:0; color:#38bdf8;"><i class="fa-solid fa-chart-column" style="margin-right:8px;"></i>Métricas Comerciales Individuales</h3>
                <div style="display:flex; gap:15px; margin-bottom:15px; flex-wrap:wrap;">
                  <div style="background:rgba(3,105,161,0.2); padding:12px; border-radius:8px; flex:1; text-align:center; border: 1px solid rgba(56,189,248,0.2);">
                    <h4 style="margin:0; color:#38bdf8; font-size:11px; text-transform:uppercase;">Concretadas</h4>
                    <p style="font-size:20px; font-weight:700; margin:6px 0; color:#f8fafc;">${numAceptadas}</p>
                  </div>
                  <div style="background:rgba(185,28,28,0.2); padding:12px; border-radius:8px; flex:1; text-align:center; border: 1px solid rgba(239,68,68,0.2);">
                    <h4 style="margin:0; color:#f87171; font-size:11px; text-transform:uppercase;">Rechazadas</h4>
                    <p style="font-size:20px; font-weight:700; margin:6px 0; color:#f8fafc;">${numRechazadas}</p>
                  </div>
                  <div style="background:rgba(217,119,6,0.2); padding:12px; border-radius:8px; flex:1; text-align:center; border: 1px solid rgba(251,191,36,0.2);">
                    <h4 style="margin:0; color:#fbbf24; font-size:11px; text-transform:uppercase;">En Proceso</h4>
                    <p style="font-size:20px; font-weight:700; margin:6px 0; color:#f8fafc;">${numEnviadas}</p>
                  </div>
                </div>
                <p style="font-size:13px; color:#cbd5e1;">Ventas Netas Concretadas (Aceptadas): <b style="color:#f8fafc;">$${formatoDinero(totalVendidoAceptado)}</b></p>
                <h3 style="color:#34d399; margin:12px 0 0 0; font-size:15px;"><i class="fa-solid fa-money-bill-trend-up"></i> Comisión Total (Facturas Timbradas): $${formatoDinero(totalComisionGanadaMensual)} MXN</h3>
              </div>

              <div class="contenedor-tarjeta" style="width: 380px;">
                <h3 style="margin-top:0;"><i class="fa-solid fa-receipt" style="color:#38bdf8; margin-right:8px;"></i>Registrar Gasto</h3>
                <form action="/guardar-gasto" method="POST" enctype="multipart/form-data">
                  <input type="text" name="concepto" required placeholder="Concepto (Ej. Gasolina / Viáticos)" class="input-industrial">
                  <label style="font-size:11px; font-weight:600; color:#94a3b8; text-transform:uppercase;">Categoría:</label>
                  <select name="categoria" class="input-industrial" style="margin-top:5px;">
                    <option value="Vehículo / Gasolina">🚗 Vehículo / Gasolina</option>
                    <option value="Viáticos">🍽️ Viáticos / Comidas</option>
                    <option value="Teléfono / Planes">📱 Teléfono</option>
                    <option value="Otros">🏢 Otros</option>
                  </select>
                  <input type="number" step="0.01" name="monto" required placeholder="Monto ($)" class="input-industrial">
                  <label style="font-size:11px; font-weight:600; color:#94a3b8; text-transform:uppercase;">Comprobante:</label>
                  <input type="file" name="comprobante" accept=".pdf,.jpg,.png" class="input-industrial" style="font-size:11px; margin-bottom:15px; padding:6px;">
                  <button type="submit" class="btn-primario" style="width:100%; background:linear-gradient(135deg, #0284c7, #0369a1); justify-content: center;"><i class="fa-solid fa-floppy-disk"></i> Registrar Gasto</button>
                </form>
              </div>
            </div>

            <div style="display: flex; gap: 24px; margin-bottom: 24px; flex-wrap: wrap;">
              <div class="contenedor-tarjeta" style="flex:1;">
                <h3 style="margin-top:0;"><i class="fa-solid fa-wallet" style="color:#38bdf8; margin-right:8px;"></i>Resumen de Gastos</h3>
                <p style="font-size:13px; color:#cbd5e1;"><i class="fa-solid fa-calendar-day" style="color:#38bdf8;"></i> Gasto Hoy: <strong style="color:#38bdf8;">$${formatoDinero(gastoTotalDiario)} MXN</strong></p>
                <p style="font-size:13px; color:#cbd5e1;"><i class="fa-solid fa-calendar-days" style="color:#f87171;"></i> Gasto Mensual (Cierre): <strong style="color:#f87171;">$${formatoDinero(gastoTotalMensual)} MXN</strong></p>
                <small style="color:#94a3b8;">Sincronizado con Control Administrativo.</small>
              </div>
              <div class="contenedor-tarjeta" style="flex:2; overflow-x:auto;">
                <h3 style="margin-top:0;"><i class="fa-solid fa-list-check" style="color:#38bdf8; margin-right:8px;"></i>Mis Gastos Registrados</h3>
                <table>
                  <tr><th>Concepto</th><th>Monto</th><th>Fecha</th><th>Comprobante</th></tr>
                  ${tablaGastosIndividuales}
                </table>
              </div>
            </div>

            <div class="contenedor-tarjeta">
              <h3 style="margin-top:0;"><i class="fa-solid fa-file-invoice-dollar" style="color:#38bdf8; margin-right:8px;"></i>Desglose de Ventas y Comisiones por Factura Timbrada</h3>
              <table>
                <tr>
                  <th>Folio</th>
                  <th>Cliente</th>
                  <th>Estatus</th>
                  <th>Monto Total</th>
                  <th>Estatus Contabilidad</th>
                  <th>Comisión (3% o 5%)</th>
                </tr>
                ${tablaDesgloseClientes}
              </table>
            </div>
          </body>
        </html>
      `);
        });
    });
});

const uploadGasto = upload.single('comprobante');
app.post('/guardar-gasto', verificarSesion, uploadGasto, (req, res) => {
    const u = req.session.usuario;
    const { concepto, categoria, monto } = req.body;
    let comp = null;
    if (req.file) {
        const ext = path.extname(req.file.originalname);
        comp = req.file.filename + ext;
        fs.renameSync(req.file.path, path.join(__dirname, 'uploads/gastos', comp));
    }
    db.run("INSERT INTO gastos_operativos (vendedor, concepto, categoria, monto, comprobante) VALUES (?, ?, ?, ?, ?)", [u.username, concepto, categoria, monto, comp], () => res.redirect('/gastos-ventas'));
});

app.listen(port, () => console.log(`ERP Industrial Tech corriendo en http://localhost:${port}`));