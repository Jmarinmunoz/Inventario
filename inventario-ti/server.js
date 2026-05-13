const express = require("express")
const path = require("path")
const cors = require("cors")
const session = require("express-session")
const ExcelJS = require("exceljs")

// ============================
// HELPERS
// ============================

function displayCellOrDash(v) {
  if (v == null || String(v).trim() === "") return "-"
  return String(v).trim()
}

function normalizeEstadoDisplay(s) {
  if (s == null || String(s).trim() === "") return "-"
  const t = String(s).trim()
  const lower = t.toLowerCase()
  if (lower === "activo") return "Activo"
  if (lower === "inactivo") return "Inactivo"
  if (lower === "eliminado") return "De baja"
  return t.charAt(0).toUpperCase() + t.slice(1).toLowerCase()
}

const FECHA_DISPLAY_TZ = process.env.INVENTARIO_TZ || "America/Santiago"

function formatFechaLocal(val) {
  if (val == null || String(val).trim() === "") return "-"
  const raw = String(val).trim()
  if (raw === "-") return "-"
  let iso = raw.includes("T") ? raw : raw.replace(" ", "T")
  if (!/[zZ]$|[+-][0-9]{2}:?[0-9]{2}$/.test(iso)) iso = `${iso}Z`
  const d = new Date(iso)
  if (Number.isNaN(d.getTime())) return raw
  try {
    const s = new Intl.DateTimeFormat("sv-SE", {
      timeZone: FECHA_DISPLAY_TZ,
      year: "numeric", month: "2-digit", day: "2-digit",
      hour: "2-digit", minute: "2-digit", hour12: false
    }).format(d)
    const [datePart, timePart = ""] = s.split(" ")
    const [y, m, dd] = datePart.split("-")
    const hm = timePart ? timePart.split(":").slice(0, 2).join(":") : ""
    return hm ? `${dd}-${m}-${y} ${hm}` : `${dd}-${m}-${y}`
  } catch {
    const dd = String(d.getUTCDate()).padStart(2, "0")
    const mm = String(d.getUTCMonth() + 1).padStart(2, "0")
    const yyyy = d.getUTCFullYear()
    const hh = String(d.getUTCHours()).padStart(2, "0")
    const min = String(d.getUTCMinutes()).padStart(2, "0")
    return `${dd}-${mm}-${yyyy} ${hh}:${min}`
  }
}

function formatFechaCreacion(val) { return formatFechaLocal(val) }

function anioCalendarioActualZona() {
  try {
    const parts = new Intl.DateTimeFormat("en-CA", {
      timeZone: FECHA_DISPLAY_TZ, year: "numeric"
    }).formatToParts(new Date())
    const y = parts.find((p) => p.type === "year")
    if (y && y.value) return Number(y.value)
  } catch {}
  return new Date().getFullYear()
}

// ---- Excel ----

function thinBorder() {
  const c = { argb: "FFB4B4B4" }
  return {
    top: { style: "thin", color: c }, left: { style: "thin", color: c },
    bottom: { style: "thin", color: c }, right: { style: "thin", color: c }
  }
}

function excelColLetter(n) {
  let s = "", x = n
  while (x > 0) {
    const m = (x - 1) % 26
    s = String.fromCharCode(65 + m) + s
    x = Math.floor((x - 1) / 26)
  }
  return s
}

const EXPORT_COLS = [
  { key: "codigo", header: "Código", center: true },
  { key: "departamento", header: "Área" },
  { key: "marca", header: "Marca" },
  { key: "modelo", header: "Modelo" },
  { key: "tipo", header: "Tipo" },
  { key: "licencias", header: "Licencias" },
  { key: "anio_compra", header: "Año compra", center: true },
  { key: "sistema_operativo", header: "Sistema operativo" },
  { key: "usuario", header: "Usuario" },
  { key: "estado", header: "Estado", center: true },
  { key: "motivo_baja", header: "Motivo baja" },
  { key: "comentario_inactivo", header: "Comentario (inactivo)" },
  { key: "fecha_creacion", header: "Fecha creación", center: true }
]

async function buildInventarioExcel(rows) {
  const workbook = new ExcelJS.Workbook()
  workbook.creator = "Inventario TI"
  workbook.created = new Date()
  const sheet = workbook.addWorksheet("Inventario", {
    views: [{ state: "frozen", ySplit: 1, activeCell: "A2" }]
  })
  const headerRow = sheet.addRow(EXPORT_COLS.map((c) => c.header))
  headerRow.height = 24
  headerRow.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: "FFFFFFFF" }, size: 11 }
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1F4E79" } }
    cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true }
    cell.border = thinBorder()
  })
  const effectiveRows = rows.length > 0 ? rows : [null]
  let activos = 0, inactivos = 0
  for (const row of rows) {
    const e = normalizeEstadoDisplay(row.estado)
    if (e === "Activo") activos++
    else if (e === "Inactivo") inactivos++
  }
  let dataRowIndex = 0
  for (const row of effectiveRows) {
    const values = EXPORT_COLS.map((col) => {
      if (row == null) return "-"
      if (col.key === "fecha_creacion") return formatFechaCreacion(row.fecha_creacion)
      if (col.key === "estado") return normalizeEstadoDisplay(row.estado)
      if (col.key === "codigo") return row.codigo != null ? String(row.codigo) : "-"
      if (col.key === "departamento") return displayCellOrDash(nombreAreaMostrar(row.departamento))
      return displayCellOrDash(row[col.key])
    })
    const r = sheet.addRow(values)
    r.height = 20
    const stripeRow = dataRowIndex % 2 === 1
    dataRowIndex++
    r.eachCell((cell, colNumber) => {
      const col = EXPORT_COLS[colNumber - 1]
      const horiz = col.center ? "center" : col.key === "codigo" ? "center" : "left"
      cell.border = thinBorder()
      cell.font = { size: 11 }
      cell.alignment = { vertical: "middle", horizontal: horiz, wrapText: true }
      if (stripeRow) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFDCE6F1" } }
    })
  }
  const lastDataRow = 1 + effectiveRows.length
  const lastCol = excelColLetter(EXPORT_COLS.length)
  sheet.autoFilter = `A1:${lastCol}${lastDataRow}`
  for (let r = 2; r <= lastDataRow; r++) {
    const row = effectiveRows[r - 2]
    if (row == null) continue
    EXPORT_COLS.forEach((col, ci) => {
      const cell = sheet.getCell(r, ci + 1)
      if (col.key === "estado") {
        const v = normalizeEstadoDisplay(row.estado)
        if (v === "Activo") {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFC6EFCE" } }
          cell.font = { size: 11, color: { argb: "FF006100" } }
        } else if (v === "Inactivo") {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFC7CE" } }
          cell.font = { size: 11, color: { argb: "FF9C0006" } }
        } else if (v === "De baja") {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE8E8E8" } }
          cell.font = { size: 11, color: { argb: "FF444444" } }
        }
      }
      if (col.key === "comentario_inactivo") {
        const v = displayCellOrDash(row.comentario_inactivo)
        if (v !== "-") {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFE699" } }
          cell.font = { size: 11, bold: true, color: { argb: "FF7F1D00" } }
        }
      }
    })
  }
  const totales = ["TOTALES", `Total equipos: ${rows.length}`, `Activos: ${activos}`, `Inactivos: ${inactivos}`, ...Array(EXPORT_COLS.length - 4).fill("-")]
  const totalRow = sheet.addRow(totales)
  totalRow.height = 22
  totalRow.eachCell((cell, colNumber) => {
    cell.border = thinBorder()
    cell.alignment = { vertical: "middle", horizontal: colNumber <= 4 ? "left" : "center", wrapText: false }
    cell.font = colNumber === 1 ? { bold: true, size: 11, color: { argb: "FF1F4E79" } } : { size: 11 }
    if (colNumber > 1 && colNumber <= 4) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF2F2F2" } }
  })
  const maxRow = sheet.rowCount
  for (let c = 1; c <= EXPORT_COLS.length; c++) {
    let maxW = EXPORT_COLS[c - 1].header.length + 2
    for (let r = 1; r <= maxRow; r++) {
      const v = sheet.getCell(r, c).value != null ? String(sheet.getCell(r, c).value) : ""
      maxW = Math.max(maxW, Math.min(v.length + 2, 60))
    }
    sheet.getColumn(c).width = Math.min(maxW, 48)
  }
  return workbook.xlsx.writeBuffer()
}

// ---- HTML helpers ----

function escapeHtml(value) {
  if (value == null) return ""
  return String(value).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;")
}

const PREFIJO_CODIGO_UNICO = "COD"

function codigoEquipoMostrar(row) {
  if (!row) return ""
  return row.codigo != null ? String(row.codigo).trim() : ""
}

const NOMBRE_AREA_POR_CODIGO = {
  TI: "TI", BOD: "Bodega", FRI: "Frigorifico", RH: "RRHH",
  GER: "Gerencia", ADM: "Administracion", PAC: "Packing", MAN: "Mantencion"
}

function nombreAreaMostrar(codigoDepartamento) {
  if (codigoDepartamento == null) return ""
  const t = String(codigoDepartamento).trim()
  if (t === "" || t === "-") return t
  return NOMBRE_AREA_POR_CODIGO[t.toUpperCase()] || t
}

function normalizeTextoBusqueda(s) {
  return String(s || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase()
}

const CODIGOS_DEPARTAMENTO_MINUS = new Set(Object.keys(NOMBRE_AREA_POR_CODIGO).map((c) => c.toLowerCase()))

function codigosDepartamentoQueCoincidenConBusqueda(texto) {
  if (!texto || !texto.trim()) return []
  const q = normalizeTextoBusqueda(texto.trim())
  if (!q) return []
  if (CODIGOS_DEPARTAMENTO_MINUS.has(q) && q !== "ti") return []
  const out = []
  for (const [code, name] of Object.entries(NOMBRE_AREA_POR_CODIGO)) {
    if (normalizeTextoBusqueda(name).includes(q)) out.push(code)
  }
  return out
}

// ============================
// MOCK DATA (DEMO — sin persistencia)
// ============================

const MOCK_USUARIOS = [
  { id: 1, nombre: "admin", password: "admin123", activo: 1, es_superadmin: 1 },
  { id: 2, nombre: "demo",  password: "demo123",  activo: 1, es_superadmin: 0 }
]

let mockEquipos = [
  { codigo:"COD-0001", departamento:"TI",  marca:"Dell",      modelo:"Latitude 5520",       tipo:"Notebooks",  licencias:"Licencia Office standard", usuario:"Carlos Ramirez",  desperfecto:null, estado:"Activo",    motivo_baja:null, comentario_inactivo:null,                                    anio_compra:"2022", sistema_operativo:"Windows 11 Pro",  fecha_creacion:"2022-03-15 10:30:00" },
  { codigo:"COD-0002", departamento:"ADM", marca:"HP",        modelo:"ProBook 450 G8",      tipo:"Notebooks",  licencias:"Licencia Office standard", usuario:"Maria Gonzalez",  desperfecto:null, estado:"Activo",    motivo_baja:null, comentario_inactivo:null,                                    anio_compra:"2021", sistema_operativo:"Windows 10 Pro",  fecha_creacion:"2021-08-20 09:15:00" },
  { codigo:"COD-0003", departamento:"BOD", marca:"Canon",     modelo:"PIXMA G3160",         tipo:"Impresora",  licencias:"No",                       usuario:"Bodega General",  desperfecto:null, estado:"Activo",    motivo_baja:null, comentario_inactivo:null,                                    anio_compra:"2023", sistema_operativo:null,              fecha_creacion:"2023-01-10 14:00:00" },
  { codigo:"COD-0004", departamento:"GER", marca:"Dell",      modelo:"OptiPlex 7080",       tipo:"Escritorio", licencias:"Licencia Office standard", usuario:"Roberto Silva",   desperfecto:null, estado:"Activo",    motivo_baja:null, comentario_inactivo:null,                                    anio_compra:"2021", sistema_operativo:"Windows 11 Home", fecha_creacion:"2021-05-12 11:45:00" },
  { codigo:"COD-0005", departamento:"TI",  marca:"Tp Link",   modelo:"TL-SG108",            tipo:"Switch",     licencias:"No",                       usuario:"Red TI",          desperfecto:null, estado:"Activo",    motivo_baja:null, comentario_inactivo:null,                                    anio_compra:"2020", sistema_operativo:null,              fecha_creacion:"2020-11-03 08:30:00" },
  { codigo:"COD-0006", departamento:"ADM", marca:"HP",        modelo:"LaserJet Pro M404dn", tipo:"Impresora",  licencias:"No",                       usuario:"Administracion",  desperfecto:null, estado:"Activo",    motivo_baja:null, comentario_inactivo:null,                                    anio_compra:"2022", sistema_operativo:null,              fecha_creacion:"2022-06-01 09:00:00" },
  { codigo:"COD-0007", departamento:"RH",  marca:"Lenovo",    modelo:"ThinkPad E15",        tipo:"Notebooks",  licencias:"Licencia Office standard", usuario:"Ana Martinez",    desperfecto:null, estado:"Activo",    motivo_baja:null, comentario_inactivo:null,                                    anio_compra:"2021", sistema_operativo:"Windows 10 Pro",  fecha_creacion:"2021-09-28 13:20:00" },
  { codigo:"COD-0008", departamento:"TI",  marca:"Asus",      modelo:"RT-AX88U",            tipo:"Router",     licencias:"No",                       usuario:"Red TI",          desperfecto:null, estado:"Activo",    motivo_baja:null, comentario_inactivo:null,                                    anio_compra:"2023", sistema_operativo:null,              fecha_creacion:"2023-02-14 10:00:00" },
  { codigo:"COD-0009", departamento:"PAC", marca:"HP",        modelo:"Compaq Elite 8300",   tipo:"Escritorio", licencias:"No",                       usuario:"Juan Torres",     desperfecto:null, estado:"Inactivo",  motivo_baja:null, comentario_inactivo:"Requiere reemplazo de fuente de poder", anio_compra:"2019", sistema_operativo:"Windows 10",      fecha_creacion:"2019-07-22 15:30:00" },
  { codigo:"COD-0010", departamento:"FRI", marca:"Lenovo",    modelo:"IdeaPad 3",           tipo:"Notebooks",  licencias:"No",                       usuario:"Pedro Diaz",      desperfecto:null, estado:"Activo",    motivo_baja:null, comentario_inactivo:null,                                    anio_compra:"2022", sistema_operativo:"Windows 11 Home", fecha_creacion:"2022-10-05 08:00:00" },
  { codigo:"COD-0011", departamento:"BOD", marca:"Zebra",     modelo:"ZT230",               tipo:"Impresora",  licencias:"No",                       usuario:"Etiquetado",      desperfecto:null, estado:"Activo",    motivo_baja:null, comentario_inactivo:null,                                    anio_compra:"2021", sistema_operativo:null,              fecha_creacion:"2021-03-18 12:30:00" },
  { codigo:"COD-0012", departamento:"GER", marca:"Dell",      modelo:"Latitude 7400",       tipo:"Notebooks",  licencias:"Licencia Office standard", usuario:"Carmen Flores",   desperfecto:null, estado:"Activo",    motivo_baja:null, comentario_inactivo:null,                                    anio_compra:"2023", sistema_operativo:"Windows 11 Pro",  fecha_creacion:"2023-04-20 09:45:00" },
  { codigo:"COD-0013", departamento:"MAN", marca:"Hikvision", modelo:"DS-3E0105P",          tipo:"Switch",     licencias:"No",                       usuario:"Mantencion",      desperfecto:null, estado:"Activo",    motivo_baja:null, comentario_inactivo:null,                                    anio_compra:"2022", sistema_operativo:null,              fecha_creacion:"2022-08-15 11:00:00" },
  { codigo:"COD-0014", departamento:"PAC", marca:"Acer",      modelo:"Aspire 5 A515",       tipo:"Notebooks",  licencias:"No",                       usuario:"Luis Herrera",    desperfecto:null, estado:"Activo",    motivo_baja:null, comentario_inactivo:null,                                    anio_compra:"2021", sistema_operativo:"Windows 10 Pro",  fecha_creacion:"2021-11-30 14:15:00" },
  { codigo:"COD-0015", departamento:"TI",  marca:"HP",        modelo:"Elite 800 G9",        tipo:"Escritorio", licencias:"Licencia Office standard", usuario:"Soporte TI",      desperfecto:null, estado:"Eliminado", motivo_baja:"Daño irreparable en placa madre", comentario_inactivo:null, anio_compra:"2020", sistema_operativo:"Windows 11 Pro",  fecha_creacion:"2020-12-01 10:00:00" },
]

let mockHistorialEquipos = [
  { id:1, equipo_codigo:"COD-0001", usuario_anterior:"Jose Perez",  usuario_nuevo:"Carlos Ramirez",  estado_anterior:null,     estado_nuevo:null,       comentario:null,                                      tipo_cambio:"transferencia", fecha_cambio:"2023-01-10 09:00:00" },
  { id:2, equipo_codigo:"COD-0001", usuario_anterior:null,          usuario_nuevo:null,              estado_anterior:null,     estado_nuevo:null,        comentario:"Área: GER → TI",                         tipo_cambio:"ubicacion",     fecha_cambio:"2023-01-10 09:05:00" },
  { id:3, equipo_codigo:"COD-0009", usuario_anterior:null,          usuario_nuevo:null,              estado_anterior:"Activo", estado_nuevo:"Inactivo",  comentario:"Requiere reemplazo de fuente de poder",   tipo_cambio:"estado",        fecha_cambio:"2023-06-15 11:30:00" },
  { id:4, equipo_codigo:"COD-0015", usuario_anterior:null,          usuario_nuevo:null,              estado_anterior:"Activo", estado_nuevo:"Eliminado", comentario:"Daño irreparable en placa madre",         tipo_cambio:"estado",        fecha_cambio:"2024-02-20 16:00:00" },
]

let mockNotas = [
  { id:1, equipo_codigo:"COD-0001", texto:"Se actualizó la RAM de 8GB a 16GB. Rendimiento mejorado.",                             fecha_creacion:"2023-03-10 14:00:00" },
  { id:2, equipo_codigo:"COD-0009", texto:"Pendiente solicitar cotización para nueva fuente de poder. Equipo sin uso hasta entonces.", fecha_creacion:"2023-06-20 09:00:00" },
  { id:3, equipo_codigo:"COD-0012", texto:"Instalado antivirus corporativo y configurado acceso VPN.",                             fecha_creacion:"2023-05-02 11:00:00" },
]

let nextCodigoN = 16
let nextHistorialEquiposId = 5
let nextNotaId = 4

const NOTA_MAX_LENGTH = 4000

// ============================
// EXPRESS SETUP
// ============================

const app = express()
const SESSION_SECRET = process.env.SESSION_SECRET || "inventario-ti-demo-secret-2024"

app.use(express.json())
app.use(express.urlencoded({ extended: true }))
app.use(cors({ origin: true, credentials: true }))
app.use(session({
  secret: SESSION_SECRET,
  resave: false,
  saveUninitialized: false,
  name: "inventario.sid",
  cookie: { httpOnly: true, sameSite: "lax", maxAge: 7 * 24 * 60 * 60 * 1000 }
}))

// ============================
// AUTH HELPERS
// ============================

function esSuperadminSesion(req) {
  return !!(req.session && req.session.esSuperadmin)
}

function quiereRespuestaJson(req) {
  const a = req.get("Accept") || ""
  return a.includes("application/json")
}

function requireAuth(req, res, next) {
  if (req.session && req.session.userId) return next()
  if (quiereRespuestaJson(req) || req.method !== "GET") {
    return res.status(401).json({ error: "Debe iniciar sesión", code: "AUTH_REQUIRED" })
  }
  return res.redirect("/login")
}

function requireSuperadmin(req, res, next) {
  if (!req.session || !req.session.userId) {
    return res.status(401).json({ error: "Debe iniciar sesión", code: "AUTH_REQUIRED" })
  }
  if (!esSuperadminSesion(req)) {
    return res.status(403).json({ error: "Solo el superadministrador puede realizar esta acción" })
  }
  next()
}

function codigoEquipoDesdeParams(req) {
  const raw = req.params.codigo
  if (raw == null) return null
  const s = String(raw).trim()
  return s.length > 0 ? s : null
}

function siguienteCodigoUnico() {
  const reCod = new RegExp(`^${PREFIJO_CODIGO_UNICO}-(\\d+)$`, "i")
  let maxN = 0
  for (const eq of mockEquipos) {
    const m = reCod.exec(String(eq.codigo || "").trim())
    if (m) maxN = Math.max(maxN, parseInt(m[1], 10))
  }
  const n = Math.max(maxN + 1, nextCodigoN)
  nextCodigoN = n + 1
  return `${PREFIJO_CODIGO_UNICO}-${String(n).padStart(4, "0")}`
}

function registrarHistorialEquipo(entry) {
  mockHistorialEquipos.push({
    id: nextHistorialEquiposId++,
    equipo_codigo: entry.equipo_codigo,
    usuario_anterior: entry.usuario_anterior || null,
    usuario_nuevo: entry.usuario_nuevo || null,
    estado_anterior: entry.estado_anterior || null,
    estado_nuevo: entry.estado_nuevo || null,
    comentario: entry.comentario || null,
    tipo_cambio: entry.tipo_cambio,
    fecha_cambio: new Date().toISOString().replace("T", " ").split(".")[0]
  })
}

// ============================
// AUTH ROUTES
// ============================

app.get("/login", (req, res) => {
  if (req.session && req.session.userId) return res.redirect("/")
  res.sendFile(path.join(__dirname, "public", "login.html"))
})

app.post("/login", (req, res) => {
  const nombreRaw = req.body && typeof req.body.nombre === "string" ? req.body.nombre : ""
  const passRaw   = req.body && typeof req.body.password === "string" ? req.body.password : ""
  const nombre = nombreRaw.trim()
  const password = passRaw
  if (!nombre || !password) return res.status(400).json({ error: "Indique nombre y contraseña" })

  const user = MOCK_USUARIOS.find((u) => u.nombre.toLowerCase() === nombre.toLowerCase())
  if (!user || user.password !== password) {
    return res.status(401).json({ error: "Credenciales incorrectas" })
  }
  if (user.activo !== 1) {
    return res.status(403).json({ error: "Usuario desactivado. Contacte al administrador." })
  }
  const esSuper = user.es_superadmin === 1
  req.session.userId = user.id
  req.session.nombre = user.nombre
  req.session.esSuperadmin = esSuper
  res.json({ ok: true, user: { nombre: user.nombre, esSuperadmin: esSuper } })
})

app.get("/api/session", (req, res) => {
  if (!req.session || !req.session.userId) {
    return res.status(401).json({ authenticated: false })
  }
  res.json({
    authenticated: true,
    user: { id: req.session.userId, nombre: req.session.nombre, esSuperadmin: esSuperadminSesion(req) }
  })
})

app.post("/logout", (req, res) => {
  req.session.destroy(() => {
    res.clearCookie("inventario.sid", { path: "/" })
    res.json({ ok: true })
  })
})

// ============================
// USER MANAGEMENT (superadmin)
// ============================

app.get("/api/usuarios", requireAuth, requireSuperadmin, (req, res) => {
  res.json(MOCK_USUARIOS.map(({ id, nombre, activo, es_superadmin }) => ({ id, nombre, activo, es_superadmin })))
})

app.post("/api/usuarios", requireAuth, requireSuperadmin, (req, res) => {
  const nombre    = (req.body && typeof req.body.nombre === "string" ? req.body.nombre : "").trim()
  const password  = req.body && typeof req.body.password === "string" ? req.body.password : ""
  const esSuperIn = req.body && req.body.es_superadmin
  const es_superadmin = (esSuperIn === true || esSuperIn === 1 || esSuperIn === "1") ? 1 : 0

  if (!nombre || !password) return res.status(400).json({ error: "Nombre y contraseña son obligatorios" })
  if (password.length < 4) return res.status(400).json({ error: "La contraseña debe tener al menos 4 caracteres" })
  if (MOCK_USUARIOS.find((u) => u.nombre.toLowerCase() === nombre.toLowerCase())) {
    return res.status(409).json({ error: "Ya existe un usuario con ese nombre" })
  }
  const newId = Math.max(...MOCK_USUARIOS.map((u) => u.id)) + 1
  MOCK_USUARIOS.push({ id: newId, nombre, password, activo: 1, es_superadmin })
  res.status(201).json({ ok: true, id: newId, nombre, activo: 1, es_superadmin })
})

app.patch("/api/usuarios/:id", requireAuth, requireSuperadmin, (req, res) => {
  const id = Number(req.params.id)
  if (!Number.isInteger(id) || id < 1) return res.status(400).json({ error: "Id inválido" })
  const idx = MOCK_USUARIOS.findIndex((u) => u.id === id)
  if (idx === -1) return res.status(404).json({ error: "Usuario no encontrado" })

  const target = MOCK_USUARIOS[idx]
  const nuevoNombre = (req.body && typeof req.body.nombre === "string" ? req.body.nombre.trim() : target.nombre) || target.nombre
  if (!nuevoNombre) return res.status(400).json({ error: "El nombre no puede estar vacío" })

  const activoIn = req.body && Object.prototype.hasOwnProperty.call(req.body, "activo") ? req.body.activo : null
  const nuevoActivo = activoIn !== null ? (activoIn === true || activoIn === 1 || activoIn === "1" ? 1 : 0) : target.activo
  if (nuevoActivo !== 1 && id === Number(req.session.userId)) {
    return res.status(400).json({ error: "No puede desactivar su propia cuenta" })
  }

  const esSuperIn = req.body && Object.prototype.hasOwnProperty.call(req.body, "es_superadmin") ? req.body.es_superadmin : null
  const nuevoEsSuper = esSuperIn !== null ? (esSuperIn === true || esSuperIn === 1 || esSuperIn === "1" ? 1 : 0) : target.es_superadmin

  const passRaw = req.body && typeof req.body.password === "string" ? req.body.password : ""
  const nuevoPassword = passRaw && passRaw.length >= 4 ? passRaw : target.password

  MOCK_USUARIOS[idx] = { ...target, nombre: nuevoNombre, activo: nuevoActivo, es_superadmin: nuevoEsSuper, password: nuevoPassword }
  res.json({ ok: true, id, nombre: nuevoNombre, activo: nuevoActivo, es_superadmin: nuevoEsSuper })
})

// ============================
// MAIN PAGES
// ============================

function sendIndexHtml(req, res) {
  res.setHeader("Cache-Control", "no-store, no-cache, must-revalidate, private")
  res.setHeader("Pragma", "no-cache")
  res.setHeader("Expires", "0")
  res.sendFile(path.join(__dirname, "public", "index.html"))
}

app.get("/", requireAuth, sendIndexHtml)
app.get("/index.html", requireAuth, sendIndexHtml)

app.use(express.static(path.join(__dirname, "public"), {
  index: false, etag: false, lastModified: false,
  setHeaders(res) {
    res.setHeader("Cache-Control", "no-store, no-cache, must-revalidate, private")
    res.setHeader("Pragma", "no-cache")
    res.setHeader("Expires", "0")
  }
}))

// ============================
// DASHBOARD
// ============================

function dashboardVistaEquiposConfig(tipo) {
  const map = {
    todos:     { titulo: "Total — todos los equipos",    descripcion: "Incluye activos, inactivos y dados de baja.", filter: () => true },
    activos:   { titulo: "Activos",                       descripcion: "Equipos actualmente marcados como activos.",  filter: (e) => e.estado === "Activo" },
    inactivos: { titulo: "Inactivos",                     descripcion: "Equipos inactivos (requieren motivo al pasar a inactivo).", filter: (e) => e.estado === "Inactivo" }
  }
  return map[tipo] || null
}

app.get("/dashboard", requireAuth, (req, res) => {
  const total    = mockEquipos.length
  const activos  = mockEquipos.filter((e) => e.estado === "Activo").length
  const inactivos= mockEquipos.filter((e) => e.estado === "Inactivo").length
  const eliminados = mockEquipos.filter((e) => e.estado === "Eliminado").length

  res.send(`<style>
body{font-family:Arial;background:#f4f6f9;margin:40px;}
.grid{display:grid;grid-template-columns:repeat(4,1fr);gap:20px;}
.card{background:white;padding:25px;border-radius:10px;box-shadow:0 4px 10px rgba(0,0,0,0.1);}
.card-link{text-decoration:none;color:inherit;display:block;}
.card-link .card{transition:transform .15s ease, box-shadow .15s ease;}
.card-link:hover .card{transform:translateY(-2px);box-shadow:0 8px 16px rgba(0,0,0,0.12);}
.numero{font-size:35px;font-weight:bold;color:#0B47BF;}
.btn{background:#262626;color:#fff;border:none;padding:10px 14px;border-radius:6px;cursor:pointer;font-size:15px;font-family:inherit;}
</style>
<h1>Dashboard TI</h1>
<p style="color:#666;font-size:14px;margin:0 0 20px;">Haz clic en cada tarjeta para ver el detalle de esos equipos.</p>
<div class="grid">
<a class="card-link" href="/dashboard/vista/todos" title="Ver todos los equipos registrados">
<div class="card"><b>Total</b><div class="numero">${total}</div><div style="font-size:12px;color:#666;margin-top:6px;">Ver listado completo</div></div>
</a>
<a class="card-link" href="/dashboard/vista/activos" title="Ver equipos activos">
<div class="card"><b>Activos</b><div class="numero">${activos}</div><div style="font-size:12px;color:#666;margin-top:6px;">Ver equipos en uso</div></div>
</a>
<a class="card-link" href="/dashboard/vista/inactivos" title="Ver equipos inactivos">
<div class="card"><b>Inactivos</b><div class="numero">${inactivos}</div><div style="font-size:12px;color:#666;margin-top:6px;">Ver equipos inactivos</div></div>
</a>
<a class="card-link" href="/historial-bajas" title="Ver equipos dados de baja">
<div class="card"><b>Dados de baja</b><div class="numero">${eliminados}</div><div style="font-size:12px;color:#666;margin-top:6px;">Ver historial de bajas</div></div>
</a>
</div>
<br>
<button type="button" class="btn" onclick="window.location.href='/'">📋 Volver al inventario</button>`)
})

app.get("/dashboard/vista/:tipo", requireAuth, (req, res) => {
  const cfg = dashboardVistaEquiposConfig(req.params.tipo)
  if (!cfg) return res.redirect("/dashboard")
  const list = mockEquipos.filter(cfg.filter)
  const filas = list.map((r) => `
<tr>
<td>${escapeHtml(codigoEquipoMostrar(r))}</td>
<td>${escapeHtml(r.marca || "")}</td>
<td>${escapeHtml(r.modelo || "")}</td>
<td>${escapeHtml(r.tipo || "")}</td>
<td>${escapeHtml(nombreAreaMostrar(r.departamento) || "-")}</td>
<td>${escapeHtml(r.usuario || "")}</td>
<td>${escapeHtml(normalizeEstadoDisplay(r.estado))}</td>
<td><a href="/equipo/${encodeURIComponent(r.codigo)}">Ver detalle</a></td>
</tr>`).join("")

  res.send(`<html><head><meta charset="UTF-8">
<style>
body{font-family:Arial;background:#f4f6f9;margin:0;padding:30px;}
.card{background:#fff;border-radius:12px;padding:22px;box-shadow:0 4px 12px rgba(0,0,0,.08);}
table{width:100%;border-collapse:collapse;margin-top:14px;}
th,td{padding:10px;border-bottom:1px solid #e5e7eb;text-align:left;vertical-align:top;}
th{background:#0B47BF;color:#fff;}
a{color:#0B47BF;text-decoration:none;}
.acciones{margin-top:16px;display:flex;gap:10px;flex-wrap:wrap;}
.btn{background:#262626;color:#fff;border:none;padding:10px 14px;border-radius:6px;cursor:pointer;}
.sub{color:#666;font-size:14px;margin:6px 0 0;}
</style>
</head><body>
<div class="card">
<h2 style="margin:0 0 6px;">${escapeHtml(cfg.titulo)}</h2>
<p class="sub">${escapeHtml(cfg.descripcion)}</p>
<div class="sub"><b>Registros:</b> ${list.length}</div>
<table>
<thead><tr><th>Código único</th><th>Marca</th><th>Modelo</th><th>Tipo</th><th>Área</th><th>Usuario</th><th>Estado</th><th>Detalle</th></tr></thead>
<tbody>${filas || '<tr><td colspan="8">No hay equipos en esta categoría.</td></tr>'}</tbody>
</table>
<div class="acciones">
<button class="btn" type="button" onclick="window.location.href='/dashboard'">⬅️ Volver al dashboard</button>
<button class="btn" type="button" onclick="window.location.href='/'">📋 Volver al inventario</button>
</div>
</div></body></html>`)
})

app.get("/historial-bajas", requireAuth, (req, res) => {
  const rows = mockEquipos.filter((e) => e.estado === "Eliminado")
  const filas = rows.map((r) => `
<tr>
<td>${escapeHtml(codigoEquipoMostrar(r))}</td>
<td>${escapeHtml(r.marca || "")}</td>
<td>${escapeHtml(r.modelo || "")}</td>
<td>${escapeHtml(r.tipo || "")}</td>
<td>${escapeHtml(nombreAreaMostrar(r.departamento) || "-")}</td>
<td>${escapeHtml(r.usuario || "")}</td>
<td>${escapeHtml(r.motivo_baja || "-")}</td>
<td><a href="/equipo/${encodeURIComponent(r.codigo)}">Ver detalle</a></td>
</tr>`).join("")

  res.send(`<html><head><meta charset="UTF-8">
<style>
body{font-family:Arial;background:#f4f6f9;margin:0;padding:30px;}
.card{background:#fff;border-radius:12px;padding:20px;box-shadow:0 4px 12px rgba(0,0,0,.08);}
table{width:100%;border-collapse:collapse;margin-top:12px;}
th,td{padding:10px;border-bottom:1px solid #e5e7eb;text-align:left;vertical-align:top;}
th{background:#0B47BF;color:#fff;}
a{color:#0B47BF;text-decoration:none;}
.acciones{margin-top:14px;display:flex;gap:10px;}
.btn{background:#262626;color:#fff;border:none;padding:10px 14px;border-radius:6px;cursor:pointer;}
</style>
</head><body>
<div class="card">
<h2 style="margin:0 0 8px;">Historial de equipos dados de baja</h2>
<div style="color:#666;font-size:14px;">Total: ${rows.length}</div>
<table>
<thead><tr><th>Código único</th><th>Marca</th><th>Modelo</th><th>Tipo</th><th>Área</th><th>Usuario</th><th>Motivo baja</th><th>Detalle</th></tr></thead>
<tbody>${filas || '<tr><td colspan="8">No hay equipos dados de baja.</td></tr>'}</tbody>
</table>
<div class="acciones">
<button class="btn" type="button" onclick="window.location.href='/dashboard'">⬅️ Volver al dashboard</button>
<button class="btn" type="button" onclick="window.location.href='/'">📋 Volver al inventario</button>
</div>
</div></body></html>`)
})

// ============================
// CREAR EQUIPO
// ============================

const DEPARTAMENTOS_VALIDOS = ["TI","BOD","FRI","RH","GER","ADM","PAC","MAN"]
const TIPOS_VALIDOS = ["Antena","Router","Repetidores","Switch","Notebook","Notebooks","Escritorio","Impresora"]
const TIPOS_CON_SISTEMA_OPERATIVO = ["Notebooks","Escritorio"]
const TIPOS_CON_LICENCIAS = ["Notebooks","Escritorio"]
const SISTEMAS_OPERATIVOS_VALIDOS = ["Windows 7","Windows 10","Windows 10 Pro","Windows 11 Home","Windows 11 Pro","Linux"]
const LICENCIAS_VALIDAS = ["No","Licencia Office standard"]
const ESTADOS_VALIDOS = ["Activo","Inactivo"]

app.post("/equipos", requireAuth, (req, res) => {
  const marca       = (req.body && typeof req.body.marca === "string"       ? req.body.marca       : "").trim()
  const modelo      = (req.body && typeof req.body.modelo === "string"      ? req.body.modelo      : "").trim()
  const tipoRaw     = (req.body && typeof req.body.tipo === "string"        ? req.body.tipo        : "").trim()
  const licencias   = (req.body && typeof req.body.licencias === "string"   ? req.body.licencias   : "").trim()
  const usuario     = (req.body && typeof req.body.usuario === "string"     ? req.body.usuario     : "").trim()
  const estado      = (req.body && typeof req.body.estado === "string"      ? req.body.estado      : "").trim()
  const departamento= (req.body && typeof req.body.departamento === "string"? req.body.departamento: "").trim().toUpperCase()
  const comentario  = (req.body && typeof req.body.comentario === "string"  ? req.body.comentario  : "").trim()
  const anioCompra  = (req.body && typeof req.body.anio_compra === "string" ? req.body.anio_compra : "").trim()
  const sistemaOpCliente = (req.body && typeof req.body.sistema_operativo === "string" ? req.body.sistema_operativo : "").trim()
  const tipoCanonico = tipoRaw === "Notebook" ? "Notebooks" : tipoRaw

  if (!marca || !modelo || !tipoCanonico || !licencias || !usuario || !estado || !anioCompra) {
    return res.status(400).json({ error: "Todos los campos del formulario son obligatorios" })
  }
  if (!/^\d{4}$/.test(anioCompra)) {
    return res.status(400).json({ error: "El año de compra debe tener exactamente 4 dígitos" })
  }
  if (Number(anioCompra) > anioCalendarioActualZona()) {
    return res.status(400).json({ error: `El año de compra no puede ser mayor a ${anioCalendarioActualZona()}` })
  }
  if (!DEPARTAMENTOS_VALIDOS.includes(departamento)) return res.status(400).json({ error: "Seleccione un área válida" })
  if (!TIPOS_VALIDOS.includes(tipoCanonico)) return res.status(400).json({ error: "Seleccione un tipo de equipo valido" })
  if (!ESTADOS_VALIDOS.includes(estado)) return res.status(400).json({ error: "Seleccione un estado valido" })
  if (TIPOS_CON_LICENCIAS.includes(tipoCanonico)) {
    if (!LICENCIAS_VALIDAS.includes(licencias)) return res.status(400).json({ error: "Seleccione una licencia valida" })
  } else if (licencias !== "No") {
    return res.status(400).json({ error: "La licencia solo aplica a Notebooks o Escritorio" })
  }

  let sistemaOperativoDb = null
  if (TIPOS_CON_SISTEMA_OPERATIVO.includes(tipoCanonico)) {
    if (!SISTEMAS_OPERATIVOS_VALIDOS.includes(sistemaOpCliente)) return res.status(400).json({ error: "Seleccione un sistema operativo válido" })
    sistemaOperativoDb = sistemaOpCliente
  } else if (sistemaOpCliente !== "") {
    return res.status(400).json({ error: "El sistema operativo solo aplica a Notebooks y Escritorio" })
  }
  if (estado === "Inactivo" && !comentario) {
    return res.status(400).json({ error: "Debe ingresar el motivo al marcar el equipo como Inactivo" })
  }

  const codigoUnico = siguienteCodigoUnico()
  const now = new Date().toISOString().replace("T", " ").split(".")[0]
  mockEquipos.unshift({
    codigo: codigoUnico,
    departamento,
    marca,
    modelo,
    tipo: tipoCanonico,
    licencias,
    usuario,
    desperfecto: null,
    estado,
    motivo_baja: null,
    comentario_inactivo: estado === "Inactivo" ? comentario : null,
    anio_compra: anioCompra,
    sistema_operativo: sistemaOperativoDb,
    fecha_creacion: now
  })

  res.status(201).json({ ok: true, codigo: codigoUnico, departamento, departamento_nombre: nombreAreaMostrar(departamento) })
})

// ============================
// LISTAR / BUSCAR EQUIPOS
// ============================

app.get("/equipos", requireAuth, (req, res) => {
  res.json(mockEquipos.filter((e) => e.estado !== "Eliminado"))
})

const ESTADOS_FILTRO = ["Todos", "Activo", "Inactivo"]

app.get("/buscar", requireAuth, (req, res) => {
  const rawQ = typeof req.query.q === "string" ? req.query.q.trim() : ""
  const rawEstado = typeof req.query.estado === "string" ? req.query.estado : "Todos"
  const estadoFiltro = ESTADOS_FILTRO.includes(rawEstado) ? rawEstado : "Todos"

  let result = mockEquipos.filter((e) => e.estado !== "Eliminado")

  if (estadoFiltro === "Activo" || estadoFiltro === "Inactivo") {
    result = result.filter((e) => e.estado === estadoFiltro)
  }

  if (rawQ !== "") {
    const q = normalizeTextoBusqueda(rawQ)
    const deptCodes = codigosDepartamentoQueCoincidenConBusqueda(rawQ)
    result = result.filter((e) => {
      const matchDept = deptCodes.length > 0 && deptCodes.includes((e.departamento || "").toUpperCase())
      return (
        matchDept ||
        normalizeTextoBusqueda(e.marca || "").includes(q) ||
        normalizeTextoBusqueda(e.modelo || "").includes(q) ||
        normalizeTextoBusqueda(e.usuario || "").includes(q) ||
        normalizeTextoBusqueda(e.tipo || "").includes(q) ||
        normalizeTextoBusqueda(e.codigo || "").includes(q) ||
        normalizeTextoBusqueda(e.anio_compra || "").includes(q) ||
        normalizeTextoBusqueda(e.sistema_operativo || "").includes(q)
      )
    })
  }

  res.json(result)
})

// ============================
// DETALLE EQUIPO
// ============================

app.get("/equipo/:codigo", requireAuth, (req, res) => {
  const codigo = codigoEquipoDesdeParams(req)
  if (!codigo) return res.status(400).send("Codigo de equipo invalido")
  const row = mockEquipos.find((e) => e.codigo === codigo)
  if (!row) return res.send("Equipo no encontrado")

  const motivoDisplay = row.estado === "Inactivo" ? "block" : "none"
  const urlEtiqueta = `/etiqueta/${encodeURIComponent(row.codigo)}?c=${encodeURIComponent(codigoEquipoMostrar(row))}`

  res.send(`<html>
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<script src="https://cdn.jsdelivr.net/npm/sweetalert2@11"></script>
<style>
body{font-family:Arial;background:#f4f6f9;margin:0;}
.container{padding:40px;display:flex;justify-content:center;}
.card{background:white;padding:30px;border-radius:12px;box-shadow:0 4px 15px rgba(0,0,0,0.1);width:100%;max-width:560px;animation:fadeIn 0.5s ease;box-sizing:border-box;}
h1{margin-bottom:8px;}
.subtitulo{color:#666;font-size:14px;margin:0 0 20px;}
.item{margin-bottom:10px;}
.codigo{font-size:22px;font-weight:bold;color:#0B47BF;margin-bottom:15px;}
.acciones{margin-top:20px;display:flex;flex-direction:column;gap:12px;}
button{background:#0B47BF;color:white;border:none;padding:10px 14px;cursor:pointer;border-radius:6px;font-size:15px;}
.btn-editar{background:#9DBF21;}
a.link-volver{text-decoration:none;color:#0B47BF;margin-top:8px;display:inline-block;}
.bloque-notas{margin-top:24px;padding-top:22px;border-top:1px solid #e5e7eb;}
.bloque-notas h2{margin:0 0 10px;font-size:17px;color:#333;}
.notas-accion{margin-bottom:12px;}
.btn-nota{background:#6c757d;width:100%;max-width:280px;}
.notas-lista{list-style:none;margin:0;padding:0;max-height:320px;overflow-y:auto;border:1px solid #e9ecef;border-radius:8px;background:#fafafa;}
.nota-item{padding:12px 14px;border-bottom:1px solid #e9ecef;}
.nota-item:last-child{border-bottom:none;}
.nota-fecha{display:block;font-size:12px;color:#666;margin-bottom:6px;}
.nota-texto{font-size:14px;color:#222;white-space:pre-wrap;word-break:break-word;}
.nota-vacio{padding:14px;color:#666;font-size:14px;text-align:center;}
@keyframes fadeIn{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:translateY(0)}}
</style>
</head>
<body>
<div class="container">
<div class="card">
<h1>Detalles del equipo</h1>
<div class="codigo">${escapeHtml(codigoEquipoMostrar(row))}</div>
<div class="item"><b>Marca:</b> ${escapeHtml(row.marca)}</div>
<div class="item"><b>Modelo:</b> ${escapeHtml(row.modelo)}</div>
<div class="item"><b>Área:</b> ${escapeHtml(nombreAreaMostrar(row.departamento))}</div>
<div class="item"><b>Tipo:</b> ${escapeHtml(row.tipo)}</div>
<div class="item"><b>Usuario:</b> <span id="usuarioTexto">${escapeHtml(row.usuario)}</span></div>
<div class="item"><b>Licencias:</b> ${escapeHtml(row.licencias)}</div>
<div class="item"><b>Año de compra:</b> ${escapeHtml(displayCellOrDash(row.anio_compra))}</div>
<div class="item"><b>Sistema operativo:</b> ${escapeHtml(displayCellOrDash(row.sistema_operativo))}</div>
<div class="item"><b>Estado:</b> ${escapeHtml(normalizeEstadoDisplay(row.estado))}</div>
<div class="item" style="display:${motivoDisplay}"><b>Motivo inactivo:</b> ${escapeHtml(row.comentario_inactivo || "")}</div>
<div class="item" style="display:${row.estado === "Eliminado" ? "block" : "none"}"><b>Motivo baja:</b> ${escapeHtml(row.motivo_baja || "")}</div>
<div class="acciones">
<button type="button" onclick='window.open(${JSON.stringify(urlEtiqueta)})'>🖨️ Generar etiqueta</button>
<a class="btn-editar" href="/equipo/${encodeURIComponent(row.codigo)}/editar" style="text-align:center;padding:10px 14px;border-radius:6px;color:white;text-decoration:none;font-size:15px;">✏️ Editar equipo</a>
<button type="button" onclick="window.location.href='/'">📋 Volver al inventario</button>
</div>
<div class="bloque-notas">
<h2>Notas y comentarios</h2>
<p style="margin:0 0 12px;font-size:13px;color:#666;">Registre cambios realizados al equipo u otras observaciones.</p>
<div class="notas-accion">
<button type="button" class="btn-nota" id="btnAgregarNota">📝 Agregar nota</button>
</div>
<ul class="notas-lista" id="listaNotas" aria-live="polite"></ul>
</div>
</div>
</div>
<script>
(function(){
var codigoEquipo=${JSON.stringify(row.codigo)};
var lista=document.getElementById("listaNotas");
function escapeText(s){return String(s==null?"":s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");}
function renderNotas(items){
if(!lista)return;
if(!items||!items.length){lista.innerHTML='<li class="nota-vacio">No hay notas registradas.</li>';return;}
lista.innerHTML=items.map(function(n){return'<li class="nota-item"><span class="nota-fecha">'+escapeText(n.fecha)+'</span><div class="nota-texto">'+escapeText(n.texto)+'</div></li>';}).join("");
}
function cargarNotas(){
fetch("/equipo/"+encodeURIComponent(codigoEquipo)+"/notas",{credentials:"include",headers:{Accept:"application/json"}})
.then(function(res){return res.ok?res.json():Promise.reject();})
.then(renderNotas)
.catch(function(){lista.innerHTML='<li class="nota-vacio">No se pudo cargar el historial de notas.</li>';});
}
var btn=document.getElementById("btnAgregarNota");
if(btn){btn.addEventListener("click",function(){
if(typeof Swal==="undefined"){alert("No se puede abrir el formulario de nota.");return;}
Swal.fire({title:"Agregar nota",input:"textarea",inputLabel:"Comentario o cambio realizado",inputPlaceholder:"Ej.: Se cambió disco SSD, reinstalación de Windows…",showCancelButton:true,confirmButtonText:"Guardar",cancelButtonText:"Cancelar",confirmButtonColor:"#0B47BF",cancelButtonColor:"#6c757d",inputAttributes:{"aria-label":"Texto de la nota"},inputValidator:function(v){var t=v!=null?String(v).trim():"";if(!t)return"Escriba un texto para la nota";if(t.length>${NOTA_MAX_LENGTH})return"El texto no puede superar ${NOTA_MAX_LENGTH} caracteres";}}).then(function(result){
if(!result.isConfirmed)return;
var texto=String(result.value).trim();
fetch("/equipo/"+encodeURIComponent(codigoEquipo)+"/notas",{method:"POST",credentials:"include",headers:{"Content-Type":"application/json",Accept:"application/json"},body:JSON.stringify({texto:texto})})
.then(function(res){return res.json().then(function(data){return{res:res,data:data};});})
.then(function(r){
if(!r.res.ok){Swal.fire({icon:"error",title:"Error",text:(r.data&&r.data.error)||"No se pudo guardar la nota",confirmButtonColor:"#0B47BF"});return;}
cargarNotas();
Swal.fire({icon:"success",title:"Nota guardada",timer:1600,showConfirmButton:false});
}).catch(function(){Swal.fire({icon:"error",title:"Error",text:"Sin conexión o servidor no disponible.",confirmButtonColor:"#0B47BF"});});
});});}
cargarNotas();
})();
</script>
</body>
</html>`)
})

// ============================
// NOTAS
// ============================

app.get("/equipo/:codigo/notas", requireAuth, (req, res) => {
  const codigo = codigoEquipoDesdeParams(req)
  if (!codigo) return res.status(400).json({ error: "Codigo de equipo invalido" })
  if (!mockEquipos.find((e) => e.codigo === codigo)) return res.status(404).json({ error: "Equipo no encontrado" })
  const list = mockNotas
    .filter((n) => n.equipo_codigo === codigo)
    .sort((a, b) => b.id - a.id)
    .map((n) => ({ id: n.id, texto: n.texto, fecha: formatFechaLocal(n.fecha_creacion) }))
  res.json(list)
})

app.post("/equipo/:codigo/notas", requireAuth, (req, res) => {
  const codigo = codigoEquipoDesdeParams(req)
  if (!codigo) return res.status(400).json({ error: "Codigo de equipo invalido" })
  if (!mockEquipos.find((e) => e.codigo === codigo)) return res.status(404).json({ error: "Equipo no encontrado" })
  const texto = (req.body && typeof req.body.texto === "string" ? req.body.texto : "").trim()
  if (!texto) return res.status(400).json({ error: "La nota no puede estar vacia" })
  if (texto.length > NOTA_MAX_LENGTH) return res.status(400).json({ error: `La nota no puede superar ${NOTA_MAX_LENGTH} caracteres` })
  const now = new Date().toISOString().replace("T", " ").split(".")[0]
  mockNotas.push({ id: nextNotaId++, equipo_codigo: codigo, texto, fecha_creacion: now })
  res.status(201).json({ ok: true, id: nextNotaId - 1, fecha: formatFechaLocal(now) })
})

// ============================
// EDITAR EQUIPO
// ============================

app.get("/equipo/:codigo/editar", requireAuth, (req, res) => {
  const codigo = codigoEquipoDesdeParams(req)
  if (!codigo) return res.status(400).send("Codigo de equipo invalido")
  const row = mockEquipos.find((e) => e.codigo === codigo)
  if (!row) return res.send("Equipo no encontrado")
  const urlEtiqueta = `/etiqueta/${encodeURIComponent(row.codigo)}?c=${encodeURIComponent(codigoEquipoMostrar(row))}`

  res.send(`<html>
<head>
<meta charset="UTF-8">
<script src="https://cdn.jsdelivr.net/npm/sweetalert2@11"></script>
<style>
body{font-family:Arial;background:#f4f6f9;margin:0;}
.container{padding:40px;display:flex;justify-content:center;}
.card{background:white;padding:30px;border-radius:12px;box-shadow:0 4px 15px rgba(0,0,0,0.1);width:500px;animation:fadeIn 0.5s ease;}
h1{margin-bottom:20px;}
.codigo{font-size:22px;font-weight:bold;color:#0B47BF;margin-bottom:15px;}
button{background:#9DBF21;color:white;border:none;padding:10px;margin-top:10px;cursor:pointer;border-radius:6px;}
@keyframes fadeIn{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:translateY(0)}}
.form-cambios{margin-top:18px;padding-top:18px;border-top:1px solid #eee;}
.form-cambios label{display:block;margin-bottom:8px;}
.form-cambios input,.form-cambios select{width:100%;padding:10px;border-radius:6px;border:1px solid #ccc;font-size:15px;margin-bottom:10px;}
.form-cambios textarea{width:100%;padding:10px;border-radius:6px;border:1px solid #ccc;font-size:15px;font-family:Arial,sans-serif;resize:vertical;min-height:72px;box-sizing:border-box;margin-bottom:10px;}
.form-cambios button[type="submit"]{width:100%;}
.msg-cambios{margin:10px 0 0;font-size:14px;min-height:20px;}
.historial-cambios{margin-top:14px;max-height:240px;overflow:auto;border:1px solid #e9ecef;border-radius:8px;padding:10px;background:#fafafa;}
.historial-tabs{display:flex;gap:8px;margin:10px 0 10px;}
.historial-tab{background:#e9ecef;color:#333;border:none;padding:8px 12px;border-radius:999px;cursor:pointer;font-size:13px;}
.historial-tab.activo{background:#0B47BF;color:#fff;}
.historial-panel{display:none;}
.historial-panel.activo{display:block;}
.historial-cambios ul{margin:0;padding-left:18px;}
.historial-cambios li{margin-bottom:8px;}
</style>
</head>
<body>
<div class="container">
<div class="card">
<h1>Editar equipo</h1>
<p style="color:#666;font-size:14px;margin:-8px 0 16px;">Modificar estado. Si pasa de Inactivo a Activo, indique qué se reparó.</p>
<form id="formCambios" class="form-cambios" autocomplete="off">
<label for="usuarioEquipo">Usuario asignado</label>
<input id="usuarioEquipo" name="usuario" type="text" value="${escapeHtml(row.usuario || "")}" required>
<label for="departamentoEquipo">Área</label>
<select id="departamentoEquipo" name="departamento" required aria-label="Área">
<option value="" ${row.departamento ? "" : "selected"}>Seleccionar</option>
<option value="TI"  ${row.departamento==="TI"  ? "selected":""}>TI</option>
<option value="BOD" ${row.departamento==="BOD" ? "selected":""}>Bodega</option>
<option value="FRI" ${row.departamento==="FRI" ? "selected":""}>Frigorifico</option>
<option value="RH"  ${row.departamento==="RH"  ? "selected":""}>RRHH</option>
<option value="GER" ${row.departamento==="GER" ? "selected":""}>Gerencia</option>
<option value="ADM" ${row.departamento==="ADM" ? "selected":""}>Administracion</option>
<option value="PAC" ${row.departamento==="PAC" ? "selected":""}>Packing</option>
<option value="MAN" ${row.departamento==="MAN" ? "selected":""}>Mantencion</option>
</select>
<p style="margin:0 0 10px;font-size:12px;color:#666;">El código del equipo no cambia; solo se actualizan los campos editables.</p>
<label for="estadoEquipo">Estado</label>
<select id="estadoEquipo" name="estado" aria-label="Estado del equipo">
<option value="Activo"    ${row.estado==="Activo"    ? "selected":""}>Activo</option>
<option value="Inactivo"  ${row.estado==="Inactivo"  ? "selected":""}>Inactivo</option>
<option value="Eliminado" ${row.estado==="Eliminado" ? "selected":""}>Dar de baja</option>
</select>
<div id="wrapComentarioInactivo" style="display:${row.estado==="Inactivo"?"block":"none"}">
<label for="comentarioCambio">Motivo de inactividad</label>
<textarea id="comentarioCambio" name="comentario" rows="3" placeholder="Obligatorio solo al pasar a Inactivo">${escapeHtml(row.comentario_inactivo||"")}</textarea>
</div>
<div id="wrapComentarioReparacion" style="display:none">
<label for="comentarioReparacion">¿Qué se reparó del equipo?</label>
<textarea id="comentarioReparacion" name="comentario_reactivacion" rows="3" placeholder="Obligatorio al pasar de Inactivo a Activo"></textarea>
</div>
<div id="wrapMotivoBaja" style="display:none">
<label for="motivoBaja">Motivo de baja</label>
<textarea id="motivoBaja" name="motivo_baja" rows="3" placeholder="Obligatorio para dar de baja el equipo"></textarea>
</div>
<button type="submit">Guardar cambios</button>
<p id="msgCambios" class="msg-cambios" role="status" aria-live="polite"></p>
</form>
<div class="historial-cambios">
<b>Historial de cambios</b>
<div class="historial-tabs">
<button type="button" id="tabTransferencias" class="historial-tab activo">Transferencias (<span id="countTransferencias">0</span>)</button>
<button type="button" id="tabEstados" class="historial-tab">Estados (<span id="countEstados">0</span>)</button>
<button type="button" id="tabUbicacion" class="historial-tab">Ubicación (<span id="countUbicacion">0</span>)</button>
</div>
<div id="panelTransferencias" class="historial-panel activo"><ul id="listaTransferencias"></ul></div>
<div id="panelEstados" class="historial-panel"><ul id="listaEstados"></ul></div>
<div id="panelUbicacion" class="historial-panel"><ul id="listaUbicacion"></ul></div>
</div>
<br>
<button type="button" onclick='window.open(${JSON.stringify(urlEtiqueta)})'>🖨️ Generar Etiqueta</button>
<br>
<button type="button" onclick="window.location.href='/equipo/${encodeURIComponent(row.codigo)}'">⬅️ Volver a lectura</button>
<br>
<button type="button" onclick="window.location.href='/'">📋 Volver al inventario</button>
</div>
</div>
<script>
(function(){
var codigoEquipo=${JSON.stringify(row.codigo)};
var form=document.getElementById("formCambios");
var msg=document.getElementById("msgCambios");
var usuarioInput=document.getElementById("usuarioEquipo");
var departamentoSel=document.getElementById("departamentoEquipo");
var sel=document.getElementById("estadoEquipo");
var wrap=document.getElementById("wrapComentarioInactivo");
var ta=document.getElementById("comentarioCambio");
var wrapRep=document.getElementById("wrapComentarioReparacion");
var taRep=document.getElementById("comentarioReparacion");
var wrapMotivoBaja=document.getElementById("wrapMotivoBaja");
var motivoBajaInput=document.getElementById("motivoBaja");
var tabTransferencias=document.getElementById("tabTransferencias");
var tabEstados=document.getElementById("tabEstados");
var tabUbicacion=document.getElementById("tabUbicacion");
var panelTransferencias=document.getElementById("panelTransferencias");
var panelEstados=document.getElementById("panelEstados");
var panelUbicacion=document.getElementById("panelUbicacion");
var listaTransferencias=document.getElementById("listaTransferencias");
var listaEstados=document.getElementById("listaEstados");
var listaUbicacion=document.getElementById("listaUbicacion");
var countTransferencias=document.getElementById("countTransferencias");
var countEstados=document.getElementById("countEstados");
var countUbicacion=document.getElementById("countUbicacion");
var estadoInicialEquipo=${JSON.stringify(row.estado||"Activo")};
var NOMBRE_AREA=${JSON.stringify(NOMBRE_AREA_POR_CODIGO)};
function nombreAreaFromCodigo(c){if(c==null)return"-";var t=String(c).trim();if(t===""||t==="-")return t;return NOMBRE_AREA[t.toUpperCase()]||t;}
function humanizarComentarioUbicacion(s){var t=String(s==null?"":s).trim();var m=/^Área:\s*(.+?)\s*→\s*(.+)$/.exec(t);if(!m)return String(s==null?"":s);return"Área: "+nombreAreaFromCodigo(m[1].trim())+" → "+nombreAreaFromCodigo(m[2].trim());}
function escapeText(s){return String(s==null?"":s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");}
function etiquetaEstadoMostrar(e){if(e==="Eliminado")return"De baja";return e==null?"":String(e);}
function activarTab(tab){
tabTransferencias.classList.toggle("activo",tab==="transferencias");
tabEstados.classList.toggle("activo",tab==="estados");
tabUbicacion.classList.toggle("activo",tab==="ubicacion");
panelTransferencias.classList.toggle("activo",tab==="transferencias");
panelEstados.classList.toggle("activo",tab==="estados");
panelUbicacion.classList.toggle("activo",tab==="ubicacion");
}
tabTransferencias.addEventListener("click",function(){activarTab("transferencias");});
tabEstados.addEventListener("click",function(){activarTab("estados");});
tabUbicacion.addEventListener("click",function(){activarTab("ubicacion");});
function renderTransferencia(item){var c=item.comentario?" | Comentario: "+escapeText(item.comentario):"";return"<li><b>"+escapeText(item.usuario_anterior||"Sin usuario")+"</b> → <b>"+escapeText(item.usuario_nuevo||"")+"</b> | "+escapeText(item.fecha_cambio)+c+"</li>";}
function renderEstado(item){var c="";if(item.comentario){var ct=escapeText(item.comentario);c=item.estado_nuevo==="Inactivo"?" | Motivo: "+ct:" | "+ct;}var ant=etiquetaEstadoMostrar(item.estado_anterior);var nuevo=etiquetaEstadoMostrar(item.estado_nuevo);return"<li><b>"+escapeText(ant)+"</b> → <b>"+escapeText(nuevo)+"</b> | "+escapeText(item.fecha_cambio)+c+"</li>";}
function renderUbicacion(item){var d=item.comentario?escapeText(humanizarComentarioUbicacion(item.comentario)):"";return"<li>"+d+" | "+escapeText(item.fecha_cambio)+"</li>";}
function cargarHistorialCambios(){
Promise.all([
fetch("/equipo/"+encodeURIComponent(codigoEquipo)+"/historial-cambios?tipo=transferencia",{credentials:"include",headers:{Accept:"application/json"}}).then(function(r){return r.ok?r.json():[];}),
fetch("/equipo/"+encodeURIComponent(codigoEquipo)+"/historial-cambios?tipo=estado",{credentials:"include",headers:{Accept:"application/json"}}).then(function(r){return r.ok?r.json():[];}),
fetch("/equipo/"+encodeURIComponent(codigoEquipo)+"/historial-cambios?tipo=ubicacion",{credentials:"include",headers:{Accept:"application/json"}}).then(function(r){return r.ok?r.json():[]; })
]).then(function(results){
var t=Array.isArray(results[0])?results[0]:[];
var e=Array.isArray(results[1])?results[1]:[];
var u=Array.isArray(results[2])?results[2]:[];
countTransferencias.textContent=String(t.length);
countEstados.textContent=String(e.length);
if(countUbicacion)countUbicacion.textContent=String(u.length);
listaTransferencias.innerHTML=t.length?t.map(renderTransferencia).join(""):"<li>Sin transferencias registradas.</li>";
listaEstados.innerHTML=e.length?e.map(renderEstado).join(""):"<li>Sin cambios de estado registrados.</li>";
if(listaUbicacion)listaUbicacion.innerHTML=u.length?u.map(renderUbicacion).join(""):"<li>Sin cambios de área registrados.</li>";
}).catch(function(){listaTransferencias.innerHTML="<li>No se pudo cargar el historial.</li>";listaEstados.innerHTML="<li>No se pudo cargar el historial.</li>";if(listaUbicacion)listaUbicacion.innerHTML="<li>No se pudo cargar el historial.</li>";});}
function syncComentarioUI(){
var esInactivo=sel.value==="Inactivo";
var esEliminado=sel.value==="Eliminado";
var esActivo=sel.value==="Activo";
wrap.style.display=esInactivo?"block":"none";
wrapMotivoBaja.style.display=esEliminado?"block":"none";
var mostrarRep=esActivo&&estadoInicialEquipo==="Inactivo";
if(wrapRep&&taRep){wrapRep.style.display=mostrarRep?"block":"none";if(!mostrarRep)taRep.value="";}
if(!esInactivo&&!esEliminado)ta.value="";
if(!esEliminado)motivoBajaInput.value="";
}
sel.addEventListener("change",function(){msg.textContent="";syncComentarioUI();});
syncComentarioUI();
form.addEventListener("submit",function(e){
e.preventDefault();
msg.textContent="";
var usuario=(usuarioInput.value||"").trim();
var estado=sel.value;
var comentario=(ta.value||"").trim();
var motivoBaja=(motivoBajaInput.value||"").trim();
var comentarioRep=taRep?(taRep.value||"").trim():"";
if(!usuario){msg.textContent="Ingrese el usuario asignado.";msg.style.color="#c0392b";if(typeof Swal!=="undefined")Swal.fire({icon:"warning",title:"Falta información",text:"Ingrese el usuario asignado.",confirmButtonColor:"#262626"});return;}
if(estado==="Inactivo"&&!comentario){msg.textContent="Ingrese el motivo para marcar el equipo como Inactivo.";msg.style.color="#c0392b";if(typeof Swal!=="undefined")Swal.fire({icon:"warning",title:"Motivo requerido",text:"Ingrese el motivo para marcar el equipo como Inactivo.",confirmButtonColor:"#262626"});return;}
if(estado==="Activo"&&estadoInicialEquipo==="Inactivo"&&!comentarioRep){msg.textContent="Indique qué se reparó del equipo para volver a Activo.";msg.style.color="#c0392b";if(typeof Swal!=="undefined")Swal.fire({icon:"warning",title:"Comentario requerido",text:"Describa qué se reparó del equipo al pasar de Inactivo a Activo.",confirmButtonColor:"#262626"});return;}
if(estado==="Eliminado"&&!motivoBaja){msg.textContent="Ingrese el motivo para dar de baja el equipo.";msg.style.color="#c0392b";if(typeof Swal!=="undefined")Swal.fire({icon:"warning",title:"Motivo requerido",text:"Ingrese el motivo para dar de baja el equipo.",confirmButtonColor:"#262626"});return;}
var doEnviar=function(){
var payload={usuario:usuario,estado:estado,comentario:estado==="Inactivo"?comentario:"",motivo_baja:estado==="Eliminado"?motivoBaja:"",comentario_reactivacion:estado==="Activo"&&estadoInicialEquipo==="Inactivo"?comentarioRep:"",departamento:departamentoSel?departamentoSel.value:""};
fetch("/equipo/"+encodeURIComponent(codigoEquipo)+"/cambios",{method:"PATCH",credentials:"include",headers:{"Content-Type":"application/json",Accept:"application/json"},body:JSON.stringify(payload)})
.then(function(res){return res.json().then(function(data){return{res:res,data:data};});})
.then(function(r){
if(!r.res.ok){var et=(r.data&&r.data.error)||"No se pudieron guardar los cambios";msg.textContent=et;msg.style.color="#c0392b";if(typeof Swal!=="undefined")Swal.fire({icon:"error",title:"No se pudo guardar",text:et,confirmButtonColor:"#262626"});return;}
var d=r.data;
estadoInicialEquipo=d.estado||estadoInicialEquipo;
if(taRep)taRep.value="";
syncComentarioUI();
msg.textContent=d.registros_historial>0?"Cambios guardados correctamente.":"No hubo cambios para guardar.";
msg.style.color="#1e7e34";
if(d.estado==="Eliminado"){if(typeof Swal!=="undefined"){Swal.fire({icon:"success",title:"Equipo dado de baja",text:"El equipo se registró correctamente.",confirmButtonColor:"#262626"}).then(function(){window.location.href="/";});}else{window.location.href="/";} return;}
if(typeof Swal!=="undefined"&&d.registros_historial>0)Swal.fire({icon:"success",title:"Listo",text:"Cambios guardados correctamente.",toast:true,position:"top-end",showConfirmButton:false,timer:2200,timerProgressBar:true});
cargarHistorialCambios();
});};
if(estado==="Eliminado"&&typeof Swal!=="undefined"){Swal.fire({title:"¿Está seguro?",text:"¿Está seguro que desea dar de baja este equipo?",icon:"warning",showCancelButton:true,confirmButtonColor:"#d33",cancelButtonColor:"#6c757d",confirmButtonText:"Sí, dar de baja",cancelButtonText:"Cancelar"}).then(function(result){if(result.isConfirmed)doEnviar();});}else{doEnviar();}
});
cargarHistorialCambios();
})();
</script>
</body>
</html>`)
})

// ============================
// HISTORIAL DE CAMBIOS
// ============================

app.get("/equipo/:codigo/historial-cambios", requireAuth, (req, res) => {
  const codigo = codigoEquipoDesdeParams(req)
  if (!codigo) return res.status(400).json({ error: "Codigo de equipo invalido" })
  if (!mockEquipos.find((e) => e.codigo === codigo)) return res.status(404).json({ error: "Equipo no encontrado" })

  const tipoRaw = typeof req.query.tipo === "string" ? req.query.tipo.trim() : ""
  const tipo = ["transferencia","estado","ubicacion"].includes(tipoRaw) ? tipoRaw : ""

  let list = mockHistorialEquipos.filter((h) => h.equipo_codigo === codigo)
  if (tipo) list = list.filter((h) => h.tipo_cambio === tipo)
  list = list.sort((a, b) => b.id - a.id)
  const out = list.map((r) => ({ ...r, fecha_cambio: formatFechaLocal(r.fecha_cambio) }))
  res.json(out)
})

// ============================
// GUARDAR CAMBIOS EQUIPO (PATCH principal)
// ============================

const DEPARTAMENTOS_EDICION = ["TI","BOD","FRI","RH","GER","ADM","PAC","MAN"]
const ESTADOS_EDICION = ["Activo","Inactivo","Eliminado"]

app.patch("/equipo/:codigo/cambios", requireAuth, (req, res) => {
  const codigo = codigoEquipoDesdeParams(req)
  if (!codigo) return res.status(400).json({ error: "Codigo de equipo invalido" })

  const usuarioNuevo      = (req.body && typeof req.body.usuario === "string"       ? req.body.usuario       : "").trim()
  const nuevoEstado       = (req.body && typeof req.body.estado === "string"        ? req.body.estado        : "").trim()
  const comentarioEstado  = (req.body && typeof req.body.comentario === "string"    ? req.body.comentario    : "").trim()
  const motivoBaja        = (req.body && typeof req.body.motivo_baja === "string"   ? req.body.motivo_baja   : "").trim()
  const comentarioReact   = (req.body && typeof req.body.comentario_reactivacion === "string" ? req.body.comentario_reactivacion : "").trim()
  const departamentoNuevo = (req.body && typeof req.body.departamento === "string"  ? req.body.departamento  : "").trim().toUpperCase()

  if (!usuarioNuevo) return res.status(400).json({ error: "Debe indicar el usuario del equipo" })
  if (!DEPARTAMENTOS_EDICION.includes(departamentoNuevo)) return res.status(400).json({ error: "Seleccione un área válida" })
  if (!ESTADOS_EDICION.includes(nuevoEstado)) return res.status(400).json({ error: "Estado debe ser Activo, Inactivo o Eliminado" })
  if (nuevoEstado === "Inactivo" && !comentarioEstado) return res.status(400).json({ error: "Debe ingresar el motivo al marcar el equipo como Inactivo" })
  if (nuevoEstado === "Eliminado" && !motivoBaja) return res.status(400).json({ error: "Debe ingresar el motivo para dar de baja el equipo" })

  const idx = mockEquipos.findIndex((e) => e.codigo === codigo)
  if (idx === -1) return res.status(404).json({ error: "Equipo no encontrado" })
  const row = mockEquipos[idx]

  const usuarioAnterior     = (row.usuario || "").trim()
  const estadoAnterior      = row.estado || "Activo"
  const departamentoAnterior= (row.departamento || "").trim().toUpperCase()

  if (estadoAnterior === "Inactivo" && nuevoEstado === "Activo" && !comentarioReact) {
    return res.status(400).json({ error: "Debe describir qué se reparó del equipo al volver a Activo" })
  }

  const cambioUsuario     = usuarioAnterior !== usuarioNuevo
  const cambioEstado      = estadoAnterior !== nuevoEstado
  const cambioDepartamento= departamentoAnterior !== departamentoNuevo

  if (!cambioUsuario && !cambioEstado && !cambioDepartamento) {
    return res.json({ ok: true, codigo, usuario: usuarioAnterior, estado: estadoAnterior, comentario_inactivo: row.comentario_inactivo || null, motivo_baja: row.motivo_baja || null, departamento: departamentoAnterior, departamento_nombre: nombreAreaMostrar(departamentoAnterior), registros_historial: 0 })
  }

  const comentarioDb = nuevoEstado === "Inactivo" ? comentarioEstado : null
  const motivoBajaDb = nuevoEstado === "Eliminado" ? motivoBaja : null

  mockEquipos[idx] = { ...row, usuario: usuarioNuevo, estado: nuevoEstado, comentario_inactivo: comentarioDb, motivo_baja: motivoBajaDb, departamento: departamentoNuevo }

  let registros = 0
  if (cambioDepartamento) {
    registrarHistorialEquipo({ equipo_codigo: codigo, comentario: `Área: ${nombreAreaMostrar(departamentoAnterior)||"-"} → ${nombreAreaMostrar(departamentoNuevo)}`, tipo_cambio: "ubicacion" })
    registros++
  }
  if (cambioUsuario) {
    registrarHistorialEquipo({ equipo_codigo: codigo, usuario_anterior: usuarioAnterior, usuario_nuevo: usuarioNuevo, tipo_cambio: "transferencia" })
    registros++
  }
  if (cambioEstado) {
    const comentHistorial = nuevoEstado === "Inactivo" ? (comentarioEstado || null) : nuevoEstado === "Eliminado" ? (motivoBaja || null) : (nuevoEstado === "Activo" && estadoAnterior === "Inactivo" ? `Reparación: ${comentarioReact}` : null)
    registrarHistorialEquipo({ equipo_codigo: codigo, estado_anterior: estadoAnterior, estado_nuevo: nuevoEstado, comentario: comentHistorial, tipo_cambio: "estado" })
    registros++
  }
  if (cambioEstado && nuevoEstado === "Eliminado") {
    mockNotas = mockNotas.filter((n) => n.equipo_codigo !== codigo)
  }

  res.json({ ok: true, codigo, usuario: usuarioNuevo, estado: nuevoEstado, comentario_inactivo: comentarioDb, motivo_baja: motivoBajaDb, departamento: departamentoNuevo, departamento_nombre: nombreAreaMostrar(departamentoNuevo), registros_historial: registros })
})

app.patch("/equipo/:codigo", requireAuth, (req, res) => {
  const codigo = codigoEquipoDesdeParams(req)
  if (!codigo) return res.status(400).json({ error: "Codigo de equipo invalido" })
  const nuevoEstado  = (req.body && typeof req.body.estado === "string" ? req.body.estado : "").trim()
  const comentario   = (req.body && typeof req.body.comentario === "string" ? req.body.comentario : "").trim()
  if (!["Activo","Inactivo"].includes(nuevoEstado)) return res.status(400).json({ error: "Estado debe ser Activo o Inactivo" })
  if (nuevoEstado === "Inactivo" && !comentario) return res.status(400).json({ error: "Debe ingresar el motivo al marcar el equipo como Inactivo" })
  const idx = mockEquipos.findIndex((e) => e.codigo === codigo)
  if (idx === -1) return res.status(404).json({ error: "Equipo no encontrado" })
  const estadoAnterior = mockEquipos[idx].estado || "Activo"
  const comentarioDb = nuevoEstado === "Inactivo" ? comentario : null
  mockEquipos[idx] = { ...mockEquipos[idx], estado: nuevoEstado, comentario_inactivo: comentarioDb }
  if (estadoAnterior !== nuevoEstado) {
    registrarHistorialEquipo({ equipo_codigo: codigo, estado_anterior: estadoAnterior, estado_nuevo: nuevoEstado, comentario: comentario || null, tipo_cambio: "estado" })
  }
  res.json({ ok: true, estado: nuevoEstado, comentario_inactivo: comentarioDb })
})

app.patch("/equipo/:codigo/transferir", requireAuth, (req, res) => {
  const codigo = codigoEquipoDesdeParams(req)
  if (!codigo) return res.status(400).json({ error: "Codigo de equipo invalido" })
  const usuarioNuevo = (req.body && typeof req.body.usuario_nuevo === "string" ? req.body.usuario_nuevo : "").trim()
  if (!usuarioNuevo) return res.status(400).json({ error: "Debe indicar el nuevo usuario" })
  const idx = mockEquipos.findIndex((e) => e.codigo === codigo)
  if (idx === -1) return res.status(404).json({ error: "Equipo no encontrado" })
  const usuarioAnterior = (mockEquipos[idx].usuario || "").trim()
  if (usuarioAnterior === usuarioNuevo) return res.status(400).json({ error: "El nuevo usuario debe ser distinto al usuario actual" })
  mockEquipos[idx] = { ...mockEquipos[idx], usuario: usuarioNuevo }
  registrarHistorialEquipo({ equipo_codigo: codigo, usuario_anterior: usuarioAnterior, usuario_nuevo: usuarioNuevo, tipo_cambio: "transferencia" })
  res.json({ ok: true, codigo, usuario_anterior: usuarioAnterior, usuario_nuevo: usuarioNuevo, comentario: null })
})

// ============================
// EXPORTAR EXCEL
// ============================

app.get("/exportar", requireAuth, (req, res) => {
  const rows = [...mockEquipos].sort((a, b) => new Date(b.fecha_creacion) - new Date(a.fecha_creacion))
  buildInventarioExcel(rows)
    .then((buffer) => {
      const now = new Date()
      const stamp = `${now.getFullYear()}${String(now.getMonth()+1).padStart(2,"0")}${String(now.getDate()).padStart(2,"0")}-${String(now.getHours()).padStart(2,"0")}${String(now.getMinutes()).padStart(2,"0")}${String(now.getSeconds()).padStart(2,"0")}`
      const filename = `inventario-ti-${stamp}.xlsx`
      res.setHeader("Content-Disposition", `attachment; filename="${filename}"`)
      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
      res.setHeader("Cache-Control", "no-store, no-cache, must-revalidate, private")
      res.setHeader("Pragma", "no-cache")
      res.setHeader("Expires", "0")
      res.send(Buffer.from(buffer))
    })
    .catch(() => res.status(500).send("Error al generar el archivo Excel"))
})

// ============================
// ETIQUETA (impresión)
// ============================

app.get("/etiqueta/:codigo", requireAuth, (req, res) => {
  const codigoParam = codigoEquipoDesdeParams(req)
  const codigoQuery = typeof req.query.c === "string" ? req.query.c.trim() : ""
  if (!codigoParam) return res.status(400).send("Codigo de equipo invalido")
  const row = mockEquipos.find((e) => e.codigo === codigoParam)
  if (!row) return res.send("Equipo no encontrado")
  const codigoMostrar = codigoEquipoMostrar(row) || codigoQuery || ""
  const codigoTxt = escapeHtml(codigoMostrar)

  res.send(`<html>
<head>
<style>
body{font-family:Arial,sans-serif;display:flex;justify-content:center;align-items:center;min-height:100vh;background:#f4f6f9;margin:0;}
.vista-etiqueta{display:flex;flex-direction:column;align-items:center;gap:12px;}
.etiqueta{width:50mm;height:25mm;border:2px solid #000;padding:2mm;box-sizing:border-box;background:white;text-align:center;display:flex;align-items:center;justify-content:center;overflow:visible;}
.codigo{font-size:12mm;font-weight:bold;line-height:1.1;color:#000;-webkit-print-color-adjust:exact;print-color-adjust:exact;word-break:break-all;}
.acciones{text-align:center;}
@media print{
@page{size:50mm 25mm;margin:0;}
html,body{width:50mm;height:25mm;margin:0;padding:0;background:#fff !important;}
body{display:block;}
.vista-etiqueta{display:block;width:50mm;height:25mm;margin:0;padding:0;gap:0;}
.etiqueta{width:50mm;height:25mm;border:none;padding:1mm;margin:0;box-sizing:border-box;}
.acciones{display:none !important;}
}
</style>
</head>
<body>
<div class="vista-etiqueta">
<div class="etiqueta"><div class="codigo">${codigoTxt || "—"}</div></div>
<div class="acciones"><button type="button" onclick="window.print()">🖨️ Imprimir</button></div>
</div>
</body>
</html>`)
})

// ============================
// EXPORT (Vercel serverless)
// ============================

module.exports = app

if (require.main === module) {
  const PORT = process.env.PORT || 3001
  app.listen(PORT, "0.0.0.0", () => {
    console.log(`\nServidor DEMO corriendo en http://localhost:${PORT}`)
    console.log("─────────────────────────────────────")
    console.log("Credenciales de acceso:")
    console.log("  Usuario: admin | Contraseña: admin123  (SuperAdmin)")
    console.log("  Usuario: demo  | Contraseña: demo123")
    console.log("─────────────────────────────────────")
    console.log("NOTA: Los datos son demo. Los cambios se pierden al reiniciar.\n")
  })
}
