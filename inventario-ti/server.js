const express = require("express")
const path = require("path")
const sqlite3 = require("sqlite3").verbose()
const cors = require("cors")
const session = require("express-session")
const bcrypt = require("bcryptjs")
const ExcelJS = require("exceljs")

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

/**
 * Zona horaria para mostrar fechas (SQLite CURRENT_TIMESTAMP se guarda en UTC).
 * Definir INVENTARIO_TZ en el entorno para otra zona, ej. America/Santiago
 */
const FECHA_DISPLAY_TZ = process.env.INVENTARIO_TZ || "America/Santiago"

function esErrorSqliteBusy(err) {
  const m = String((err && err.message) || "")
  return m.includes("SQLITE_BUSY") || m.includes("database is locked")
}

function responderBloqueoBd(res) {
  return res.status(503).json({
    error:
      "La base de datos está en uso por otra aplicación. Cierre esa app e intente nuevamente."
  })
}

/** Año calendario actual en FECHA_DISPLAY_TZ (para validar año de compra). */
function anioCalendarioActualZona() {
  try {
    const parts = new Intl.DateTimeFormat("en-CA", {
      timeZone: FECHA_DISPLAY_TZ,
      year: "numeric"
    }).formatToParts(new Date())
    const y = parts.find((p) => p.type === "year")
    if (y && y.value) return Number(y.value)
  } catch {
    // ignorar
  }
  return new Date().getFullYear()
}

/**
 * Convierte un valor fecha de SQLite (asumido UTC) a texto dd-mm-aaaa HH:mm en FECHA_DISPLAY_TZ.
 */
function formatFechaLocal(val) {
  if (val == null || String(val).trim() === "") return "-"
  const raw = String(val).trim()
  if (raw === "-") return "-"
  let iso = raw.includes("T") ? raw : raw.replace(" ", "T")
  if (!/[zZ]$|[+-][0-9]{2}:?[0-9]{2}$/.test(iso)) {
    iso = `${iso}Z`
  }
  const d = new Date(iso)
  if (Number.isNaN(d.getTime())) return raw
  try {
    const s = new Intl.DateTimeFormat("sv-SE", {
      timeZone: FECHA_DISPLAY_TZ,
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
      hour: "2-digit",
      minute: "2-digit",
      hour12: false
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

function formatFechaCreacion(val) {
  return formatFechaLocal(val)
}

function thinBorder() {
  const c = { argb: "FFB4B4B4" }
  return {
    top: { style: "thin", color: c },
    left: { style: "thin", color: c },
    bottom: { style: "thin", color: c },
    right: { style: "thin", color: c }
  }
}

function excelColLetter(n) {
  let s = ""
  let x = n
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
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FF1F4E79" }
    }
    cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true }
    cell.border = thinBorder()
  })

  const effectiveRows = rows.length > 0 ? rows : [null]
  let activos = 0
  let inactivos = 0
  for (const row of rows) {
    const e = normalizeEstadoDisplay(row.estado)
    if (e === "Activo") activos += 1
    else if (e === "Inactivo") inactivos += 1
  }

  let dataRowIndex = 0
  for (const row of effectiveRows) {
    const values = EXPORT_COLS.map((col) => {
      if (row == null) return "-"
      if (col.key === "fecha_creacion") return formatFechaCreacion(row.fecha_creacion)
      if (col.key === "estado") return normalizeEstadoDisplay(row.estado)
      if (col.key === "codigo") {
        return row.codigo != null ? String(row.codigo) : "-"
      }
      if (col.key === "departamento") {
        return displayCellOrDash(nombreAreaMostrar(row.departamento))
      }
      return displayCellOrDash(row[col.key])
    })
    const r = sheet.addRow(values)
    r.height = 20
    const stripeRow = dataRowIndex % 2 === 1
    dataRowIndex += 1
    r.eachCell((cell, colNumber) => {
      const col = EXPORT_COLS[colNumber - 1]
      const horiz = col.center ? "center" : col.key === "codigo" ? "center" : "left"
      cell.border = thinBorder()
      cell.font = { size: 11 }
      cell.alignment = { vertical: "middle", horizontal: horiz, wrapText: true }
      if (stripeRow) {
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFDCE6F1" }
        }
      }
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
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFC6EFCE" }
          }
          cell.font = { size: 11, color: { argb: "FF006100" } }
        } else if (v === "Inactivo") {
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFFFC7CE" }
          }
          cell.font = { size: 11, color: { argb: "FF9C0006" } }
        } else if (v === "De baja") {
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFE8E8E8" }
          }
          cell.font = { size: 11, color: { argb: "FF444444" } }
        }
      }
      if (col.key === "comentario_inactivo") {
        const v = displayCellOrDash(row.comentario_inactivo)
        if (v !== "-") {
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFFFE699" }
          }
          cell.font = { size: 11, bold: true, color: { argb: "FF7F1D00" } }
        }
      }
    })
  }

  const totales = [
    "TOTALES",
    `Total equipos: ${rows.length}`,
    `Activos: ${activos}`,
    `Inactivos: ${inactivos}`,
    ...Array(EXPORT_COLS.length - 4).fill("-")
  ]
  const totalRow = sheet.addRow(totales)
  totalRow.height = 22
  totalRow.eachCell((cell, colNumber) => {
    cell.border = thinBorder()
    cell.alignment = { vertical: "middle", horizontal: colNumber <= 4 ? "left" : "center", wrapText: false }
    cell.font =
      colNumber === 1
        ? { bold: true, size: 11, color: { argb: "FF1F4E79" } }
        : { size: 11 }
    if (colNumber > 1 && colNumber <= 4) {
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFF2F2F2" }
      }
    }
  })

  const maxRow = sheet.rowCount
  for (let c = 1; c <= EXPORT_COLS.length; c++) {
    let maxW = EXPORT_COLS[c - 1].header.length + 2
    for (let r = 1; r <= maxRow; r++) {
      const cell = sheet.getCell(r, c)
      const v = cell.value != null ? String(cell.value) : ""
      maxW = Math.max(maxW, Math.min(v.length + 2, 60))
    }
    sheet.getColumn(c).width = Math.min(maxW, 48)
  }

  return workbook.xlsx.writeBuffer()
}

function escapeHtml(value) {
  if (value == null) return ""
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
}

const PREFIJO_CODIGO_UNICO = "COD"

/** Identificador visible del equipo (PK). */
function codigoEquipoMostrar(row) {
  if (!row) return ""
  const c = row.codigo != null ? String(row.codigo).trim() : ""
  return c
}

/** Código guardado en BD → nombre de área (sin siglas entre paréntesis). */
const NOMBRE_AREA_POR_CODIGO = {
  TI: "TI",
  BOD: "Bodega",
  FRI: "Frigorifico",
  RH: "RRHH",
  GER: "Gerencia",
  ADM: "Administracion",
  PAC: "Packing",
  MAN: "Mantencion"
}

function nombreAreaMostrar(codigoDepartamento) {
  if (codigoDepartamento == null) return ""
  const t = String(codigoDepartamento).trim()
  if (t === "" || t === "-") return t
  const cod = t.toUpperCase()
  return NOMBRE_AREA_POR_CODIGO[cod] || t
}

function normalizeTextoBusqueda(s) {
  return String(s || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
}

const CODIGOS_DEPARTAMENTO_MINUS = new Set(
  Object.keys(NOMBRE_AREA_POR_CODIGO).map((c) => c.toLowerCase())
)

/** Códigos de departamento cuyo nombre de área contiene el texto (no por código solo salvo TI, que se muestra como "TI"). */
function codigosDepartamentoQueCoincidenConBusqueda(texto) {
  if (texto == null || String(texto).trim() === "") return []
  const q = normalizeTextoBusqueda(texto.trim())
  if (q === "") return []
  if (CODIGOS_DEPARTAMENTO_MINUS.has(q) && q !== "ti") {
    return []
  }
  const out = []
  for (const [code, name] of Object.entries(NOMBRE_AREA_POR_CODIGO)) {
    if (normalizeTextoBusqueda(name).includes(q)) {
      out.push(code)
    }
  }
  return out
}

const app = express()

const rawPortEnv = process.env.PORT || process.env.IISNODE_HTTPPORT
const portEnvEsSoloNumero =
  rawPortEnv == null ||
  rawPortEnv === "" ||
  typeof rawPortEnv === "number" ||
  (typeof rawPortEnv === "string" && /^\d+$/.test(String(rawPortEnv).trim()))
if (!portEnvEsSoloNumero || process.env.IISNODE_VERSION) {
  app.set("trust proxy", 1)
}

const SESSION_SECRET =
  process.env.SESSION_SECRET || "inventario-ti-dev-cambiar-en-produccion"

app.use(express.json())
app.use(express.urlencoded({ extended: true }))
app.use(
  cors({
    origin: true,
    credentials: true
  })
)
app.use(
  session({
    secret: SESSION_SECRET,
    resave: false,
    saveUninitialized: false,
    name: "inventario.sid",
    cookie: {
      httpOnly: true,
      sameSite: "lax",
      maxAge: 7 * 24 * 60 * 60 * 1000
    }
  })
)

/* ========================
   BASE DE DATOS
======================== */

const dbPath = path.join(__dirname, "data", "inventario.db")
const db = new sqlite3.Database(dbPath)

/** Parámetro de ruta :codigo del equipo (PK); Express ya decodifica el path. */
function codigoEquipoDesdeParams(req) {
  const raw = req.params.codigo
  if (raw == null) return null
  const s = String(raw).trim()
  return s.length > 0 ? s : null
}

function migrateRemovePlantasColumn(done) {
  db.all(`PRAGMA table_info(equipos)`, (infoErr, cols) => {
    if (infoErr) return done(infoErr)
    const hasPlantas = Array.isArray(cols) && cols.some((c) => c && c.name === "plantas")
    if (!hasPlantas) return done()

    db.run("BEGIN IMMEDIATE TRANSACTION", (beginErr) => {
      if (beginErr) return done(beginErr)

      const rollbackAndDone = (err) => {
        db.run("ROLLBACK", () => done(err))
      }

      db.run(
        `CREATE TABLE IF NOT EXISTS equipos_tmp_no_plantas (
id INTEGER PRIMARY KEY AUTOINCREMENT,
codigo TEXT,
departamento TEXT,
planta TEXT,
marca TEXT,
modelo TEXT,
tipo TEXT,
licencias TEXT,
usuario TEXT,
desperfecto TEXT,
estado TEXT,
motivo_baja TEXT,
comentario_inactivo TEXT,
anio_compra TEXT,
sistema_operativo TEXT,
fecha_creacion DATETIME DEFAULT CURRENT_TIMESTAMP
)`,
        (createErr) => {
          if (createErr) return rollbackAndDone(createErr)

          db.run(
            `INSERT INTO equipos_tmp_no_plantas
             (id,codigo,departamento,planta,marca,modelo,tipo,licencias,usuario,desperfecto,estado,motivo_baja,comentario_inactivo,anio_compra,sistema_operativo,fecha_creacion)
             SELECT id,COALESCE(NULLIF(TRIM(codigo),''),NULLIF(TRIM(codigo_unico),''),codigo),departamento,planta,marca,modelo,tipo,licencias,usuario,desperfecto,estado,motivo_baja,comentario_inactivo,anio_compra,sistema_operativo,fecha_creacion
             FROM equipos`,
            (copyErr) => {
              if (copyErr) return rollbackAndDone(copyErr)

              db.run(`DROP TABLE equipos`, (dropErr) => {
                if (dropErr) return rollbackAndDone(dropErr)

                db.run(`ALTER TABLE equipos_tmp_no_plantas RENAME TO equipos`, (renameErr) => {
                  if (renameErr) return rollbackAndDone(renameErr)

                  db.run(
                    `CREATE UNIQUE INDEX IF NOT EXISTS idx_equipos_codigo_unico ON equipos(codigo)`,
                    (indexErr) => {
                      if (indexErr) return rollbackAndDone(indexErr)
                      db.run("COMMIT", (commitErr) => done(commitErr || null))
                    }
                  )
                })
              })
            }
          )
        }
      )
    })
  })
}

function migrateDropPlantaColumn(callback) {
  db.all(`PRAGMA table_info(equipos)`, (e1, cols1) => {
    if (e1) return callback(e1)
    const hasPlantaEquipos = Array.isArray(cols1) && cols1.some((c) => c && c.name === "planta")
    const dropEquipos = hasPlantaEquipos
      ? (next) => db.run(`ALTER TABLE equipos DROP COLUMN planta`, (e) => next(e || null))
      : (next) => next(null)

    dropEquipos((e2) => {
      if (e2) return callback(e2)
      db.all(`PRAGMA table_info(usuarios)`, (e3, cols2) => {
        if (e3) return callback(e3)
        const hasPlantaUsuarios = Array.isArray(cols2) && cols2.some((c) => c && c.name === "planta")
        if (!hasPlantaUsuarios) return callback(null)
        db.run(`ALTER TABLE usuarios DROP COLUMN planta`, (e4) => callback(e4 || null))
      })
    })
  })
}

/** Limpia registros huérfanos cuando hubo borrados directos en equipos. */
function migrateCleanupOrphanEquipoData(callback) {
  db.all(`PRAGMA table_info(equipo_notas)`, (pragmaErr, nCols) => {
    if (pragmaErr) return callback(pragmaErr)
    const notasPorCodigo =
      Array.isArray(nCols) && nCols.some((c) => c && c.name === "equipo_codigo")
    if (notasPorCodigo) {
      db.run(
        `DELETE FROM equipo_notas WHERE equipo_codigo NOT IN (SELECT codigo FROM equipos)`,
        (e1) => {
          if (e1) return callback(e1)
          db.run(
            `DELETE FROM historial_equipos WHERE equipo_codigo NOT IN (SELECT codigo FROM equipos)`,
            (e2) => {
              if (e2) return callback(e2)
              db.run(
                `DELETE FROM historial WHERE equipo_codigo NOT IN (SELECT codigo FROM equipos)`,
                (e3) => callback(e3 || null)
              )
            }
          )
        }
      )
      return
    }
    db.run(
      `DELETE FROM equipo_notas WHERE id_equipo NOT IN (SELECT id FROM equipos)`,
      (e1) => {
        if (e1) return callback(e1)
        db.run(
          `DELETE FROM historial_equipos WHERE id_equipo NOT IN (SELECT id FROM equipos)`,
          (e2) => {
            if (e2) return callback(e2)
            db.run(
              `DELETE FROM historial WHERE equipo_id NOT IN (SELECT id FROM equipos)`,
              (e3) => callback(e3 || null)
            )
          }
        )
      }
    )
  })
}

/** Solo tablas legacy con columna id: asegura codigo (y copia desde codigo_unico si hace falta). */
function migrateEnsureCodigoUnico(done) {
  const reCod = new RegExp(`^${PREFIJO_CODIGO_UNICO}-(\\d+)$`, "i")
  db.all(`PRAGMA table_info(equipos)`, (pragmaErr, cols) => {
    if (pragmaErr) return done(pragmaErr)
    const tieneIdLegacy =
      Array.isArray(cols) && cols.some((c) => c && c.name === "id")
    if (!tieneIdLegacy) {
      return done()
    }
    const tieneCodigoUnico = Array.isArray(cols) && cols.some((c) => c && c.name === "codigo_unico")
    const syncFromUnico = tieneCodigoUnico
      ? (cb) => {
          db.run(
            `UPDATE equipos SET codigo = TRIM(codigo_unico) WHERE (codigo IS NULL OR TRIM(COALESCE(codigo,'')) = '') AND TRIM(COALESCE(codigo_unico,'')) != ''`,
            (uErr) => cb(uErr || null)
          )
        }
      : (cb) => cb(null)

    syncFromUnico((syncErr) => {
      if (syncErr) return done(syncErr)
      db.all(`SELECT id, codigo FROM equipos ORDER BY id ASC`, (err, rows) => {
        if (err) return done(err)
        const list = rows || []
        let maxN = 0
        for (const r of list) {
          const c = r.codigo != null ? String(r.codigo).trim() : ""
          const m = reCod.exec(c)
          if (m) maxN = Math.max(maxN, parseInt(m[1], 10))
        }
        const toFix = list.filter((r) => !r.codigo || String(r.codigo).trim() === "")
        if (toFix.length === 0) {
          return done()
        }
        db.run("BEGIN IMMEDIATE", (bErr) => {
          if (bErr) return done(bErr)
          let n = maxN + 1
          let i = 0
          const step = () => {
            if (i >= toFix.length) {
              return db.run("COMMIT", (cErr) => {
                if (cErr) return db.run("ROLLBACK", () => done(cErr))
                done()
              })
            }
            const id = toFix[i].id
            const code = `${PREFIJO_CODIGO_UNICO}-${String(n).padStart(4, "0")}`
            n += 1
            i += 1
            db.run("UPDATE equipos SET codigo = ? WHERE id = ?", [code, id], (uErr) => {
              if (uErr) return db.run("ROLLBACK", () => done(uErr))
              step()
            })
          }
          step()
        })
      })
    })
  })
}

/** Elimina columna codigo_unico; un solo codigo (PK). Recrea triggers de borrado en cascada. */
function migrateDropCodigoUnicoColumn(done) {
  db.all(`PRAGMA table_info(equipos)`, (pragmaErr, cols) => {
    if (pragmaErr) return done(pragmaErr)
    if (!Array.isArray(cols) || !cols.some((c) => c && c.name === "codigo_unico")) {
      return done()
    }

    const fail = (e) => db.run("ROLLBACK", () => done(e))

    db.run("BEGIN IMMEDIATE", (bErr) => {
      if (bErr) return done(bErr)

      db.run(`DROP TRIGGER IF EXISTS trg_equipos_delete_historial`, (t0) => {
        if (t0) return fail(t0)
        db.run(`DROP TRIGGER IF EXISTS trg_equipos_delete_historial_equipos`, (t1) => {
          if (t1) return fail(t1)
          db.run(`DROP TRIGGER IF EXISTS trg_equipos_delete_notas`, (t2) => {
            if (t2) return fail(t2)
            db.run(`DROP INDEX IF EXISTS idx_equipos_codigo_unico_natural`, () => {})
            db.run(`DROP INDEX IF EXISTS idx_equipos_codigo_unico`, () => {})
            db.run(
              `CREATE TABLE equipos_unificados (
                codigo TEXT PRIMARY KEY NOT NULL,
                departamento TEXT,
                planta TEXT,
                marca TEXT,
                modelo TEXT,
                tipo TEXT,
                licencias TEXT,
                usuario TEXT,
                desperfecto TEXT,
                estado TEXT,
                motivo_baja TEXT,
                comentario_inactivo TEXT,
                anio_compra TEXT,
                sistema_operativo TEXT,
                fecha_creacion DATETIME DEFAULT CURRENT_TIMESTAMP
              )`,
              (ce) => {
                if (ce) return fail(ce)
                db.run(
                  `INSERT INTO equipos_unificados (codigo, departamento, planta, marca, modelo, tipo, licencias, usuario, desperfecto, estado, motivo_baja, comentario_inactivo, anio_compra, sistema_operativo, fecha_creacion)
                   SELECT
                     CASE
                       WHEN TRIM(COALESCE(codigo,'')) != '' THEN TRIM(codigo)
                       WHEN TRIM(COALESCE(codigo_unico,'')) != '' THEN TRIM(codigo_unico)
                       ELSE 'VACIO-' || ROWID
                     END,
                     departamento, planta, marca, modelo, tipo, licencias, usuario, desperfecto, estado, motivo_baja, comentario_inactivo, anio_compra, sistema_operativo, fecha_creacion
                   FROM equipos`,
                  (ie) => {
                    if (ie) return fail(ie)
                    db.run(`DROP TABLE equipos`, (de) => {
                      if (de) return fail(de)
                      db.run(`ALTER TABLE equipos_unificados RENAME TO equipos`, (re) => {
                        if (re) return fail(re)
                        db.run(
                          `CREATE TRIGGER trg_equipos_delete_historial
                           AFTER DELETE ON equipos
                           BEGIN
                             DELETE FROM historial WHERE equipo_codigo = OLD.codigo;
                           END`,
                          (tc1) => {
                            if (tc1) return fail(tc1)
                            db.run(
                              `CREATE TRIGGER trg_equipos_delete_historial_equipos
                               AFTER DELETE ON equipos
                               BEGIN
                                 DELETE FROM historial_equipos WHERE equipo_codigo = OLD.codigo;
                               END`,
                              (tc2) => {
                                if (tc2) return fail(tc2)
                                db.run(
                                  `CREATE TRIGGER trg_equipos_delete_notas
                                   AFTER DELETE ON equipos
                                   BEGIN
                                     DELETE FROM equipo_notas WHERE equipo_codigo = OLD.codigo;
                                   END`,
                                  (tc3) => {
                                    if (tc3) return fail(tc3)
                                    db.run("COMMIT", (c) => done(c || null))
                                  }
                                )
                              }
                            )
                          }
                        )
                      })
                    })
                  }
                )
              }
            )
          })
        })
      })
    })
  })
}

/** Equipos pasan de PK numérica (id) a PK por codigo (TEXT); tablas hijas usan equipo_codigo. */
function migrateEquiposPrimaryKeyCodigo(done) {
  db.all(`PRAGMA table_info(equipos)`, (pragmaErr, cols) => {
    if (pragmaErr) return done(pragmaErr)
    if (!Array.isArray(cols) || !cols.some((c) => c && c.name === "id")) {
      return done()
    }

    db.all(`SELECT * FROM equipos ORDER BY id ASC`, (selErr, equiposList) => {
      if (selErr) return done(selErr)
      const list = equiposList || []
      const used = new Set()
      const idToCodigo = new Map()
      for (const r of list) {
        let cod =
          (r.codigo != null && String(r.codigo).trim() !== "" && String(r.codigo).trim()) ||
          (r.codigo_unico != null &&
            String(r.codigo_unico).trim() !== "" &&
            String(r.codigo_unico).trim()) ||
          `LEGACY-${r.id}`
        if (used.has(cod)) cod = `${cod}__${r.id}`
        used.add(cod)
        idToCodigo.set(r.id, cod)
      }

      const fail = (e) => {
        db.run("ROLLBACK", () => done(e))
      }

      db.run("BEGIN IMMEDIATE", (beginErr) => {
        if (beginErr) return done(beginErr)

        db.run(`DROP TRIGGER IF EXISTS trg_equipos_delete_historial`, (e1) => {
          if (e1) return fail(e1)
          db.run(`DROP TRIGGER IF EXISTS trg_equipos_delete_historial_equipos`, (e2) => {
            if (e2) return fail(e2)
            db.run(`DROP TRIGGER IF EXISTS trg_equipos_delete_notas`, (e3) => {
              if (e3) return fail(e3)
              db.run(`ALTER TABLE equipos RENAME TO equipos_old_pk_mig`, (e4) => {
                if (e4) return fail(e4)
                db.run(
                  `CREATE TABLE equipos (
                    codigo TEXT PRIMARY KEY NOT NULL,
                    departamento TEXT,
                    planta TEXT,
                    marca TEXT,
                    modelo TEXT,
                    tipo TEXT,
                    licencias TEXT,
                    usuario TEXT,
                    desperfecto TEXT,
                    estado TEXT,
                    motivo_baja TEXT,
                    comentario_inactivo TEXT,
                    anio_compra TEXT,
                    sistema_operativo TEXT,
                    fecha_creacion DATETIME DEFAULT CURRENT_TIMESTAMP
                  )`,
                  (e5) => {
                    if (e5) return fail(e5)
                    const stmt = `INSERT INTO equipos (codigo,departamento,planta,marca,modelo,tipo,licencias,usuario,desperfecto,estado,motivo_baja,comentario_inactivo,anio_compra,sistema_operativo,fecha_creacion) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`
                    let idx = 0
                    const insertEquiposLoop = () => {
                      if (idx >= list.length) {
                        return copyHistorial()
                      }
                      const r = list[idx]
                      idx += 1
                      const cod = idToCodigo.get(r.id)
                      db.run(
                        stmt,
                        [
                          cod,
                          r.departamento,
                          r.planta,
                          r.marca,
                          r.modelo,
                          r.tipo,
                          r.licencias,
                          r.usuario,
                          r.desperfecto,
                          r.estado,
                          r.motivo_baja,
                          r.comentario_inactivo,
                          r.anio_compra,
                          r.sistema_operativo,
                          r.fecha_creacion
                        ],
                        (ie) => {
                          if (ie) return fail(ie)
                          insertEquiposLoop()
                        }
                      )
                    }
                    insertEquiposLoop()

                    function copyHistorial() {
                      db.all(`SELECT * FROM historial`, (he, hRows) => {
                        if (he) return fail(he)
                        db.run(`ALTER TABLE historial RENAME TO historial_old_pk_mig`, (re) => {
                          if (re) return fail(re)
                          db.run(
                            `CREATE TABLE historial (
                              id INTEGER PRIMARY KEY AUTOINCREMENT,
                              equipo_codigo TEXT,
                              cambio TEXT,
                              fecha DATETIME DEFAULT CURRENT_TIMESTAMP
                            )`,
                            (ce) => {
                              if (ce) return fail(ce)
                              const hr = hRows || []
                              let hi = 0
                              const insH = () => {
                                if (hi >= hr.length) {
                                  db.run(`DROP TABLE historial_old_pk_mig`, (de) => {
                                    if (de) return fail(de)
                                    copyHistorialEquipos()
                                  })
                                  return
                                }
                                const h = hr[hi]
                                hi += 1
                                const legacyId = h.equipo_id
                                const eqc =
                                  legacyId != null && idToCodigo.has(legacyId)
                                    ? idToCodigo.get(legacyId)
                                    : null
                                db.run(
                                  `INSERT INTO historial (id, equipo_codigo, cambio, fecha) VALUES (?,?,?,?)`,
                                  [h.id, eqc, h.cambio, h.fecha],
                                  (ie) => {
                                    if (ie) return fail(ie)
                                    insH()
                                  }
                                )
                              }
                              insH()
                            }
                          )
                        })
                      })
                    }

                    function copyHistorialEquipos() {
                      db.all(`SELECT * FROM historial_equipos`, (he2, heRows) => {
                        if (he2) return fail(he2)
                        db.run(`ALTER TABLE historial_equipos RENAME TO historial_equipos_old_pk_mig`, (re2) => {
                          if (re2) return fail(re2)
                          db.run(
                            `CREATE TABLE historial_equipos (
                              id INTEGER PRIMARY KEY AUTOINCREMENT,
                              equipo_codigo TEXT NOT NULL,
                              usuario_anterior TEXT,
                              usuario_nuevo TEXT,
                              estado_anterior TEXT,
                              estado_nuevo TEXT,
                              comentario TEXT,
                              tipo_cambio TEXT,
                              fecha_cambio DATETIME DEFAULT CURRENT_TIMESTAMP
                            )`,
                            (ce2) => {
                              if (ce2) return fail(ce2)
                              const hers = heRows || []
                              let k = 0
                              const insHe = () => {
                                if (k >= hers.length) {
                                  db.run(`DROP TABLE historial_equipos_old_pk_mig`, (d2) => {
                                    if (d2) return fail(d2)
                                    copyNotas()
                                  })
                                  return
                                }
                                const row = hers[k]
                                k += 1
                                const eqc = idToCodigo.get(row.id_equipo)
                                if (eqc == null) {
                                  return insHe()
                                }
                                db.run(
                                  `INSERT INTO historial_equipos (id, equipo_codigo, usuario_anterior, usuario_nuevo, estado_anterior, estado_nuevo, comentario, tipo_cambio, fecha_cambio) VALUES (?,?,?,?,?,?,?,?,?)`,
                                  [
                                    row.id,
                                    eqc,
                                    row.usuario_anterior,
                                    row.usuario_nuevo,
                                    row.estado_anterior,
                                    row.estado_nuevo,
                                    row.comentario,
                                    row.tipo_cambio,
                                    row.fecha_cambio
                                  ],
                                  (ie) => {
                                    if (ie) return fail(ie)
                                    insHe()
                                  }
                                )
                              }
                              insHe()
                            }
                          )
                        })
                      })
                    }

                    function copyNotas() {
                      db.all(`SELECT * FROM equipo_notas`, (ne, nRows) => {
                        if (ne) return fail(ne)
                        db.run(`ALTER TABLE equipo_notas RENAME TO equipo_notas_old_pk_mig`, (re3) => {
                          if (re3) return fail(re3)
                          db.run(
                            `CREATE TABLE equipo_notas (
                              id INTEGER PRIMARY KEY AUTOINCREMENT,
                              equipo_codigo TEXT NOT NULL,
                              texto TEXT NOT NULL,
                              fecha_creacion DATETIME DEFAULT CURRENT_TIMESTAMP
                            )`,
                            (ce3) => {
                              if (ce3) return fail(ce3)
                              db.run(
                                `CREATE INDEX IF NOT EXISTS idx_equipo_notas_equipo ON equipo_notas(equipo_codigo)`,
                                () => {}
                              )
                              const nr = nRows || []
                              let ni = 0
                              const insN = () => {
                                if (ni >= nr.length) {
                                  db.run(`DROP TABLE equipo_notas_old_pk_mig`, (d3) => {
                                    if (d3) return fail(d3)
                                    finalizePkMig()
                                  })
                                  return
                                }
                                const row = nr[ni]
                                ni += 1
                                const eqc = idToCodigo.get(row.id_equipo)
                                if (eqc == null) {
                                  return insN()
                                }
                                db.run(
                                  `INSERT INTO equipo_notas (id, equipo_codigo, texto, fecha_creacion) VALUES (?,?,?,?)`,
                                  [row.id, eqc, row.texto, row.fecha_creacion],
                                  (ie) => {
                                    if (ie) return fail(ie)
                                    insN()
                                  }
                                )
                              }
                              insN()
                            }
                          )
                        })
                      })
                    }

                    function finalizePkMig() {
                      db.run(`DROP TABLE equipos_old_pk_mig`, (d4) => {
                        if (d4) return fail(d4)
                        db.run(
                          `CREATE TRIGGER trg_equipos_delete_historial
                          AFTER DELETE ON equipos
                          BEGIN
                            DELETE FROM historial WHERE equipo_codigo = OLD.codigo;
                          END`,
                          (te) => {
                            if (te) return fail(te)
                            db.run(
                              `CREATE TRIGGER trg_equipos_delete_historial_equipos
                              AFTER DELETE ON equipos
                              BEGIN
                                DELETE FROM historial_equipos WHERE equipo_codigo = OLD.codigo;
                              END`,
                              (te2) => {
                                if (te2) return fail(te2)
                                db.run(
                                  `CREATE TRIGGER trg_equipos_delete_notas
                                  AFTER DELETE ON equipos
                                  BEGIN
                                    DELETE FROM equipo_notas WHERE equipo_codigo = OLD.codigo;
                                  END`,
                                  (te3) => {
                                    if (te3) return fail(te3)
                                    db.run("COMMIT", (cErr) => done(cErr || null))
                                  }
                                )
                              }
                            )
                          }
                        )
                      })
                    }
                  }
                )
              })
            })
          })
        })
      })
    })
  })
}

function siguienteCodigoUnico(cb) {
  const reCod = new RegExp(`^${PREFIJO_CODIGO_UNICO}-(\\d+)$`, "i")
  db.all(
    `SELECT codigo FROM equipos WHERE codigo IS NOT NULL AND TRIM(codigo) != ''`,
    (err, rows) => {
      if (err) return cb(err)
      let maxN = 0
      for (const r of rows || []) {
        const m = reCod.exec(String(r.codigo).trim())
        if (m) maxN = Math.max(maxN, parseInt(m[1], 10))
      }
      cb(null, `${PREFIJO_CODIGO_UNICO}-${String(maxN + 1).padStart(4, "0")}`)
    }
  )
}

db.serialize(()=>{

db.run(`
CREATE TABLE IF NOT EXISTS equipos (
codigo TEXT PRIMARY KEY NOT NULL,
departamento TEXT,
marca TEXT,
modelo TEXT,
tipo TEXT,
licencias TEXT,
usuario TEXT,
desperfecto TEXT,
estado TEXT,
motivo_baja TEXT,
comentario_inactivo TEXT,
anio_compra TEXT,
sistema_operativo TEXT,
fecha_creacion DATETIME DEFAULT CURRENT_TIMESTAMP
)
`)

db.run(`ALTER TABLE equipos ADD COLUMN codigo TEXT`, (err)=>{
    // ignorar error si la columna ya existe
})

db.run(`ALTER TABLE equipos ADD COLUMN estado TEXT`, (err)=>{
    // ignorar error si la columna ya existe
})

db.run(`ALTER TABLE equipos ADD COLUMN comentario_inactivo TEXT`, (err)=>{
    // ignorar error si la columna ya existe
})

db.run(`ALTER TABLE equipos ADD COLUMN departamento TEXT`, (err)=>{
    // ignorar error si la columna ya existe
})

db.run(`ALTER TABLE equipos ADD COLUMN anio_compra TEXT`, (err)=>{
  // ignorar error si la columna ya existe
})

db.run(`ALTER TABLE equipos ADD COLUMN sistema_operativo TEXT`, (err)=>{
  // ignorar error si la columna ya existe
})

migrateRemovePlantasColumn((migrationErr) => {
  if (migrationErr) {
    console.error("migracion: no se pudo eliminar columna plantas", migrationErr)
  }
  migrateEnsureCodigoUnico((codigoErr) => {
    if (codigoErr) {
      console.error("migracion: no se pudo asegurar codigo en equipos (legacy)", codigoErr)
    }
    migrateCleanupOrphanEquipoData((orphanErr) => {
      if (orphanErr) {
        console.error("migracion: no se pudo limpiar historial/notas huerfanas", orphanErr)
      }
      migrateEquiposPrimaryKeyCodigo((pkErr) => {
        if (pkErr) {
          console.error("migracion: no se pudo migrar PK de equipos a codigo", pkErr)
          return
        }
        migrateDropCodigoUnicoColumn((dropErr) => {
          if (dropErr) {
            console.error("migracion: no se pudo unificar columna codigo (quitar codigo_unico)", dropErr)
          }
          migrateDropPlantaColumn((plantaDropErr) => {
            if (plantaDropErr) {
              console.error("migracion: no se pudo eliminar columna planta", plantaDropErr)
            }
          })
        })
      })
    })
  })
})

db.run(`
CREATE TABLE IF NOT EXISTS historial (
id INTEGER PRIMARY KEY AUTOINCREMENT,
equipo_codigo TEXT,
cambio TEXT,
fecha DATETIME DEFAULT CURRENT_TIMESTAMP
)
`)

db.run(`
CREATE TABLE IF NOT EXISTS historial_equipos (
id INTEGER PRIMARY KEY AUTOINCREMENT,
equipo_codigo TEXT NOT NULL,
usuario_anterior TEXT,
usuario_nuevo TEXT,
estado_anterior TEXT,
estado_nuevo TEXT,
comentario TEXT,
tipo_cambio TEXT,
fecha_cambio DATETIME DEFAULT CURRENT_TIMESTAMP
)
`)

db.run(`ALTER TABLE historial_equipos ADD COLUMN estado_anterior TEXT`, (err)=>{
  // ignorar error si la columna ya existe
})

db.run(`ALTER TABLE historial_equipos ADD COLUMN estado_nuevo TEXT`, (err)=>{
  // ignorar error si la columna ya existe
})

db.run(`ALTER TABLE historial_equipos ADD COLUMN tipo_cambio TEXT`, (err)=>{
  // ignorar error si la columna ya existe
})

db.run(`
CREATE TABLE IF NOT EXISTS equipo_notas (
id INTEGER PRIMARY KEY AUTOINCREMENT,
equipo_codigo TEXT NOT NULL,
texto TEXT NOT NULL,
fecha_creacion DATETIME DEFAULT CURRENT_TIMESTAMP
)
`)

db.run(`CREATE INDEX IF NOT EXISTS idx_equipo_notas_equipo ON equipo_notas(equipo_codigo)`, (err)=>{
  // ignorar error si ya existe
})

db.run(`
CREATE TRIGGER IF NOT EXISTS trg_equipos_delete_historial
AFTER DELETE ON equipos
BEGIN
  DELETE FROM historial WHERE equipo_codigo = OLD.codigo;
END;
`)

db.run(`
CREATE TRIGGER IF NOT EXISTS trg_equipos_delete_historial_equipos
AFTER DELETE ON equipos
BEGIN
  DELETE FROM historial_equipos WHERE equipo_codigo = OLD.codigo;
END;
`)

db.run(`
CREATE TRIGGER IF NOT EXISTS trg_equipos_delete_notas
AFTER DELETE ON equipos
BEGIN
  DELETE FROM equipo_notas WHERE equipo_codigo = OLD.codigo;
END;
`)

  db.run(`
CREATE TABLE IF NOT EXISTS usuarios (
id INTEGER PRIMARY KEY AUTOINCREMENT,
nombre TEXT NOT NULL UNIQUE,
password_hash TEXT NOT NULL
)
`)

  db.run(`ALTER TABLE usuarios ADD COLUMN activo INTEGER DEFAULT 1`, () => {})
  db.run(`ALTER TABLE usuarios ADD COLUMN es_superadmin INTEGER DEFAULT 0`, () => {})

  const nombreAdminEnv = process.env.INVENTARIO_ADMIN_USER || "admin"
  db.run(
    `UPDATE usuarios SET activo = 1 WHERE activo IS NULL`,
    () => {}
  )
  db.run(
    `UPDATE usuarios SET es_superadmin = 0 WHERE es_superadmin IS NULL`,
    () => {}
  )
  db.run(
    `UPDATE usuarios SET es_superadmin = 1 WHERE LOWER(TRIM(nombre)) = LOWER(TRIM(?))`,
    [nombreAdminEnv],
    () => {}
  )

  db.get(`SELECT COUNT(*) AS c FROM usuarios`, (countErr, crow) => {
    if (countErr) {
      console.error("usuarios: no se pudo contar filas", countErr)
      return
    }
    if (crow && crow.c > 0) return

    const nombreAdmin = process.env.INVENTARIO_ADMIN_USER || "admin"
    const passAdmin = process.env.INVENTARIO_ADMIN_PASSWORD || "admin123"
    const hash = bcrypt.hashSync(passAdmin, 10)
    const esSuper =
      String(nombreAdmin).trim().toLowerCase() ===
      String(nombreAdminEnv).trim().toLowerCase()
      ? 1
      : 0
    db.run(
      `INSERT INTO usuarios (nombre, password_hash, activo, es_superadmin) VALUES (?,?,?,?)`,
      [nombreAdmin, hash, 1, esSuper ? 1 : 0],
      (insErr) => {
        if (insErr) {
          console.error("usuarios: no se pudo crear usuario inicial", insErr)
          return
        }
        console.log(
          `Usuario inicial: "${nombreAdmin}". Cambie la clave en produccion (variable INVENTARIO_ADMIN_PASSWORD).`
        )
      }
    )
  })

})

/* ========================
   AUTENTICACION
======================== */

function esSuperadminSesion(req) {
  return !!(req.session && req.session.esSuperadmin)
}

function equipoAccesible(row) {
  return !!row
}

function quiereRespuestaJson(req) {
  const a = req.get("Accept") || ""
  return a.includes("application/json")
}

function requireAuth(req, res, next) {
  if (req.session && req.session.userId) {
    return next()
  }
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

function sendIndexHtml(req, res) {
  res.setHeader("Cache-Control", "no-store, no-cache, must-revalidate, private")
  res.setHeader("Pragma", "no-cache")
  res.setHeader("Expires", "0")
  res.sendFile(path.join(__dirname, "public", "index.html"))
}

app.get("/login", (req, res) => {
  if (req.session && req.session.userId) {
    return res.redirect("/")
  }
  res.sendFile(path.join(__dirname, "public", "login.html"))
})

app.post("/login", (req, res) => {
  const nombreRaw =
    req.body && typeof req.body.nombre === "string" ? req.body.nombre : ""
  const passRaw =
    req.body && typeof req.body.password === "string" ? req.body.password : ""
  const nombre = nombreRaw.trim()
  const password = passRaw
  if (!nombre || !password) {
    return res.status(400).json({ error: "Indique nombre y contraseña" })
  }
  db.get(
    `SELECT id, nombre, password_hash, activo, es_superadmin FROM usuarios WHERE LOWER(TRIM(nombre)) = LOWER(TRIM(?))`,
    [nombre],
    (err, row) => {
      if (err) {
        return res.status(500).json({ error: "Error al iniciar sesión" })
      }
      if (!row || !bcrypt.compareSync(password, String(row.password_hash || ""))) {
        return res.status(401).json({ error: "Credenciales incorrectas" })
      }
      const activo =
        row.activo == null || row.activo === "" ? 1 : Number(row.activo)
      if (activo !== 1) {
        return res.status(403).json({ error: "Usuario desactivado. Contacte al administrador." })
      }
      const esSuper = Number(row.es_superadmin) === 1
      req.session.userId = row.id
      req.session.nombre = row.nombre
      req.session.esSuperadmin = esSuper
      res.json({
        ok: true,
        user: { nombre: row.nombre, esSuperadmin: esSuper }
      })
    }
  )
})

app.get("/api/session", (req, res) => {
  if (!req.session || !req.session.userId) {
    return res.status(401).json({ authenticated: false })
  }
  res.json({
    authenticated: true,
    user: {
      id: req.session.userId,
      nombre: req.session.nombre,
      esSuperadmin: esSuperadminSesion(req)
    }
  })
})

app.get(
  "/api/usuarios",
  requireAuth,
  requireSuperadmin,
  (req, res) => {
    db.all(
      `SELECT id, nombre, activo, es_superadmin FROM usuarios ORDER BY nombre COLLATE NOCASE`,
      [],
      (err, rows) => {
        if (err) {
          return res.status(500).json({ error: "Error al listar usuarios" })
        }
        res.json(rows || [])
      }
    )
  }
)

app.post("/api/usuarios", requireAuth, requireSuperadmin, (req, res) => {
  const nombreRaw =
    req.body && typeof req.body.nombre === "string" ? req.body.nombre : ""
  const passRaw =
    req.body && typeof req.body.password === "string" ? req.body.password : ""
  const nombre = nombreRaw.trim()
  const password = passRaw
  const esSuperIn = req.body && Object.prototype.hasOwnProperty.call(req.body, "es_superadmin")
    ? req.body.es_superadmin
    : undefined
  let esSuperNuevo = 0
  if (esSuperIn !== undefined) {
    if (
      esSuperIn === true ||
      esSuperIn === 1 ||
      esSuperIn === "1"
    ) {
      esSuperNuevo = 1
    } else if (
      esSuperIn === false ||
      esSuperIn === 0 ||
      esSuperIn === "0"
    ) {
      esSuperNuevo = 0
    } else {
      return res.status(400).json({ error: "Rol inválido" })
    }
  }
  if (!nombre || !password) {
    return res
      .status(400)
      .json({ error: "Nombre y contraseña son obligatorios" })
  }
  if (password.length < 4) {
    return res
      .status(400)
      .json({ error: "La contraseña debe tener al menos 4 caracteres" })
  }
  const hash = bcrypt.hashSync(password, 10)
  db.run(
    `INSERT INTO usuarios (nombre, password_hash, activo, es_superadmin) VALUES (?,?,?,?)`,
    [nombre, hash, 1, esSuperNuevo],
    function (insErr) {
      if (insErr) {
        if (String(insErr.message || "").includes("UNIQUE")) {
          return res.status(409).json({ error: "Ya existe un usuario con ese nombre" })
        }
        if (esErrorSqliteBusy(insErr)) return responderBloqueoBd(res)
        return res.status(500).json({ error: "Error al crear usuario" })
      }
      res.status(201).json({
        ok: true,
        id: this.lastID,
        nombre,
        activo: 1,
        es_superadmin: esSuperNuevo
      })
    }
  )
})

app.patch("/api/usuarios/:id", requireAuth, requireSuperadmin, (req, res) => {
  const id = Number(req.params.id)
  if (!Number.isInteger(id) || id < 1) {
    return res.status(400).json({ error: "Id inválido" })
  }

  db.get(
    `SELECT id, nombre, activo, es_superadmin FROM usuarios WHERE id = ?`,
    [id],
    (gErr, target) => {
      if (gErr) {
        return res.status(500).json({ error: "Error al buscar usuario" })
      }
      if (!target) {
        return res.status(404).json({ error: "Usuario no encontrado" })
      }

      const nombreRaw =
        req.body && typeof req.body.nombre === "string" ? req.body.nombre : undefined
      const passRaw =
        req.body && typeof req.body.password === "string" ? req.body.password : undefined
      const activoIn = req.body && Object.prototype.hasOwnProperty.call(req.body, "activo")
        ? req.body.activo
        : undefined

      const esSuperActual = Number(target.es_superadmin) === 1
      let nuevoEsSuper = esSuperActual
      if (
        req.body &&
        Object.prototype.hasOwnProperty.call(req.body, "es_superadmin")
      ) {
        const raw = req.body.es_superadmin
        if (raw === true || raw === 1 || raw === "1") {
          nuevoEsSuper = true
        } else if (raw === false || raw === 0 || raw === "0") {
          nuevoEsSuper = false
        } else {
          return res.status(400).json({ error: "Rol inválido" })
        }
      }

      let nuevoNombre =
        nombreRaw !== undefined ? nombreRaw.trim() : String(target.nombre || "")
      if (!nuevoNombre) {
        return res.status(400).json({ error: "El nombre no puede estar vacío" })
      }

      const activoActual =
        target.activo == null || target.activo === "" ? 1 : Number(target.activo)
      let nuevoActivo = activoActual
      if (activoIn !== undefined) {
        nuevoActivo =
          activoIn === true || activoIn === 1 || activoIn === "1" ? 1 : 0
      }

      if (nuevoActivo !== 1 && id === Number(req.session.userId)) {
        return res.status(400).json({ error: "No puede desactivar su propia cuenta" })
      }

      const applyUpdate = () => {
        let newHash = null
        if (passRaw != null && String(passRaw).length > 0) {
          if (String(passRaw).length < 4) {
            return res.status(400).json({
              error: "La contraseña debe tener al menos 4 caracteres"
            })
          }
          newHash = bcrypt.hashSync(passRaw, 10)
        }

        const pieces = []
        const vals = []
        if (nuevoNombre !== target.nombre) {
          pieces.push("nombre = ?")
          vals.push(nuevoNombre)
        }
        if (nuevoActivo !== activoActual) {
          pieces.push("activo = ?")
          vals.push(nuevoActivo)
        }
        if (newHash) {
          pieces.push("password_hash = ?")
          vals.push(newHash)
        }
        if (nuevoEsSuper !== esSuperActual) {
          pieces.push("es_superadmin = ?")
          vals.push(nuevoEsSuper ? 1 : 0)
        }
        if (pieces.length === 0) {
          return res.json({
            ok: true,
            id,
            nombre: nuevoNombre,
            activo: nuevoActivo,
            es_superadmin: nuevoEsSuper
          })
        }
        vals.push(id)
        db.run(
          `UPDATE usuarios SET ${pieces.join(", ")} WHERE id = ?`,
          vals,
          function (uErr) {
            if (uErr) {
              if (String(uErr.message || "").includes("UNIQUE")) {
                return res
                  .status(409)
                  .json({ error: "Ya existe un usuario con ese nombre" })
              }
              if (esErrorSqliteBusy(uErr)) return responderBloqueoBd(res)
              return res.status(500).json({ error: "Error al actualizar usuario" })
            }
            res.json({
              ok: true,
              id,
              nombre: nuevoNombre,
              activo: nuevoActivo,
              es_superadmin: nuevoEsSuper
            })
          }
        )
      }

      if (esSuperActual && (nuevoActivo !== 1 || !nuevoEsSuper)) {
        db.get(
          `SELECT COUNT(*) AS c FROM usuarios WHERE es_superadmin = 1 AND (activo = 1 OR activo IS NULL) AND id != ?`,
          [id],
          (cErr, crow) => {
            if (cErr) {
              return res.status(500).json({ error: "Error al validar superadmins" })
            }
            const otros = crow && crow.c != null ? Number(crow.c) : 0
            if (otros < 1) {
              return res.status(400).json({
                error:
                  "Debe existir al menos otro superadministrador activo antes de desactivar o quitar el rol superadmin a este usuario"
              })
            }
            applyUpdate()
          }
        )
        return
      }

      applyUpdate()
    }
  )
})

app.post("/logout", (req, res) => {
  req.session.destroy((e) => {
    if (e) {
      return res.status(500).json({ error: "No se pudo cerrar la sesión" })
    }
    res.clearCookie("inventario.sid", { path: "/" })
    res.json({ ok: true })
  })
})

app.get("/", requireAuth, sendIndexHtml)
app.get("/index.html", requireAuth, sendIndexHtml)

app.use(express.static(path.join(__dirname, "public"), { index: false, etag: false, lastModified: false, setHeaders(res) { res.setHeader("Cache-Control", "no-store, no-cache, must-revalidate, private"); res.setHeader("Pragma", "no-cache"); res.setHeader("Expires", "0"); } }))

/* ========================
   DASHBOARD
======================== */

function dashboardVistaEquiposConfig(tipo) {
  const map = {
    todos: {
      titulo: "Total — todos los equipos",
      descripcion: "Incluye activos, inactivos y dados de baja.",
      sql: `SELECT codigo, departamento, marca, modelo, tipo, usuario, estado
            FROM equipos ORDER BY datetime(fecha_creacion) DESC, codigo DESC`,
      params: []
    },
    activos: {
      titulo: "Activos",
      descripcion: "Equipos actualmente marcados como activos.",
      sql: `SELECT codigo, departamento, marca, modelo, tipo, usuario, estado
            FROM equipos WHERE estado = 'Activo' ORDER BY datetime(fecha_creacion) DESC, codigo DESC`,
      params: []
    },
    inactivos: {
      titulo: "Inactivos",
      descripcion: "Equipos inactivos (requieren motivo al pasar a inactivo).",
      sql: `SELECT codigo, departamento, marca, modelo, tipo, usuario, estado
            FROM equipos WHERE estado = 'Inactivo' ORDER BY datetime(fecha_creacion) DESC, codigo DESC`,
      params: []
    }
  }
  return map[tipo] || null
}

app.get("/dashboard", requireAuth, (req, res) => {
  db.get(
    `SELECT COUNT(*) AS total FROM equipos`,
    [],
    (err, total) => {
      if (err || !total) {
        return res.status(500).send("Error al cargar el dashboard")
      }
      db.get(
        `SELECT COUNT(*) AS activos FROM equipos WHERE estado='Activo'`,
        [],
        (err2, activos) => {
          if (err2 || !activos) {
            return res.status(500).send("Error al cargar el dashboard")
          }
          db.get(
            `SELECT COUNT(*) AS inactivos FROM equipos WHERE estado='Inactivo'`,
            [],
            (err3, inactivos) => {
              if (err3 || !inactivos) {
                return res.status(500).send("Error al cargar el dashboard")
              }
              db.get(
                `SELECT COUNT(*) AS eliminados FROM equipos WHERE estado='Eliminado'`,
                [],
                (err4, eliminados) => {
                  if (err4 || !eliminados) {
                    return res.status(500).send("Error al cargar el dashboard")
                  }

                  res.send(`
<style>
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
<div class="card"><b>Total</b><div class="numero">${total.total}</div><div style="font-size:12px;color:#666;margin-top:6px;">Ver listado completo</div></div>
</a>
<a class="card-link" href="/dashboard/vista/activos" title="Ver equipos activos">
<div class="card"><b>Activos</b><div class="numero">${activos.activos}</div><div style="font-size:12px;color:#666;margin-top:6px;">Ver equipos en uso</div></div>
</a>
<a class="card-link" href="/dashboard/vista/inactivos" title="Ver equipos inactivos">
<div class="card"><b>Inactivos</b><div class="numero">${inactivos.inactivos}</div><div style="font-size:12px;color:#666;margin-top:6px;">Ver equipos inactivos</div></div>
</a>
<a class="card-link" href="/historial-bajas" title="Ver equipos dados de baja">
<div class="card"><b>Dados de baja</b><div class="numero">${eliminados.eliminados}</div><div style="font-size:12px;color:#666;margin-top:6px;">Ver historial de bajas</div></div>
</a>
</div>

<br>
<button type="button" class="btn" onclick="window.location.href='/'">📋 Volver al inventario</button>
`)

                }
              )
            }
          )
        }
      )
    }
  )
})

app.get("/dashboard/vista/:tipo", requireAuth, (req, res) => {
  const cfg = dashboardVistaEquiposConfig(req.params.tipo)
  if (!cfg) {
    return res.redirect("/dashboard")
  }
  db.all(cfg.sql, cfg.params, (err, rows) => {
    if (err) {
      return res.status(500).send("Error al cargar el listado del dashboard")
    }
    const list = rows || []
    const filas = list
      .map(
        (r) => `
<tr>
<td>${escapeHtml(codigoEquipoMostrar(r))}</td>
<td>${escapeHtml(r.marca || "")}</td>
<td>${escapeHtml(r.modelo || "")}</td>
<td>${escapeHtml(r.tipo || "")}</td>
<td>${escapeHtml(nombreAreaMostrar(r.departamento) || "-")}</td>
<td>${escapeHtml(r.usuario || "")}</td>
<td>${escapeHtml(normalizeEstadoDisplay(r.estado))}</td>
<td><a href="/equipo/${encodeURIComponent(r.codigo)}">Ver detalle</a></td>
</tr>`
      )
      .join("")

    res.send(`
<html>
<head>
<meta charset="UTF-8">
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
</head>
<body>
<div class="card">
<h2 style="margin:0 0 6px;">${escapeHtml(cfg.titulo)}</h2>
<p class="sub">${escapeHtml(cfg.descripcion)}</p>
<div class="sub"><b>Registros:</b> ${list.length}</div>
<table>
<thead>
<tr>
<th>Código único</th><th>Marca</th><th>Modelo</th><th>Tipo</th><th>Área</th><th>Usuario</th><th>Estado</th><th>Detalle</th>
</tr>
</thead>
<tbody>
${filas || '<tr><td colspan="8">No hay equipos en esta categoría.</td></tr>'}
</tbody>
</table>
<div class="acciones">
<button class="btn" type="button" onclick="window.location.href='/dashboard'">⬅️ Volver al dashboard</button>
<button class="btn" type="button" onclick="window.location.href='/'">📋 Volver al inventario</button>
</div>
</div>
</body>
</html>
    `)
  })
})

app.get("/historial-bajas", requireAuth, (req, res) => {
  db.all(
    `SELECT codigo, departamento, marca, modelo, tipo, usuario, motivo_baja, fecha_creacion
     FROM equipos
     WHERE estado='Eliminado'
     ORDER BY datetime(fecha_creacion) DESC, codigo DESC`,
    [],
    (err, rows) => {
      if (err) {
        return res.status(500).send("Error al cargar historial de bajas")
      }

      const filas = (rows || [])
        .map(
          (r) => `
<tr>
<td>${escapeHtml(codigoEquipoMostrar(r))}</td>
<td>${escapeHtml(r.marca || "")}</td>
<td>${escapeHtml(r.modelo || "")}</td>
<td>${escapeHtml(r.tipo || "")}</td>
<td>${escapeHtml(nombreAreaMostrar(r.departamento) || "-")}</td>
<td>${escapeHtml(r.usuario || "")}</td>
<td>${escapeHtml(r.motivo_baja || "-")}</td>
<td><a href="/equipo/${encodeURIComponent(r.codigo)}">Ver detalle</a></td>
</tr>`
        )
        .join("")

      res.send(`
<html>
<head>
<meta charset="UTF-8">
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
</head>
<body>
<div class="card">
<h2 style="margin:0 0 8px;">Historial de equipos dados de baja</h2>
<div style="color:#666;font-size:14px;">Total: ${(rows || []).length}</div>
<table>
<thead>
<tr>
<th>Código único</th><th>Marca</th><th>Modelo</th><th>Tipo</th><th>Área</th><th>Usuario</th><th>Motivo baja</th><th>Detalle</th>
</tr>
</thead>
<tbody>
${filas || '<tr><td colspan="8">No hay equipos dados de baja.</td></tr>'}
</tbody>
</table>
<div class="acciones">
<button class="btn" type="button" onclick="window.location.href='/dashboard'">⬅️ Volver al dashboard</button>
<button class="btn" type="button" onclick="window.location.href='/'">📋 Volver al inventario</button>
</div>
</div>
</body>
</html>
      `)
    }
  )
})

/* ========================
   CREAR EQUIPO
======================== */

app.post("/equipos", requireAuth, (req, res) => {
  const marcaRaw = req.body && typeof req.body.marca === "string" ? req.body.marca : ""
  const modeloRaw = req.body && typeof req.body.modelo === "string" ? req.body.modelo : ""
  const tipoRaw = req.body && typeof req.body.tipo === "string" ? req.body.tipo : ""
  const licenciasRaw =
    req.body && typeof req.body.licencias === "string" ? req.body.licencias : ""
  const usuarioRaw = req.body && typeof req.body.usuario === "string" ? req.body.usuario : ""
  const estadoRaw = req.body && typeof req.body.estado === "string" ? req.body.estado : ""
  const marca = marcaRaw.trim()
  const modelo = modeloRaw.trim()
  const tipo = tipoRaw.trim()
  const tipoCanonico = tipo === "Notebook" ? "Notebooks" : tipo
  const licencias = licenciasRaw.trim()
  const usuario = usuarioRaw.trim()
  const estado = estadoRaw.trim()
  const departamentoRaw =
    req.body && typeof req.body.departamento === "string" ? req.body.departamento : ""
  const departamento = departamentoRaw.trim().toUpperCase()
  const comentarioRaw =
    req.body && typeof req.body.comentario === "string" ? req.body.comentario : ""
  const comentario = comentarioRaw.trim()
  const anioCompraRaw =
    req.body && typeof req.body.anio_compra === "string" ? req.body.anio_compra : ""
  const anioCompra = anioCompraRaw.trim()
  const sistemaOpRaw =
    req.body && typeof req.body.sistema_operativo === "string" ? req.body.sistema_operativo : ""
  const sistemaOpCliente = sistemaOpRaw.trim()
  const DEPARTAMENTOS_VALIDOS = ["TI", "BOD", "FRI", "RH", "GER", "ADM", "PAC", "MAN"]
  const TIPOS_VALIDOS = [
    "Antena",
    "Router",
    "Repetidores",
    "Switch",
    "Notebook",
    "Notebooks",
    "Escritorio",
    "Impresora"
  ]
  const TIPOS_CON_SISTEMA_OPERATIVO = ["Notebooks", "Escritorio"]
  const TIPOS_CON_LICENCIAS = ["Notebooks", "Escritorio"]
  const SISTEMAS_OPERATIVOS_VALIDOS = [
    "Windows 7",
    "Windows 10",
    "Windows 10 Pro",
    "Windows 11 Home",
    "Windows 11 Pro",
    "Linux"
  ]
  const LICENCIAS_VALIDAS = ["No", "Licencia Office standard"]
  const ESTADOS_VALIDOS = ["Activo", "Inactivo"]
  if (!marca || !modelo || !tipoCanonico || !licencias || !usuario || !estado || !anioCompra) {
    return res.status(400).json({ error: "Todos los campos del formulario son obligatorios" })
  }

  if (!/^\d{4}$/.test(anioCompra)) {
    return res.status(400).json({ error: "El año de compra debe tener exactamente 4 dígitos" })
  }

  const anioNumCompra = Number(anioCompra, 10)
  const anioMaxPermitido = anioCalendarioActualZona()
  if (anioNumCompra > anioMaxPermitido) {
    return res.status(400).json({
      error: `El año de compra no puede ser mayor a ${anioMaxPermitido}`
    })
  }

  let sistemaOperativoDb = null
  if (TIPOS_CON_SISTEMA_OPERATIVO.includes(tipoCanonico)) {
    if (!SISTEMAS_OPERATIVOS_VALIDOS.includes(sistemaOpCliente)) {
      return res.status(400).json({ error: "Seleccione un sistema operativo válido" })
    }
    sistemaOperativoDb = sistemaOpCliente
  } else if (sistemaOpCliente !== "") {
    return res.status(400).json({
      error: "El sistema operativo solo aplica a Notebooks y Escritorio"
    })
  }

  if (!DEPARTAMENTOS_VALIDOS.includes(departamento)) {
    return res.status(400).json({ error: "Seleccione un área válida" })
  }
  if (!TIPOS_VALIDOS.includes(tipoCanonico)) {
    return res.status(400).json({ error: "Seleccione un tipo de equipo valido" })
  }
  if (TIPOS_CON_LICENCIAS.includes(tipoCanonico)) {
    if (!LICENCIAS_VALIDAS.includes(licencias)) {
      return res.status(400).json({ error: "Seleccione una licencia valida" })
    }
  } else if (licencias !== "No") {
    return res.status(400).json({
      error: "La licencia solo aplica a Notebooks o Escritorio"
    })
  }
  if (!ESTADOS_VALIDOS.includes(estado)) {
    return res.status(400).json({ error: "Seleccione un estado valido" })
  }
  if (estado === "Inactivo" && !comentario) {
    return res.status(400).json({
      error: "Debe ingresar el motivo al marcar el equipo como Inactivo"
    })
  }

  const comentarioDb = estado === "Inactivo" ? comentario : null

  db.run("BEGIN IMMEDIATE TRANSACTION", (beginErr) => {
    if (beginErr) {
      if (esErrorSqliteBusy(beginErr)) return responderBloqueoBd(res)
      return res.status(500).json({ error: "No se pudo iniciar la creacion del equipo" })
    }

    siguienteCodigoUnico((genErr, codigoUnico) => {
      if (genErr || !codigoUnico) {
        return db.run("ROLLBACK", () => {
          res.status(500).json({ error: "No se pudo generar el codigo del equipo" })
        })
      }

      db.run(
        `INSERT INTO equipos
         (codigo,departamento,marca,modelo,tipo,licencias,usuario,desperfecto,estado,comentario_inactivo,anio_compra,sistema_operativo)
         VALUES (?,?,?,?,?,?,?,?,?,?,?,?)`,
        [
          codigoUnico,
          departamento,
          marca,
          modelo,
          tipoCanonico,
          licencias,
          usuario,
          null,
          estado,
          comentarioDb,
          anioCompra,
          sistemaOperativoDb
        ],
        function (insertErr) {
          if (insertErr) {
            return db.run("ROLLBACK", () => {
              if (String(insertErr.message || "").includes("UNIQUE")) {
                return res
                  .status(409)
                  .json({ error: "Conflicto de codigo. Intente crear el equipo nuevamente." })
              }
              if (esErrorSqliteBusy(insertErr)) return responderBloqueoBd(res)
              res.status(500).json({ error: "Error al crear equipo" })
            })
          }

          db.run(
            `INSERT INTO historial (equipo_codigo,cambio)
             VALUES (?,?)`,
            [codigoUnico, `Equipo creado (${codigoUnico})`],
            (histErr) => {
              if (histErr) {
                return db.run("ROLLBACK", () => {
                  if (esErrorSqliteBusy(histErr)) return responderBloqueoBd(res)
                  res.status(500).json({ error: "Error al registrar historial de creacion" })
                })
              }

              db.run("COMMIT", (commitErr) => {
                if (commitErr) {
                  return db.run("ROLLBACK", () => {
                    if (esErrorSqliteBusy(commitErr)) return responderBloqueoBd(res)
                    res.status(500).json({ error: "No se pudo finalizar la creacion del equipo" })
                  })
                }

                res.status(201).json({
                  ok: true,
                  codigo: codigoUnico,
                  departamento,
                  departamento_nombre: nombreAreaMostrar(departamento)
                })
              })
            }
          )
        }
      )
    })
  })
})

/* ========================
   LISTAR EQUIPOS
======================== */

app.get("/equipos", requireAuth, (req, res) => {
  db.all(
    `SELECT * FROM equipos WHERE estado!='Eliminado' ORDER BY datetime(fecha_creacion) DESC, codigo DESC`,
    [],
    (err, rows) => {
      if (err) {
        return res.status(500).json({ error: "Error al listar equipos" })
      }
      res.json(rows || [])
    }
  )
})

/* ========================
   BUSCAR (texto + estado)
======================== */

const ESTADOS_FILTRO = ["Todos", "Activo", "Inactivo"]

app.get("/buscar", requireAuth, (req, res) => {
  const rawQ = typeof req.query.q === "string" ? req.query.q.trim() : ""
  const rawEstado = typeof req.query.estado === "string" ? req.query.estado : "Todos"
  const estadoFiltro = ESTADOS_FILTRO.includes(rawEstado) ? rawEstado : "Todos"

  let sql = `SELECT * FROM equipos WHERE estado != 'Eliminado'`
  const params = []

  if (estadoFiltro === "Activo" || estadoFiltro === "Inactivo") {
    sql += ` AND estado = ?`
    params.push(estadoFiltro)
  }

  if (rawQ !== "") {
    const like = `%${rawQ}%`
    const deptCodes = codigosDepartamentoQueCoincidenConBusqueda(rawQ)
    const deptInClause =
      deptCodes.length > 0
        ? ` OR departamento IN (${deptCodes.map(() => "?").join(",")})`
        : ""
    sql += ` AND (marca LIKE ? OR modelo LIKE ? OR usuario LIKE ? OR tipo LIKE ? OR codigo LIKE ? OR IFNULL(anio_compra,'') LIKE ? OR IFNULL(sistema_operativo,'') LIKE ?${deptInClause})`
    params.push(like, like, like, like, like, like, like)
    if (deptCodes.length > 0) params.push(...deptCodes)
  }

  sql += ` ORDER BY datetime(fecha_creacion) DESC, codigo DESC`

  db.all(sql, params, (err, rows) => {
    if (err) {
      return res.status(500).json({ error: "Error al buscar equipos" })
    }
    res.json(rows)
  })
})

/* ========================
   DETALLE EQUIPO (solo lectura)
======================== */

app.get("/equipo/:codigo", requireAuth, (req, res) => {
  const codigo = codigoEquipoDesdeParams(req)
  if (!codigo) {
    return res.status(400).send("Codigo de equipo invalido")
  }

  db.get("SELECT * FROM equipos WHERE codigo = ?", [codigo], (err, row) => {
    if (!row) {
      return res.send("Equipo no encontrado")
    }

    const motivoDisplay = row.estado === "Inactivo" ? "block" : "none"
    const urlEtiqueta = `/etiqueta/${encodeURIComponent(row.codigo)}?c=${encodeURIComponent(codigoEquipoMostrar(row))}`

    res.send(`
<html>
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
var codigoEquipo = ${JSON.stringify(row.codigo)};
var lista = document.getElementById("listaNotas");
function escapeText(s){
return String(s == null ? "" : s)
.replace(/&/g, "&amp;")
.replace(/</g, "&lt;")
.replace(/>/g, "&gt;")
.replace(/"/g, "&quot;");
}
function renderNotas(items){
if (!lista) return;
if (!items || !items.length){
lista.innerHTML = '<li class="nota-vacio">No hay notas registradas.</li>';
return;
}
lista.innerHTML = items.map(function(n){
return '<li class="nota-item"><span class="nota-fecha">' + escapeText(n.fecha) + '</span><div class="nota-texto">' + escapeText(n.texto) + '</div></li>';
}).join("");
}
function cargarNotas(){
fetch("/equipo/" + encodeURIComponent(codigoEquipo) + "/notas", { credentials: "include", headers: { Accept: "application/json" } })
.then(function(res){ return res.ok ? res.json() : Promise.reject(); })
.then(renderNotas)
.catch(function(){
lista.innerHTML = '<li class="nota-vacio">No se pudo cargar el historial de notas.</li>';
});
}
var btn = document.getElementById("btnAgregarNota");
if (btn){
btn.addEventListener("click", function(){
if (typeof Swal === "undefined"){
alert("No se puede abrir el formulario de nota.");
return;
}
Swal.fire({
title: "Agregar nota",
input: "textarea",
inputLabel: "Comentario o cambio realizado",
inputPlaceholder: "Ej.: Se cambió disco SSD, reinstalación de Windows…",
showCancelButton: true,
confirmButtonText: "Guardar",
cancelButtonText: "Cancelar",
confirmButtonColor: "#0B47BF",
cancelButtonColor: "#6c757d",
inputAttributes: { "aria-label": "Texto de la nota" },
inputValidator: function(v){
var t = v != null ? String(v).trim() : "";
if (!t) return "Escriba un texto para la nota";
if (t.length > ${NOTA_MAX_LENGTH}) return "El texto no puede superar ${NOTA_MAX_LENGTH} caracteres";
}
}).then(function(result){
if (!result.isConfirmed) return;
var texto = String(result.value).trim();
fetch("/equipo/" + encodeURIComponent(codigoEquipo) + "/notas", {
method: "POST",
credentials: "include",
headers: { "Content-Type": "application/json", Accept: "application/json" },
body: JSON.stringify({ texto: texto })
})
.then(function(res){
return res.json().then(function(data){ return { res: res, data: data }; });
})
.then(function(r){
if (!r.res.ok){
var err = (r.data && r.data.error) || "No se pudo guardar la nota";
Swal.fire({ icon: "error", title: "Error", text: err, confirmButtonColor: "#0B47BF" });
return;
}
cargarNotas();
Swal.fire({ icon: "success", title: "Nota guardada", timer: 1600, showConfirmButton: false });
})
.catch(function(){
Swal.fire({ icon: "error", title: "Error", text: "Sin conexión o servidor no disponible.", confirmButtonColor: "#0B47BF" });
});
});
});
}
cargarNotas();
})();
</script>
</body>
</html>
`)
  })
})

const NOTA_MAX_LENGTH = 4000

app.get("/equipo/:codigo/notas", requireAuth, (req, res) => {
  const codigo = codigoEquipoDesdeParams(req)
  if (!codigo) {
    return res.status(400).json({ error: "Codigo de equipo invalido" })
  }
  db.get("SELECT codigo FROM equipos WHERE codigo = ?", [codigo], (gErr, eqRow) => {
    if (gErr) {
      return res.status(500).json({ error: "Error al verificar el equipo" })
    }
    if (!eqRow) {
      return res.status(404).json({ error: "Equipo no encontrado" })
    }
  db.all(
    `SELECT id, texto, fecha_creacion FROM equipo_notas WHERE equipo_codigo = ? ORDER BY fecha_creacion DESC, id DESC`,
    [codigo],
    (err, rows) => {
      if (err) {
        return res.status(500).json({ error: "Error al cargar notas" })
      }
      const list = (rows || []).map((r) => ({
        id: r.id,
        texto: r.texto,
        fecha: formatFechaLocal(r.fecha_creacion)
      }))
      res.json(list)
    })
  })
})

app.post("/equipo/:codigo/notas", requireAuth, (req, res) => {
  const codigo = codigoEquipoDesdeParams(req)
  if (!codigo) {
    return res.status(400).json({ error: "Codigo de equipo invalido" })
  }
  const raw =
    req.body && typeof req.body.texto === "string" ? req.body.texto : ""
  const texto = raw.trim()
  if (!texto) {
    return res.status(400).json({ error: "La nota no puede estar vacia" })
  }
  if (texto.length > NOTA_MAX_LENGTH) {
    return res.status(400).json({
      error: `La nota no puede superar ${NOTA_MAX_LENGTH} caracteres`
    })
  }

  db.get("SELECT codigo FROM equipos WHERE codigo = ?", [codigo], (err, row) => {
    if (err) {
      return res.status(500).json({ error: "Error al verificar el equipo" })
    }
    if (!row) {
      return res.status(404).json({ error: "Equipo no encontrado" })
    }
    db.run(
      `INSERT INTO equipo_notas (equipo_codigo, texto) VALUES (?, ?)`,
      [codigo, texto],
      function (insertErr) {
        if (insertErr) {
          return res.status(500).json({ error: "Error al guardar la nota" })
        }
        res.status(201).json({
          ok: true,
          id: this.lastID,
          fecha: formatFechaLocal(new Date().toISOString())
        })
      }
    )
  })
})

/* ========================
   EDITAR EQUIPO
======================== */

app.get("/equipo/:codigo/editar", requireAuth, (req, res) => {

const codigo = codigoEquipoDesdeParams(req)
if (!codigo) {
return res.status(400).send("Codigo de equipo invalido")
}

db.get("SELECT * FROM equipos WHERE codigo = ?",[codigo],(err,row)=>{

if(!row){
return res.send("Equipo no encontrado")
}

const urlEtiqueta = `/etiqueta/${encodeURIComponent(row.codigo)}?c=${encodeURIComponent(codigoEquipoMostrar(row))}`

res.send(`

<html>

<head>

<meta charset="UTF-8">
<script src="https://cdn.jsdelivr.net/npm/sweetalert2@11"></script>

<style>

body{
font-family:Arial;
background:#f4f6f9;
margin:0;
}

.container{
padding:40px;
display:flex;
justify-content:center;
}

.card{
background:white;
padding:30px;
border-radius:12px;
box-shadow:0 4px 15px rgba(0,0,0,0.1);
width:500px;
animation:fadeIn 0.5s ease;
}

h1{
margin-bottom:20px;
}

.item{
margin-bottom:10px;
}

.codigo{
font-size:22px;
font-weight:bold;
color:#0B47BF;
margin-bottom:15px;
}

button{
background:#9DBF21;
color:white;
border:none;
padding:10px;
margin-top:10px;
cursor:pointer;
border-radius:6px;
}

a{
text-decoration:none;
display:inline-block;
margin-top:20px;
}

@keyframes fadeIn{
from{opacity:0; transform:translateY(10px)}
to{opacity:1; transform:translateY(0)}
}

.form-cambios{
margin-top:18px;
padding-top:18px;
border-top:1px solid #eee;
}

.form-cambios label{
display:block;
margin-bottom:8px;
}

.form-cambios input,
.form-cambios select{
width:100%;
padding:10px;
border-radius:6px;
border:1px solid #ccc;
font-size:15px;
margin-bottom:10px;
}

.form-cambios textarea{
width:100%;
padding:10px;
border-radius:6px;
border:1px solid #ccc;
font-size:15px;
font-family:Arial,sans-serif;
resize:vertical;
min-height:72px;
box-sizing:border-box;
margin-bottom:10px;
}

.form-cambios button[type="submit"]{
width:100%;
}

.msg-cambios{
margin:10px 0 0;
font-size:14px;
min-height:20px;
}

.historial-cambios{
margin-top:14px;
max-height:240px;
overflow:auto;
border:1px solid #e9ecef;
border-radius:8px;
padding:10px;
background:#fafafa;
}

.historial-tabs{
display:flex;
gap:8px;
margin:10px 0 10px;
}

.historial-tab{
background:#e9ecef;
color:#333;
border:none;
padding:8px 12px;
border-radius:999px;
cursor:pointer;
font-size:13px;
}

.historial-tab.activo{
background:#0B47BF;
color:#fff;
}

.historial-panel{
display:none;
}

.historial-panel.activo{
display:block;
}

.historial-cambios ul{
margin:0;
padding-left:18px;
}

.historial-cambios li{
margin-bottom:8px;
}

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
<option value="TI" ${row.departamento === "TI" ? "selected" : ""}>TI</option>
<option value="BOD" ${row.departamento === "BOD" ? "selected" : ""}>Bodega</option>
<option value="FRI" ${row.departamento === "FRI" ? "selected" : ""}>Frigorifico</option>
<option value="RH" ${row.departamento === "RH" ? "selected" : ""}>RRHH</option>
<option value="GER" ${row.departamento === "GER" ? "selected" : ""}>Gerencia</option>
<option value="ADM" ${row.departamento === "ADM" ? "selected" : ""}>Administracion</option>
<option value="PAC" ${row.departamento === "PAC" ? "selected" : ""}>Packing</option>
<option value="MAN" ${row.departamento === "MAN" ? "selected" : ""}>Mantencion</option>
</select>

<p style="margin:0 0 10px;font-size:12px;color:#666;">El código del equipo no cambia; solo se actualizan los campos editables.</p>

<label for="estadoEquipo">Estado</label>
<select id="estadoEquipo" name="estado" aria-label="Estado del equipo">
<option value="Activo" ${row.estado === "Activo" ? "selected" : ""}>Activo</option>
<option value="Inactivo" ${row.estado === "Inactivo" ? "selected" : ""}>Inactivo</option>
<option value="Eliminado" ${row.estado === "Eliminado" ? "selected" : ""}>Dar de baja</option>
</select>
<div id="wrapComentarioInactivo" class="wrap-comentario-inactivo" style="display:${row.estado === "Inactivo" ? "block" : "none"}">
<label for="comentarioCambio">Motivo de inactividad</label>
<textarea id="comentarioCambio" name="comentario" rows="3" placeholder="Obligatorio solo al pasar a Inactivo">${escapeHtml(
    row.comentario_inactivo || ""
  )}</textarea>
</div>
<div id="wrapComentarioReparacion" class="wrap-comentario-inactivo" style="display:none">
<label for="comentarioReparacion">¿Qué se reparó del equipo?</label>
<textarea id="comentarioReparacion" name="comentario_reactivacion" rows="3" placeholder="Obligatorio al pasar de Inactivo a Activo"></textarea>
</div>
<div id="wrapMotivoBaja" class="wrap-comentario-inactivo" style="display:none">
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
<div id="panelTransferencias" class="historial-panel activo">
<ul id="listaTransferencias"></ul>
</div>
<div id="panelEstados" class="historial-panel">
<ul id="listaEstados"></ul>
</div>
<div id="panelUbicacion" class="historial-panel">
<ul id="listaUbicacion"></ul>
</div>
</div>

<br>

<button type="button" onclick='window.open(${JSON.stringify(urlEtiqueta)})'>
🖨️ Generar Etiqueta
</button>

<br>

<button type="button" onclick="window.location.href='/equipo/${encodeURIComponent(row.codigo)}'">⬅️ Volver a lectura</button>
<br>
<button type="button" onclick="window.location.href='/'">📋 Volver al inventario</button>

</div>

</div>

<script>
(function(){
var codigoEquipo = ${JSON.stringify(row.codigo)};
var form = document.getElementById("formCambios");
var span = document.getElementById("estadoTexto");
var usuarioTexto = document.getElementById("usuarioTexto");
var msg = document.getElementById("msgCambios");
var usuarioInput = document.getElementById("usuarioEquipo");
var departamentoSel = document.getElementById("departamentoEquipo");
var departamentoTextoEl = document.getElementById("departamentoTexto");
var sel = document.getElementById("estadoEquipo");
var wrap = document.getElementById("wrapComentarioInactivo");
var ta = document.getElementById("comentarioCambio");
var wrapRep = document.getElementById("wrapComentarioReparacion");
var taRep = document.getElementById("comentarioReparacion");
var wrapMotivoBaja = document.getElementById("wrapMotivoBaja");
var motivoBajaInput = document.getElementById("motivoBaja");
var motivoWrap = document.getElementById("motivoDisplayWrap");
var motivoTexto = document.getElementById("motivoTexto");
var motivoBajaWrap = document.getElementById("motivoBajaDisplayWrap");
var motivoBajaTexto = document.getElementById("motivoBajaTexto");
var tabTransferencias = document.getElementById("tabTransferencias");
var tabEstados = document.getElementById("tabEstados");
var tabUbicacion = document.getElementById("tabUbicacion");
var panelTransferencias = document.getElementById("panelTransferencias");
var panelEstados = document.getElementById("panelEstados");
var panelUbicacion = document.getElementById("panelUbicacion");
var listaTransferencias = document.getElementById("listaTransferencias");
var listaEstados = document.getElementById("listaEstados");
var listaUbicacion = document.getElementById("listaUbicacion");
var countTransferencias = document.getElementById("countTransferencias");
var countEstados = document.getElementById("countEstados");
var countUbicacion = document.getElementById("countUbicacion");
var usuarioActual = ${JSON.stringify(row.usuario || "")};
var estadoInicialEquipo = ${JSON.stringify(row.estado || "Activo")};
var NOMBRE_AREA = ${JSON.stringify(NOMBRE_AREA_POR_CODIGO)};
function nombreAreaFromCodigo(c){
if (c == null) return "-";
var t = String(c).trim();
if (t === "" || t === "-") return t;
var u = t.toUpperCase();
return NOMBRE_AREA[u] || t;
}
function humanizarComentarioUbicacion(s){
var t = String(s == null ? "" : s).trim();
var mp = /^Planta:\\s*(.+?)\\s*→\\s*(.+)$/.exec(t);
if (mp) return "Planta: " + mp[1].trim() + " → " + mp[2].trim();
var m = /^Departamento:\\s*(.+?)\\s*→\\s*(.+)$/.exec(t);
if (!m) m = /^Área:\\s*(.+?)\\s*→\\s*(.+)$/.exec(t);
if (!m) return String(s == null ? "" : s);
return "Área: " + nombreAreaFromCodigo(m[1].trim()) + " → " + nombreAreaFromCodigo(m[2].trim());
}

function escapeText(s){
return String(s == null ? "" : s)
.replace(/&/g, "&amp;")
.replace(/</g, "&lt;")
.replace(/>/g, "&gt;")
.replace(/"/g, "&quot;");
}

function etiquetaEstadoMostrar(e){
if (e === "Eliminado") return "De baja";
return e == null ? "" : String(e);
}

function activarTab(tab){
tabTransferencias.classList.toggle("activo", tab === "transferencias");
tabEstados.classList.toggle("activo", tab === "estados");
tabUbicacion.classList.toggle("activo", tab === "ubicacion");
panelTransferencias.classList.toggle("activo", tab === "transferencias");
panelEstados.classList.toggle("activo", tab === "estados");
panelUbicacion.classList.toggle("activo", tab === "ubicacion");
}

function renderTransferencia(item){
var comentario = item.comentario ? " | Comentario: " + escapeText(item.comentario) : "";
return "<li><b>" + escapeText(item.usuario_anterior || "Sin usuario") + "</b> → <b>" + escapeText(item.usuario_nuevo || "") + "</b> | " + escapeText(item.fecha_cambio) + comentario + "</li>";
}

function renderEstado(item){
var comentario = "";
if (item.comentario) {
var c = escapeText(item.comentario);
comentario =
item.estado_nuevo === "Inactivo" ? " | Motivo: " + c : " | " + c;
}
var ant = etiquetaEstadoMostrar(item.estado_anterior);
var nuevo = etiquetaEstadoMostrar(item.estado_nuevo);
return "<li><b>" + escapeText(ant) + "</b> → <b>" + escapeText(nuevo) + "</b> | " + escapeText(item.fecha_cambio) + comentario + "</li>";
}

function renderUbicacion(item){
var detalle = item.comentario ? escapeText(humanizarComentarioUbicacion(item.comentario)) : "";
return "<li>" + detalle + " | " + escapeText(item.fecha_cambio) + "</li>";
}

function cargarHistorialCambios(){
Promise.all([
fetch("/equipo/" + encodeURIComponent(codigoEquipo) + "/historial-cambios?tipo=transferencia", { credentials: "include", headers: { Accept: "application/json" } }).then(function(res){
return res.ok ? res.json() : [];
}),
fetch("/equipo/" + encodeURIComponent(codigoEquipo) + "/historial-cambios?tipo=estado", { credentials: "include", headers: { Accept: "application/json" } }).then(function(res){
return res.ok ? res.json() : [];
}),
fetch("/equipo/" + encodeURIComponent(codigoEquipo) + "/historial-cambios?tipo=ubicacion", { credentials: "include", headers: { Accept: "application/json" } }).then(function(res){
return res.ok ? res.json() : [];
})
])
.then(function(results){
var transferencias = Array.isArray(results[0]) ? results[0] : [];
var estados = Array.isArray(results[1]) ? results[1] : [];
var ubicaciones = Array.isArray(results[2]) ? results[2] : [];
countTransferencias.textContent = String(transferencias.length);
countEstados.textContent = String(estados.length);
if (countUbicacion) countUbicacion.textContent = String(ubicaciones.length);
listaTransferencias.innerHTML = transferencias.length
? transferencias.map(renderTransferencia).join("")
: "<li>Sin transferencias registradas.</li>";
listaEstados.innerHTML = estados.length
? estados.map(renderEstado).join("")
: "<li>Sin cambios de estado registrados.</li>";
if (listaUbicacion) {
listaUbicacion.innerHTML = ubicaciones.length
? ubicaciones.map(renderUbicacion).join("")
: "<li>Sin cambios de área registrados.</li>";
}
})
.catch(function(){
listaTransferencias.innerHTML = "<li>No se pudo cargar el historial.</li>";
listaEstados.innerHTML = "<li>No se pudo cargar el historial.</li>";
if (listaUbicacion) listaUbicacion.innerHTML = "<li>No se pudo cargar el historial.</li>";
});
}

function syncComentarioUI(){
var esInactivo = sel.value === "Inactivo";
var esEliminado = sel.value === "Eliminado";
var esActivo = sel.value === "Activo";
wrap.style.display = esInactivo ? "block" : "none";
wrapMotivoBaja.style.display = esEliminado ? "block" : "none";
var mostrarReparacion = esActivo && estadoInicialEquipo === "Inactivo";
if (wrapRep && taRep) {
wrapRep.style.display = mostrarReparacion ? "block" : "none";
if (!mostrarReparacion) {
taRep.value = "";
}
}
if (!esInactivo && !esEliminado) {
ta.value = "";
}
if (!esEliminado) {
motivoBajaInput.value = "";
}
}

sel.addEventListener("change", function(){
msg.textContent = "";
syncComentarioUI();
});

syncComentarioUI();

function enviarCambiosEquipo(){
var usuario = (usuarioInput.value || "").trim();
var estado = sel.value;
var comentario = (ta.value || "").trim();
var motivoBaja = (motivoBajaInput.value || "").trim();
var comentarioEstado = estado === "Inactivo" ? comentario : "";
var comentarioReparacion = taRep ? (taRep.value || "").trim() : "";
var payload = {
usuario: usuario,
estado: estado,
comentario: comentarioEstado,
motivo_baja: estado === "Eliminado" ? motivoBaja : "",
comentario_reactivacion:
  estado === "Activo" && estadoInicialEquipo === "Inactivo" ? comentarioReparacion : "",
departamento: departamentoSel ? departamentoSel.value : ""
};
fetch("/equipo/" + encodeURIComponent(codigoEquipo) + "/cambios", {
method: "PATCH",
credentials: "include",
headers: { "Content-Type": "application/json", Accept: "application/json" },
body: JSON.stringify(payload)
})
.then(function(res){ return res.json().then(function(data){ return { res: res, data: data }; }); })
.then(function(r){
if (!r.res.ok) {
var errText = (r.data && r.data.error) || "No se pudieron guardar los cambios";
msg.textContent = errText;
msg.style.color = "#c0392b";
if (typeof Swal !== "undefined") {
Swal.fire({
icon: "error",
title: "No se pudo guardar",
text: errText,
confirmButtonColor: "#262626"
});
}
return;
}
var d = r.data;
usuarioActual = d.usuario;
if (usuarioTexto) usuarioTexto.textContent = d.usuario;
if (departamentoTextoEl) {
var dn = d.departamento_nombre != null ? d.departamento_nombre : nombreAreaFromCodigo(d.departamento);
departamentoTextoEl.textContent = dn;
}
if (span) span.textContent = etiquetaEstadoMostrar(d.estado);
if (d.estado === "Activo") {
if (motivoWrap) motivoWrap.style.display = "none";
if (motivoTexto) motivoTexto.textContent = "";
if (motivoBajaWrap) motivoBajaWrap.style.display = "none";
if (motivoBajaTexto) motivoBajaTexto.textContent = "";
} else {
if (d.estado === "Inactivo") {
if (motivoWrap) motivoWrap.style.display = "block";
if (motivoTexto) motivoTexto.textContent = d.comentario_inactivo || "";
if (motivoBajaWrap) motivoBajaWrap.style.display = "none";
if (motivoBajaTexto) motivoBajaTexto.textContent = "";
} else if (d.estado === "Eliminado") {
if (motivoWrap) motivoWrap.style.display = "none";
if (motivoTexto) motivoTexto.textContent = "";
if (motivoBajaWrap) motivoBajaWrap.style.display = "block";
if (motivoBajaTexto) motivoBajaTexto.textContent = d.motivo_baja || "";
} else {
if (motivoWrap) motivoWrap.style.display = "none";
if (motivoTexto) motivoTexto.textContent = "";
if (motivoBajaWrap) motivoBajaWrap.style.display = "none";
if (motivoBajaTexto) motivoBajaTexto.textContent = "";
}
}
if (d.comentario_inactivo) {
ta.value = d.comentario_inactivo;
}
estadoInicialEquipo = d.estado || estadoInicialEquipo;
if (taRep) taRep.value = "";
syncComentarioUI();
msg.textContent = d.registros_historial > 0 ? "Cambios guardados correctamente." : "No hubo cambios para guardar.";
msg.style.color = "#1e7e34";
if (d.estado === "Eliminado") {
if (typeof Swal !== "undefined") {
Swal.fire({
icon: "success",
title: "Equipo dado de baja",
text: "El equipo se registró correctamente.",
confirmButtonColor: "#262626"
}).then(function(){ window.location.href = "/"; });
} else {
window.location.href = "/";
}
return;
}
if (typeof Swal !== "undefined" && d.registros_historial > 0) {
Swal.fire({
icon: "success",
title: "Listo",
text: "Cambios guardados correctamente.",
toast: true,
position: "top-end",
showConfirmButton: false,
timer: 2200,
timerProgressBar: true
});
}
cargarHistorialCambios();
});
}

form.addEventListener("submit", function(e){
e.preventDefault();
msg.textContent = "";
var usuario = (usuarioInput.value || "").trim();
var estado = sel.value;
var comentario = (ta.value || "").trim();
var motivoBaja = (motivoBajaInput.value || "").trim();
var comentarioEstado = estado === "Inactivo" ? comentario : "";
var comentarioReparacion = taRep ? (taRep.value || "").trim() : "";
if (!usuario) {
msg.textContent = "Ingrese el usuario asignado.";
msg.style.color = "#c0392b";
if (typeof Swal !== "undefined") {
Swal.fire({
icon: "warning",
title: "Falta información",
text: "Ingrese el usuario asignado.",
confirmButtonColor: "#262626"
});
}
return;
}
if (estado === "Inactivo" && !comentarioEstado) {
msg.textContent = "Ingrese el motivo para marcar el equipo como Inactivo.";
msg.style.color = "#c0392b";
if (typeof Swal !== "undefined") {
Swal.fire({
icon: "warning",
title: "Motivo requerido",
text: "Ingrese el motivo para marcar el equipo como Inactivo.",
confirmButtonColor: "#262626"
});
}
return;
}
if (estado === "Activo" && estadoInicialEquipo === "Inactivo" && !comentarioReparacion) {
msg.textContent = "Indique qué se reparó del equipo para volver a Activo.";
msg.style.color = "#c0392b";
if (typeof Swal !== "undefined") {
Swal.fire({
icon: "warning",
title: "Comentario requerido",
text: "Describa qué se reparó del equipo al pasar de Inactivo a Activo.",
confirmButtonColor: "#262626"
});
}
return;
}
if (estado === "Eliminado" && !motivoBaja) {
msg.textContent = "Ingrese el motivo para dar de baja el equipo.";
msg.style.color = "#c0392b";
if (typeof Swal !== "undefined") {
Swal.fire({
icon: "warning",
title: "Motivo requerido",
text: "Ingrese el motivo para dar de baja el equipo.",
confirmButtonColor: "#262626"
});
}
return;
}
if (estado === "Eliminado") {
if (typeof Swal !== "undefined") {
Swal.fire({
title: "¿Está seguro?",
text: "¿Está seguro que desea dar de baja este equipo?",
icon: "question",
showCancelButton: true,
confirmButtonText: "Sí, dar de baja",
cancelButtonText: "Cancelar",
confirmButtonColor: "#262626",
cancelButtonColor: "#6c757d"
}).then(function(result){
if (result.isConfirmed) {
enviarCambiosEquipo();
}
});
} else {
if (window.confirm("¿Está seguro que desea dar de baja este equipo?")) {
enviarCambiosEquipo();
}
}
return;
}
enviarCambiosEquipo();
});

tabTransferencias.addEventListener("click", function(){
activarTab("transferencias");
});

tabEstados.addEventListener("click", function(){
activarTab("estados");
});

if (tabUbicacion) {
tabUbicacion.addEventListener("click", function(){
activarTab("ubicacion");
});
}

cargarHistorialCambios();
})();
</script>

</body>

</html>

`)

})

})

/* ========================
   ACTUALIZAR ESTADO EQUIPO
======================== */

const ESTADOS_EDICION = ["Activo", "Inactivo", "Eliminado"]

const DEPARTAMENTOS_EDICION = ["TI", "BOD", "FRI", "RH", "GER", "ADM", "PAC", "MAN"]

function registrarHistorialEquipo(entry, done) {
  db.run(
    `INSERT INTO historial_equipos
     (equipo_codigo, usuario_anterior, usuario_nuevo, estado_anterior, estado_nuevo, comentario, tipo_cambio)
     VALUES (?, ?, ?, ?, ?, ?, ?)`,
    [
      entry.equipo_codigo,
      entry.usuario_anterior || "",
      entry.usuario_nuevo || "",
      entry.estado_anterior || "",
      entry.estado_nuevo || "",
      entry.comentario || null,
      entry.tipo_cambio
    ],
    done
  )
}

app.patch("/equipo/:codigo/cambios", requireAuth, (req, res) => {
  const codigo = codigoEquipoDesdeParams(req)
  if (!codigo) {
    return res.status(400).json({ error: "Codigo de equipo invalido" })
  }
  const usuarioNuevoRaw =
    req.body && typeof req.body.usuario === "string" ? req.body.usuario : ""
  const nuevoEstado = req.body && typeof req.body.estado === "string" ? req.body.estado : ""
  const comentarioRaw =
    req.body && typeof req.body.comentario === "string" ? req.body.comentario : ""
  const motivoBajaRaw =
    req.body && typeof req.body.motivo_baja === "string" ? req.body.motivo_baja : ""
  const reactivacionRaw =
    req.body && typeof req.body.comentario_reactivacion === "string"
      ? req.body.comentario_reactivacion
      : ""
  const usuarioNuevo = usuarioNuevoRaw.trim()
  const comentario = comentarioRaw.trim()
  const motivoBaja = motivoBajaRaw.trim()
  const comentarioReactivacion = reactivacionRaw.trim()
  const comentarioEstado = nuevoEstado === "Inactivo" ? comentario : ""
  const departamentoRaw =
    req.body && typeof req.body.departamento === "string" ? req.body.departamento : ""
  const departamentoNuevo = departamentoRaw.trim().toUpperCase()

  if (!usuarioNuevo) {
    return res.status(400).json({ error: "Debe indicar el usuario del equipo" })
  }

  if (!DEPARTAMENTOS_EDICION.includes(departamentoNuevo)) {
    return res.status(400).json({ error: "Seleccione un área válida" })
  }
  if (!ESTADOS_EDICION.includes(nuevoEstado)) {
    return res.status(400).json({ error: "Estado debe ser Activo, Inactivo o Eliminado" })
  }

  if (nuevoEstado === "Inactivo" && !comentarioEstado) {
    return res.status(400).json({
      error: "Debe ingresar el motivo al marcar el equipo como Inactivo"
    })
  }
  if (nuevoEstado === "Eliminado" && !motivoBaja) {
    return res.status(400).json({
      error: "Debe ingresar el motivo para dar de baja el equipo"
    })
  }

  const comentarioDb = nuevoEstado === "Inactivo" ? comentarioEstado : null
  const motivoBajaDb = nuevoEstado === "Eliminado" ? motivoBaja : null

  db.get("SELECT * FROM equipos WHERE codigo = ?", [codigo], (err, row) => {
    if (err) {
      return res.status(500).json({ error: "Error al buscar equipo" })
    }
    if (!row) {
      return res.status(404).json({ error: "Equipo no encontrado" })
    }

    const usuarioAnterior = (row.usuario || "").trim()
    const estadoAnterior = row.estado || "Activo"
    const departamentoAnterior = (row.departamento || "").trim().toUpperCase()
    const cambioUsuario = usuarioAnterior !== usuarioNuevo
    const cambioEstado = estadoAnterior !== nuevoEstado
    const cambioDepartamento = departamentoAnterior !== departamentoNuevo

    if (
      estadoAnterior === "Inactivo" &&
      nuevoEstado === "Activo" &&
      !comentarioReactivacion
    ) {
      return res.status(400).json({
        error: "Debe describir qué se reparó del equipo al volver a Activo"
      })
    }

    if (!cambioUsuario && !cambioEstado && !cambioDepartamento) {
      const dep = departamentoAnterior || row.departamento || ""
      return res.json({
        ok: true,
        codigo,
        usuario: usuarioAnterior,
        estado: estadoAnterior,
        comentario_inactivo: row.comentario_inactivo || null,
        motivo_baja: row.motivo_baja || null,
        departamento: dep,
        departamento_nombre: nombreAreaMostrar(dep),
        registros_historial: 0
      })
    }

    db.run(
      "UPDATE equipos SET usuario = ?, estado = ?, comentario_inactivo = ?, motivo_baja = ?, departamento = ? WHERE codigo = ?",
      [
        usuarioNuevo,
        nuevoEstado,
        comentarioDb,
        motivoBajaDb,
        departamentoNuevo,
        codigo
      ],
      function (updateErr) {
        if (updateErr) {
          if (esErrorSqliteBusy(updateErr)) return responderBloqueoBd(res)
          return res.status(500).json({ error: "Error al actualizar equipo" })
        }
        if (this.changes === 0) {
          return res.status(404).json({ error: "Equipo no encontrado" })
        }

        const tareas = []
        if (cambioDepartamento) {
          tareas.push((next) => {
            registrarHistorialEquipo(
              {
                equipo_codigo: codigo,
                usuario_anterior: "",
                usuario_nuevo: "",
                estado_anterior: "",
                estado_nuevo: "",
                comentario: `Área: ${nombreAreaMostrar(departamentoAnterior) || "-"} → ${nombreAreaMostrar(departamentoNuevo)}`,
                tipo_cambio: "ubicacion"
              },
              next
            )
          })
          db.run("INSERT INTO historial (equipo_codigo, cambio) VALUES (?, ?)", [
            codigo,
            `Área: ${nombreAreaMostrar(departamentoAnterior) || "-"} → ${nombreAreaMostrar(departamentoNuevo)}`
          ])
        }
        if (cambioUsuario) {
          tareas.push((next) => {
            registrarHistorialEquipo(
              {
                equipo_codigo: codigo,
                usuario_anterior: usuarioAnterior,
                usuario_nuevo: usuarioNuevo,
                comentario: null,
                tipo_cambio: "transferencia"
              },
              next
            )
          })
          db.run("INSERT INTO historial (equipo_codigo, cambio) VALUES (?, ?)", [
            codigo,
            `Usuario: ${usuarioAnterior || "Sin usuario"} → ${usuarioNuevo}`
          ])
        }

        if (cambioEstado) {
          tareas.push((next) => {
            registrarHistorialEquipo(
              {
                equipo_codigo: codigo,
                estado_anterior: estadoAnterior,
                estado_nuevo: nuevoEstado,
                comentario:
                  nuevoEstado === "Inactivo"
                    ? comentarioEstado || null
                    : nuevoEstado === "Eliminado"
                      ? motivoBaja || null
                      : nuevoEstado === "Activo" && estadoAnterior === "Inactivo"
                        ? `Reparación: ${comentarioReactivacion}`
                        : null,
                tipo_cambio: "estado"
              },
              next
            )
          })
          db.run("INSERT INTO historial (equipo_codigo, cambio) VALUES (?, ?)", [
            codigo,
            `Estado: ${estadoAnterior} → ${nuevoEstado}${
              nuevoEstado === "Inactivo" && comentarioEstado
                ? `. Motivo: ${comentarioEstado}`
                : nuevoEstado === "Eliminado" && motivoBaja
                  ? `. Motivo baja: ${motivoBaja}`
                  : nuevoEstado === "Activo" &&
                      estadoAnterior === "Inactivo" &&
                      comentarioReactivacion
                    ? `. Reparación: ${comentarioReactivacion}`
                    : ""
            }`
          ])
        }
        if (cambioEstado && nuevoEstado === "Eliminado") {
          tareas.push((next) => {
            db.run(`DELETE FROM equipo_notas WHERE equipo_codigo = ?`, [codigo], (delErr) => {
              next(delErr || null)
            })
          })
        }

        let completadas = 0
        const total = tareas.length
        if (total === 0) {
          return res.json({
            ok: true,
            codigo,
            usuario: usuarioNuevo,
            estado: nuevoEstado,
            comentario_inactivo: comentarioDb,
            motivo_baja: motivoBajaDb,
            departamento: departamentoNuevo,
            departamento_nombre: nombreAreaMostrar(departamentoNuevo),
            registros_historial: total
          })
        }

        function doneOne(taskErr) {
          if (taskErr) {
            if (esErrorSqliteBusy(taskErr)) return responderBloqueoBd(res)
            return res
              .status(500)
              .json({ error: "Se actualizó el equipo pero falló el historial de cambios" })
          }
          completadas += 1
          if (completadas === total) {
            res.json({
              ok: true,
              codigo,
              usuario: usuarioNuevo,
              estado: nuevoEstado,
              comentario_inactivo: comentarioDb,
              motivo_baja: motivoBajaDb,
              departamento: departamentoNuevo,
              departamento_nombre: nombreAreaMostrar(departamentoNuevo),
              registros_historial: total
            })
          }
        }

        tareas.forEach((fn) => fn(doneOne))
      }
    )
  })
})

// Compatibilidad con flujo antiguo: solo cambio de estado
app.patch("/equipo/:codigo", requireAuth, (req, res) => {
  const codigo = codigoEquipoDesdeParams(req)
  if (!codigo) {
    return res.status(400).json({ error: "Codigo de equipo invalido" })
  }
  const nuevoEstado = req.body && typeof req.body.estado === "string" ? req.body.estado : ""
  const comentarioRaw =
    req.body && typeof req.body.comentario === "string" ? req.body.comentario : ""
  const comentario = comentarioRaw.trim()

  const ESTADOS_COMPAT = ["Activo", "Inactivo"]
  if (!ESTADOS_COMPAT.includes(nuevoEstado)) {
    return res.status(400).json({ error: "Estado debe ser Activo o Inactivo" })
  }
  if (nuevoEstado === "Inactivo" && !comentario) {
    return res.status(400).json({
      error: "Debe ingresar el motivo al marcar el equipo como Inactivo"
    })
  }

  db.get(
    "SELECT estado, comentario_inactivo FROM equipos WHERE codigo = ?",
    [codigo],
    (err, row) => {
    if (err) {
      return res.status(500).json({ error: "Error al buscar equipo" })
    }
    if (!row) {
      return res.status(404).json({ error: "Equipo no encontrado" })
    }

    const estadoAnterior = row.estado || "Activo"
    const comentarioDb = nuevoEstado === "Inactivo" ? comentario : null

    db.run(
      "UPDATE equipos SET estado = ?, comentario_inactivo = ? WHERE codigo = ?",
      [nuevoEstado, comentarioDb, codigo],
      function (updateErr) {
        if (updateErr) {
          if (esErrorSqliteBusy(updateErr)) return responderBloqueoBd(res)
          return res.status(500).json({ error: "Error al actualizar estado" })
        }
        if (estadoAnterior !== nuevoEstado) {
          registrarHistorialEquipo(
            {
              equipo_codigo: codigo,
              estado_anterior: estadoAnterior,
              estado_nuevo: nuevoEstado,
              comentario: comentario || null,
              tipo_cambio: "estado"
            },
            () => {}
          )
          db.run("INSERT INTO historial (equipo_codigo, cambio) VALUES (?, ?)", [
            codigo,
            `Estado: ${estadoAnterior} → ${nuevoEstado}${comentario ? `. Motivo: ${comentario}` : ""}`
          ])
        }
        res.json({
          ok: true,
          estado: nuevoEstado,
          comentario_inactivo: comentarioDb
        })
      }
    )
  })
})

// Compatibilidad con flujo antiguo: transferencia separada
app.patch("/equipo/:codigo/transferir", requireAuth, (req, res) => {
  const codigo = codigoEquipoDesdeParams(req)
  if (!codigo) {
    return res.status(400).json({ error: "Codigo de equipo invalido" })
  }
  const usuarioNuevoRaw =
    req.body && typeof req.body.usuario_nuevo === "string"
      ? req.body.usuario_nuevo
      : ""
  const usuarioNuevo = usuarioNuevoRaw.trim()

  if (!usuarioNuevo) {
    return res.status(400).json({ error: "Debe indicar el nuevo usuario" })
  }

  db.get("SELECT usuario FROM equipos WHERE codigo = ?", [codigo], (err, row) => {
    if (err) {
      return res.status(500).json({ error: "Error al buscar equipo" })
    }
    if (!row) {
      return res.status(404).json({ error: "Equipo no encontrado" })
    }

    const usuarioAnterior = (row.usuario || "").trim()
    if (usuarioAnterior === usuarioNuevo) {
      return res.status(400).json({
        error: "El nuevo usuario debe ser distinto al usuario actual"
      })
    }

    db.run("UPDATE equipos SET usuario=? WHERE codigo=?", [usuarioNuevo, codigo], function (updateErr) {
      if (updateErr) {
        return res.status(500).json({ error: "Error al transferir equipo" })
      }

      registrarHistorialEquipo(
        {
          equipo_codigo: codigo,
          usuario_anterior: usuarioAnterior,
          usuario_nuevo: usuarioNuevo,
          comentario: null,
          tipo_cambio: "transferencia"
        },
        () => {}
      )
      db.run("INSERT INTO historial (equipo_codigo, cambio) VALUES (?, ?)", [
        codigo,
        `Usuario: ${usuarioAnterior || "Sin usuario"} → ${usuarioNuevo}`
      ])

      res.json({
        ok: true,
        codigo,
        usuario_anterior: usuarioAnterior,
        usuario_nuevo: usuarioNuevo,
        comentario: null
      })
    })
  })
})

app.get("/equipo/:codigo/historial-cambios", requireAuth, (req, res) => {
  const codigo = codigoEquipoDesdeParams(req)
  if (!codigo) {
    return res.status(400).json({ error: "Codigo de equipo invalido" })
  }
  const tipoRaw = typeof req.query.tipo === "string" ? req.query.tipo.trim() : ""
  const tipo =
    tipoRaw === "transferencia" || tipoRaw === "estado" || tipoRaw === "ubicacion"
      ? tipoRaw
      : ""

  db.get("SELECT codigo FROM equipos WHERE codigo = ?", [codigo], (gErr, eqRow) => {
    if (gErr) {
      return res.status(500).json({ error: "Error al verificar el equipo" })
    }
    if (!eqRow) {
      return res.status(404).json({ error: "Equipo no encontrado" })
    }

    let sql = `SELECT
       id,
       equipo_codigo,
       usuario_anterior,
       usuario_nuevo,
       estado_anterior,
       estado_nuevo,
       comentario,
       tipo_cambio,
       fecha_cambio
     FROM historial_equipos
     WHERE equipo_codigo = ?`
    const params = [codigo]

    if (tipo) {
      sql += ` AND tipo_cambio = ?`
      params.push(tipo)
    }

    sql += ` ORDER BY fecha_cambio DESC, id DESC`

    db.all(sql, params, (err, rows) => {
      if (err) {
        return res.status(500).json({ error: "Error al cargar historial de cambios" })
      }
      const out = (rows || []).map((r) => ({
        ...r,
        fecha_cambio: formatFechaLocal(r.fecha_cambio)
      }))
      res.json(out)
    })
  })
})

/* ========================
   EXPORTAR EXCEL
======================== */

app.get("/exportar", requireAuth, (req, res) => {
  db.all(
    "SELECT * FROM equipos ORDER BY datetime(fecha_creacion) DESC, codigo DESC",
    [],
    (err, rows) => {
    if (err) {
      return res.status(500).send("Error al exportar inventario")
    }
    buildInventarioExcel(rows || [])
      .then((buffer) => {
        const now = new Date()
        const stamp = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, "0")}${String(now.getDate()).padStart(2, "0")}-${String(now.getHours()).padStart(2, "0")}${String(now.getMinutes()).padStart(2, "0")}${String(now.getSeconds()).padStart(2, "0")}`
        const filename = `inventario-ti-${stamp}.xlsx`
        res.setHeader(
          "Content-Disposition",
          `attachment; filename="${filename}"`
        )
        res.setHeader(
          "Content-Type",
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        res.setHeader("Cache-Control", "no-store, no-cache, must-revalidate, private")
        res.setHeader("Pragma", "no-cache")
        res.setHeader("Expires", "0")
        res.send(Buffer.from(buffer))
      })
      .catch((e) => {
        console.error("exportar excel:", e)
        res.status(500).send("Error al generar el archivo Excel")
      })
  })
})

/* ========================
   ETIQUETA (impresión)
======================== */

app.get("/etiqueta/:codigo", requireAuth, (req, res) => {
  const codigoParam = codigoEquipoDesdeParams(req)
  const codigoQuery =
    typeof req.query.c === "string" ? req.query.c.trim() : ""

  if (!codigoParam) {
    return res.status(400).send("Codigo de equipo invalido")
  }

  db.get("SELECT * FROM equipos WHERE codigo = ?", [codigoParam], (err, row) => {
    if (err) {
      return res.status(500).send("Error al cargar el equipo")
    }
    if (!row) {
      return res.send("Equipo no encontrado")
    }

    const codigoDb = codigoEquipoMostrar(row)
    const codigoMostrar = codigoDb || codigoQuery || ""
    const codigoTxt = escapeHtml(codigoMostrar)

res.send(`

<html>

<head>

<style>

body{
font-family:Arial,sans-serif;
display:flex;
justify-content:center;
align-items:center;
min-height:100vh;
background:#f4f6f9;
margin:0;
}

.vista-etiqueta{
display:flex;
flex-direction:column;
align-items:center;
gap:12px;
}

.etiqueta{
width:50mm;
height:25mm;
border:2px solid #000;
padding:2mm;
box-sizing:border-box;
background:white;
text-align:center;
display:flex;
align-items:center;
justify-content:center;
overflow:visible;
}

.codigo{
font-size:12mm;
font-weight:bold;
line-height:1.1;
color:#000;
-webkit-print-color-adjust:exact;
print-color-adjust:exact;
word-break:break-all;
}

.acciones{
text-align:center;
}

@media print{
@page{
size:50mm 25mm;
margin:0;
}

html,body{
width:50mm;
height:25mm;
margin:0;
padding:0;
background:#fff !important;
}

body{
display:block;
}

.vista-etiqueta{
display:block;
width:50mm;
height:25mm;
margin:0;
padding:0;
gap:0;
}

.etiqueta{
width:50mm;
height:25mm;
border:none;
padding:1mm;
margin:0;
box-sizing:border-box;
}

.acciones{
display:none !important;
}
}

</style>

</head>

<body>

<div class="vista-etiqueta">

<div class="etiqueta">

<div class="codigo">${codigoTxt || "—"}</div>

</div>

<div class="acciones">
<button type="button" onclick="window.print()">🖨️ Imprimir</button>
</div>

</div>

</body>

</html>

`)

  })
})

/* ========================
   SERVIDOR
======================== */

/** iisnode suele definir process.env.PORT como named pipe (no solo un número). */
const listenTarget = process.env.PORT || process.env.IISNODE_HTTPPORT || 3001
const esPuertoNumerico =
  typeof listenTarget === "number" ||
  (typeof listenTarget === "string" && /^\d+$/.test(String(listenTarget).trim()))

if (esPuertoNumerico) {
  const puertoNum = Number(listenTarget) || 3001
  app.listen(puertoNum, "0.0.0.0", () => {
    console.log("Servidor corriendo en puerto", puertoNum)
  })
} else {
  app.listen(listenTarget, () => {
    console.log("Servidor bajo IIS iisnode:", String(listenTarget))
  })
}