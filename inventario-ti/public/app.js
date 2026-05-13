const fetchOptsJson = {
  credentials: "include",
  headers: { Accept: "application/json" }
}

async function bootstrapSesion() {
  const res = await fetch("/api/session", fetchOptsJson)
  if (!res.ok) {
    window.location.href = "/login"
    return null
  }
  const data = await res.json()
  if (!data.authenticated) {
    window.location.href = "/login"
    return null
  }
  const el = document.getElementById("sessionUserLabel")
  if (el) {
    const u = data.user || {}
    el.textContent = u.esSuperadmin ? "SuperAdmin" : (u.nombre || "")
  }
  sesionUsuario = data.user || null
  return data.user
}

function configurarCerrarSesion() {
  const btn = document.getElementById("btnCerrarSesion")
  if (!btn) return
  btn.addEventListener("click", async () => {
    try {
      await fetch("/logout", { method: "POST", credentials: "include" })
    } catch {
      // ignorar
    }
    window.location.href = "/login"
  })
}

function redirigirSi401(res) {
  if (res.status === 401) {
    window.location.href = "/login"
    return true
  }
  return false
}

function mostrarAvisoBloqueoBdSiAplica(res, body, msgEl) {
  if (!res || res.status !== 503) return false
  const texto =
    (body && body.error) ||
    "La base de datos está en uso por otra aplicación. Cierre el DB e intente nuevamente."
  if (msgEl) msgEl.textContent = texto
  if (typeof Swal !== "undefined") {
    Swal.fire({
      icon: "warning",
      title: "Base de datos en uso",
      text: texto,
      confirmButtonColor: "#0B47BF"
    })
  }
  return true
}

let sesionUsuario = null

async function cargarListaUsuarios() {
  const res = await fetch("/api/usuarios", {
    credentials: "include",
    headers: { Accept: "application/json" }
  })
  if (redirigirSi401(res)) {
    return
  }
  if (!res.ok) {
    return
  }
  const rows = await res.json()
  renderTablaUsuarios(rows)
}

function renderTablaUsuarios(rows) {
  const tbody = document.querySelector("#tablaUsuarios tbody")
  if (!tbody) {
    return
  }
  tbody.innerHTML = ""
  ;(rows || []).forEach((u) => {
    const activo = u.activo == null || Number(u.activo) === 1
    const superRol = Number(u.es_superadmin) === 1
    const tr = document.createElement("tr")
    const btn = document.createElement("button")
    btn.type = "button"
    btn.className = "btn-editar-usuario"
    btn.textContent = "Editar"
    btn.addEventListener("click", () => {
      abrirDialogoEditarUsuario({
        id: String(u.id),
        nombre: u.nombre,
        activo,
        es_superadmin: superRol
      })
    })
    tr.innerHTML = `
<td>${escapeHtmlUser(u.nombre)}</td>
<td>${
  activo
    ? '<span class="badge-activo">Activo</span>'
    : '<span class="badge-inactivo">Desactivado</span>'
}</td>
<td>${
  superRol
    ? '<span class="badge-super">Superadmin</span>'
    : "Usuario"
}</td>
<td></td>
`
    tr.querySelector("td:last-child").appendChild(btn)
    tbody.appendChild(tr)
  })
}

function escapeHtmlUser(s) {
  if (s == null) {
    return ""
  }
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
}

function abrirDialogoEditarUsuario(u) {
  const dlg = document.getElementById("dialogEditarUsuario")
  const msg = document.getElementById("msgEditarUsuario")
  if (!dlg) {
    return
  }
  msg.textContent = ""
  document.getElementById("editUserId").value = u.id
  document.getElementById("editUserNombre").value = u.nombre
  document.getElementById("editUserPassword").value = ""
  document.getElementById("editUserActivo").disabled = false
  document.getElementById("editUserActivo").title = ""
  const selRol = document.getElementById("editUserRol")
  if (selRol) selRol.value = u.es_superadmin ? "1" : "0"
  document.getElementById("editUserActivo").value = u.activo ? "1" : "0"
  dlg.showModal()
}

function initSuperadminUI(user) {
  if (!user || !user.esSuperadmin) {
    return
  }

  const card = document.getElementById("cardAdminUsuarios")
  if (card) {
    card.hidden = false
  }

  const formNuevo = document.getElementById("formNuevoUsuario")
  if (formNuevo) {
    formNuevo.addEventListener("submit", async (e) => {
      e.preventDefault()
      const msgEl = document.getElementById("msgNuevoUsuario")
      msgEl.textContent = ""
      const nombre = document.getElementById("nuevoUserNombre").value.trim()
      const password = document.getElementById("nuevoUserPassword").value
      const es_superadmin =
        document.getElementById("nuevoUserRol").value === "1"
      const res = await fetch("/api/usuarios", {
        method: "POST",
        credentials: "include",
        headers: {
          "Content-Type": "application/json",
          Accept: "application/json"
        },
        body: JSON.stringify({ nombre, password, es_superadmin })
      })
      if (redirigirSi401(res)) {
        return
      }
      let body = {}
      try {
        body = await res.json()
      } catch {
        body = {}
      }
      if (!res.ok) {
        if (mostrarAvisoBloqueoBdSiAplica(res, body, msgEl)) return
        msgEl.textContent = body.error || "No se pudo crear el usuario."
        return
      }
      formNuevo.reset()
      const selRol = document.getElementById("nuevoUserRol")
      if (selRol) selRol.value = "0"
      msgEl.style.color = "#1e7e34"
      msgEl.textContent = "Usuario creado correctamente."
      cargarListaUsuarios()
    })
  }

  const formEdit = document.getElementById("formEditarUsuario")
  const dlg = document.getElementById("dialogEditarUsuario")
  document.getElementById("btnCancelarEditUsuario")?.addEventListener("click", () => {
    dlg?.close()
  })
  if (formEdit && dlg) {
    formEdit.addEventListener("submit", async (e) => {
      e.preventDefault()
      const msgEd = document.getElementById("msgEditarUsuario")
      msgEd.textContent = ""
      const id = document.getElementById("editUserId").value
      const nombre = document.getElementById("editUserNombre").value.trim()
      const password = document.getElementById("editUserPassword").value
      const activo = document.getElementById("editUserActivo").value === "1"
      const es_superadmin =
        document.getElementById("editUserRol").value === "1"
      const payload = { nombre, activo, es_superadmin }
      if (password.length > 0) {
        payload.password = password
      }
      const res = await fetch("/api/usuarios/" + encodeURIComponent(id), {
        method: "PATCH",
        credentials: "include",
        headers: {
          "Content-Type": "application/json",
          Accept: "application/json"
        },
        body: JSON.stringify(payload)
      })
      if (redirigirSi401(res)) {
        return
      }
      let body = {}
      try {
        body = await res.json()
      } catch {
        body = {}
      }
      if (!res.ok) {
        if (mostrarAvisoBloqueoBdSiAplica(res, body, msgEd)) return
        msgEd.textContent = body.error || "No se pudo guardar."
        return
      }
      dlg.close()
      cargarListaUsuarios()
    })
  }

  cargarListaUsuarios()
}

/** Mapa tipo de equipo → marcas disponibles */
const MARCAS_POR_TIPO = {
  Antena: ["Starlink", "Tenda", "Hikvision", "Asus", "Microtik", "Dlink", "Bluecastle", "Linksys", "Dblue"],
  Router: ["Tp Link", "Tenda", "Hikvision", "Asus", "Microtik", "Dlink", "Bluecastle", "Linksys", "Dblue"],
  Repetidores: ["AP", "Tenda", "Hikvision", "Asus", "Microtik", "Dlink", "Bluecastle", "Linksys", "Dblue"],
  Switch: ["Tp Link", "Tenda", "Hikvision", "Asus", "Microtik", "Dlink", "Bluecastle", "Linksys", "Dblue"],
  Notebooks: ["Dell", "HP", "Lenovo", "Asus", "Clio", "Acer", "Gear", "Olidata", "Toshiba"],
  Escritorio: ["Dell", "HP", "Lenovo", "Asus", "Clio", "Acer", "Gear", "Olidata", "Toshiba", "Otros"],
  Impresora: ["Canon", "HP", "Bixolon", "Zebra", "Brother", "TSC"]
}

const PLACEHOLDER_MARCA = "Seleccionar"

/** Código de departamento (BD) → nombre de área para listados (alineado al formulario de alta). */
const AREA_POR_DEPARTAMENTO = {
  TI: "TI",
  BOD: "Bodega",
  FRI: "Frigorifico",
  RH: "RRHH",
  GER: "Gerencia",
  ADM: "Administracion",
  PAC: "Packing",
  MAN: "Mantencion"
}

function etiquetaArea(departamento) {
  if (departamento == null || String(departamento).trim() === "") return "-"
  const cod = String(departamento).trim().toUpperCase()
  return AREA_POR_DEPARTAMENTO[cod] || cod
}

/** Mismo criterio que el servidor (America/Santiago por defecto). */
function anioCompraMaximo() {
  try {
    const parts = new Intl.DateTimeFormat("en-CA", {
      timeZone: "America/Santiago",
      year: "numeric"
    }).formatToParts(new Date())
    const y = parts.find((p) => p.type === "year")
    if (y && y.value) return Number(y.value)
  } catch {
    // ignorar
  }
  return new Date().getFullYear()
}

function etiquetaEstadoMostrar(estado) {
  if (estado === "Eliminado") return "De baja"
  return estado ?? "-"
}

function actualizarSelectMarcas() {
  const tipoSelect = document.getElementById("tipo")
  const marcaSelect = document.getElementById("marca")
  const tipo = tipoSelect.value

  marcaSelect.innerHTML = ""
  const optPlaceholder = document.createElement("option")
  optPlaceholder.value = ""
  optPlaceholder.textContent = PLACEHOLDER_MARCA
  marcaSelect.appendChild(optPlaceholder)

  if (!tipo) {
    marcaSelect.disabled = true
    marcaSelect.value = ""
    return
  }

  marcaSelect.disabled = false
  const marcas = MARCAS_POR_TIPO[tipo] || []
  marcas.forEach((nombre) => {
    const opt = document.createElement("option")
    opt.value = nombre
    opt.textContent = nombre
    marcaSelect.appendChild(opt)
  })
}

function syncSistemaOperativoAlta() {
  const wrap = document.getElementById("wrapSistemaOperativo")
  const so = document.getElementById("sistemaOperativo")
  if (!wrap || !so) return
  const tipo = document.getElementById("tipo").value
  const esPc = tipo === "Notebooks" || tipo === "Escritorio"
  wrap.hidden = !esPc
  so.disabled = !esPc
  if (!esPc) {
    so.value = ""
    so.removeAttribute("required")
  } else {
    so.setAttribute("required", "required")
  }
}

function syncLicenciasAlta() {
  const tipoEl = document.getElementById("tipo")
  const licEl = document.getElementById("licencias")
  if (!tipoEl || !licEl) return
  const tipo = String(tipoEl.value || "").trim()
  const aplicaLicencia = tipo === "Notebooks" || tipo === "Escritorio"
  if (aplicaLicencia) {
    licEl.disabled = false
    return
  }
  licEl.value = "No"
  licEl.disabled = true
}

document.addEventListener("DOMContentLoaded", async () => {
  const user = await bootstrapSesion()
  if (!user) return

  configurarCerrarSesion()
  initSuperadminUI(user)
  cargarEquipos()
  syncComentarioAlta()
  actualizarSelectMarcas()
  syncSistemaOperativoAlta()
  syncLicenciasAlta()

  const anioInput = document.getElementById("anioCompra")
  if (anioInput) {
    const anioMax = anioCompraMaximo()
    anioInput.setAttribute(
      "title",
      `Año de 4 dígitos; no mayor al año en curso (${anioMax})`
    )
    anioInput.addEventListener("input", () => {
      anioInput.value = anioInput.value.replace(/\D/g, "").slice(0, 4)
    })
  }
})

const form = document.getElementById("equipoForm")
const tipoSelect = document.getElementById("tipo")
const estadoSelect = document.getElementById("estado")
const wrapComentario = document.getElementById("wrapComentarioInactivo")
const comentarioField = document.getElementById("comentarioInactivo")
const msgEquipoForm = document.getElementById("msgEquipoForm")

function syncComentarioAlta() {
  const esInactivo = estadoSelect.value === "Inactivo"
  wrapComentario.hidden = !esInactivo
  comentarioField.disabled = !esInactivo
  comentarioField.setAttribute("aria-required", esInactivo ? "true" : "false")
  if (!esInactivo) {
    comentarioField.value = ""
  }
}

tipoSelect.addEventListener("change", () => {
  msgEquipoForm.textContent = ""
  actualizarSelectMarcas()
  syncSistemaOperativoAlta()
  syncLicenciasAlta()
})

estadoSelect.addEventListener("change", () => {
  msgEquipoForm.textContent = ""
  syncComentarioAlta()
})

form.addEventListener("submit", async (e) => {
  e.preventDefault()
  msgEquipoForm.textContent = ""

  const tipo = tipoSelect.value
  const marca = document.getElementById("marca").value.trim()
  const departamento = document.getElementById("departamento").value.trim()
  const modelo = document.getElementById("modelo").value.trim()
  const licencias = document.getElementById("licencias").value.trim()
  const usuario = document.getElementById("usuario").value.trim()

  if (!tipo) {
    msgEquipoForm.textContent = "Seleccione un tipo de equipo."
    return
  }

  if (!marca) {
    msgEquipoForm.textContent = "Seleccione una marca."
    return
  }

  if (!departamento) {
    msgEquipoForm.textContent = "Seleccione un área."
    return
  }

  if (!modelo) {
    msgEquipoForm.textContent = "Ingrese el modelo."
    return
  }
  if (!licencias) {
    msgEquipoForm.textContent = "Seleccione una licencia."
    return
  }
  if (!usuario) {
    msgEquipoForm.textContent = "Ingrese el usuario."
    return
  }

  const anioCompra = document.getElementById("anioCompra").value.trim()
  if (!/^\d{4}$/.test(anioCompra)) {
    msgEquipoForm.textContent = "Ingrese el año de compra con exactamente 4 dígitos."
    return
  }

  const anioMax = anioCompraMaximo()
  const anioNum = Number(anioCompra, 10)
  if (anioNum > anioMax) {
    msgEquipoForm.textContent = `El año de compra no puede ser mayor a ${anioMax} (año en curso).`
    return
  }

  const esPc = tipo === "Notebooks" || tipo === "Escritorio"
  if (!esPc && licencias !== "No") {
    msgEquipoForm.textContent =
      "La licencia solo corresponde a Notebooks o Escritorio."
    return
  }
  const sistemaOperativo = (document.getElementById("sistemaOperativo").value || "").trim()
  if (esPc && !sistemaOperativo) {
    msgEquipoForm.textContent =
      "Seleccione el sistema operativo (Notebooks y Escritorio)."
    return
  }
  if (!esPc && sistemaOperativo) {
    msgEquipoForm.textContent =
      "El sistema operativo solo corresponde a Notebooks o Escritorio."
    return
  }

  const estado = estadoSelect.value
  const comentario = comentarioField.disabled ? "" : comentarioField.value.trim()

  if (estado === "Inactivo" && !comentario) {
    msgEquipoForm.textContent =
      "Debe ingresar el motivo al marcar el equipo como Inactivo."
    return
  }

  const data = {
    marca,
    departamento,
    modelo,
    tipo,
    licencias,
    usuario,
    anio_compra: anioCompra,
    sistema_operativo: esPc ? sistemaOperativo : "",
    estado,
    comentario: estado === "Inactivo" ? comentario : ""
  }

  const res = await fetch("/equipos", {
    method: "POST",
    credentials: "include",
    headers: {
      "Content-Type": "application/json",
      Accept: "application/json"
    },
    body: JSON.stringify(data)
  })

  if (redirigirSi401(res)) {
    return
  }

  let body = {}
  try {
    body = await res.json()
  } catch {
    body = {}
  }

  if (!res.ok) {
    if (mostrarAvisoBloqueoBdSiAplica(res, body, msgEquipoForm)) return
    msgEquipoForm.textContent =
      body.error || "No se pudo crear el equipo. Intente de nuevo."
    return
  }

  form.reset()
  msgEquipoForm.style.color = "#1e7e34"
  const codigoCreado = body.codigo
  msgEquipoForm.textContent = codigoCreado
    ? `Equipo creado con codigo ${codigoCreado}.`
    : "Equipo creado correctamente."
  actualizarSelectMarcas()
  syncComentarioAlta()
  syncSistemaOperativoAlta()
  syncLicenciasAlta()
  cargarEquipos()
})

/* =========================
CARGAR EQUIPOS (lista + filtros)
========================= */

async function cargarEquipos() {
  const texto = document.getElementById("buscar").value
  const estado = document.getElementById("filtroEstado").value

  const params = new URLSearchParams()
  params.set("estado", estado)
  if (texto.trim() !== "") {
    params.set("q", texto)
  }

  const res = await fetch(`/buscar?${params.toString()}`, {
    credentials: "include",
    headers: { Accept: "application/json" }
  })

  if (redirigirSi401(res)) {
    return
  }

  if (!res.ok) {
    return
  }

  const equipos = await res.json()
  renderTabla(equipos)
}

function renderTabla(equipos) {
  const tbody = document.querySelector("#tablaEquipos tbody")
  tbody.innerHTML = ""

  equipos.forEach((eq) => {
    tbody.innerHTML += `
<tr>
<td>${eq.codigo || "-"}</td>
<td>${eq.marca}</td>
<td>${eq.modelo}</td>
<td>${eq.tipo}</td>
<td>${etiquetaArea(eq.departamento)}</td>
<td>${eq.usuario}</td>
<td>${etiquetaEstadoMostrar(eq.estado)}</td>
<td>
<button type="button" onclick="verEquipo('${String(eq.codigo || "").replace(/'/g, "\\'")}')">Ver</button>
</td>
<td>
<button type="button" onclick="editarEquipo('${String(eq.codigo || "").replace(/'/g, "\\'")}')">Editar</button>
</td>
</tr>
`
  })
}

/* =========================
BUSCADOR Y FILTRO ESTADO
========================= */

const buscador = document.getElementById("buscar")
const filtroEstado = document.getElementById("filtroEstado")

buscador.addEventListener("keyup", () => {
  cargarEquipos()
})

filtroEstado.addEventListener("change", () => {
  cargarEquipos()
})

/* =========================
FUNCIONES
========================= */

function exportarExcel() {
  let url = "/exportar?ts=" + Date.now()
  fetch(url, { cache: "no-store", credentials: "include" })
    .then(async (res) => {
      if (res.status === 401) {
        window.location.href = "/login"
        throw new Error("Sesión expirada")
      }
      if (!res.ok) {
        let detail = ""
        try {
          detail = (await res.text()).trim()
        } catch {
          detail = ""
        }
        throw new Error(detail || "Error al exportar")
      }
      const cd = res.headers.get("Content-Disposition")
      let filename = "inventario-ti.xlsx"
      if (cd) {
        const utf8 = /filename\*=UTF-8''([^;\n]+)/i.exec(cd)
        const ascii = /filename="([^"]+)"/i.exec(cd)
        if (utf8 && utf8[1]) {
          try {
            filename = decodeURIComponent(utf8[1].trim())
          } catch {
            filename = utf8[1].trim()
          }
        } else if (ascii && ascii[1]) {
          filename = ascii[1].trim()
        }
      }
      return res.blob().then((blob) => ({ blob, filename }))
    })
    .then(({ blob, filename }) => {
      const a = document.createElement("a")
      a.href = URL.createObjectURL(blob)
      a.download = filename
      a.rel = "noopener"
      document.body.appendChild(a)
      a.click()
      a.remove()
      setTimeout(() => URL.revokeObjectURL(a.href), 2000)
    })
    .catch((err) => {
      const fallback =
        "No se pudo exportar. Reinicie el servidor (node server.js) y vuelva a intentar."
      const text = err && err.message ? err.message : fallback
      if (typeof Swal !== "undefined") {
        Swal.fire({
          icon: "error",
          title: "Error al exportar",
          text: text || fallback,
          confirmButtonColor: "#262626"
        })
      } else {
        alert(text || fallback)
      }
    })
}

function verEquipo(codigo) {
  window.location.href = "/equipo/" + encodeURIComponent(codigo)
}

function editarEquipo(codigo) {
  window.location.href = "/equipo/" + encodeURIComponent(codigo) + "/editar"
}

function irDashboard() {
  window.location.href = "/dashboard"
}
