# Inventario TI

Sistema web de gestión de inventario tecnológico desarrollado con Node.js y Express. Permite registrar, buscar, editar y dar seguimiento a equipos de TI dentro de una organización.

## Demo en vivo

🔗 [Ver demo](https://jmm-inventario.vercel.app/login)

> Los datos del demo son ficticios y se reinician automáticamente. Los cambios no persisten entre sesiones.

## Funcionalidades

- **Login** con autenticación por sesión
- **Dashboard** con resumen de equipos por estado (activos, inactivos, dados de baja)
- **Inventario** con tabla de equipos, búsqueda en tiempo real y filtros por estado
- **Registro de equipos** con validación de campos (tipo, marca, área, sistema operativo, licencias)
- **Edición de equipos** con cambio de usuario, área y estado
- **Historial de cambios** por equipo (transferencias, cambios de estado, cambios de área)
- **Notas por equipo** para registrar observaciones o intervenciones técnicas
- **Exportación a Excel** con formato profesional y colores por estado
- **Etiqueta imprimible** con código del equipo para identificación física
- **Gestión de usuarios** (disponible para SuperAdmin)

## Tecnologías

| Capa | Tecnología |
|------|-----------|
| Backend | Node.js + Express |
| Sesiones | express-session |
| Exportación | ExcelJS |
| Frontend | HTML + CSS + JavaScript vanilla |
| Alertas UI | SweetAlert2 |
| Deploy | Vercel |

## Estructura del proyecto

```
inventario-ti/
├── server.js          # Servidor Express + API REST + mock data
├── package.json
├── vercel.json        # Configuración de deploy
└── public/
    ├── index.html     # Vista principal del inventario
    ├── login.html     # Pantalla de inicio de sesión
    ├── app.js         # Lógica del frontend
    └── logosinfondo.png
```


El servidor quedará disponible en `http://localhost:3001`.

## Áreas gestionadas

El sistema contempla las siguientes áreas de la organización:

- TI
- Bodega
- Frigorífico
- RRHH
- Gerencia
- Administración
- Packing
- Mantención

## Tipos de equipos soportados

Notebooks, Escritorio, Impresora, Switch, Router, Antena, Repetidores.

## Autor

**Joaquín Marín Muñoz** — Ingeniero en Informática · Fullstack Developer

[![LinkedIn](https://img.shields.io/badge/LinkedIn-joaquin--marin--munoz-blue?logo=linkedin)](https://www.linkedin.com/in/joaquin-marin-munoz/)
[![GitHub](https://img.shields.io/badge/GitHub-Jmarinmunoz-black?logo=github)](https://github.com/Jmarinmunoz)
