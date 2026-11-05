# Seguimiento OMEGA

Página de seguimiento del Proyecto OMEGA (IPPB). Este repositorio contiene **solo código**: ningún dato del cliente.
Los datos viven en el sitio de SharePoint de IPPB y cada persona entra con su propia cuenta.

## Archivos

| Archivo | Qué es | ¿Se edita? |
| --- | --- | --- |
| `index.html` | La plataforma (pantallas, reglas, gráficos, exportación a Excel) | Solo para cambios de funcionalidad |
| `conexion.js` | Todo lo que habla con Microsoft 365: inicio de sesión, lista `OmegaDatos`, carpetas y subida de archivos | Solo si cambia la forma de conectar |
| `config.json` | Los datos de conexión: `clientId`, `tenantId`, `sitio`, `lista`, `biblioteca`, `carpetaRaiz` | Sí, una vez |

`clientId` y `tenantId` **no son secretos**: sin una cuenta de IPPB no sirven para nada.
`biblioteca` vacío = la biblioteca «Documentos» del sitio. `carpetaRaiz` = carpeta dentro de esa biblioteca donde van los procesos.

## Reglas

- Nunca subir aquí el cronograma, el maestro de procesos, `omega-datos-iniciales.json`, reportes ni capturas con datos.
- Cada cambio queda en el historial de commits; si algo se rompe, se vuelve a la versión anterior desde GitHub.

## Soporte rápido

| Lo que ves | Causa probable | Qué hacer |
| --- | --- | --- |
| «Falta configurar la conexión» | `config.json` incompleto | Completar `clientId`, `tenantId` y `sitio` |
| Error AADSTS50011 al entrar | La dirección de la página no coincide con la Redirect URI | En Entra ID → la app → Authentication, poner la dirección exacta, con `/` final |
| «Se necesita aprobación del administrador» | La organización no deja aceptar el permiso `Sites.ReadWrite.All` | Un administrador de Microsoft 365 pulsa «Grant admin consent» en la app |
| «Todavía no tienes acceso» | El correo de la persona no está en Equipo | Equipo → Dar acceso, con el correo exacto que muestra la pantalla |
| «No tiene permiso en el sitio» (403) | La persona no es miembro del sitio de SharePoint | Agregarla como Miembro del sitio |
| «No se pudo guardar» | Sin internet o SharePoint ocupado | Pulsar Reintentar; los cambios siguen en pantalla |
| Datos raros en una fila | Alguien editó la lista `OmegaDatos` a mano | Lista → la fila → Historial de versiones → restaurar |
