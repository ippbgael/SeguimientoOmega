# Seguimiento OMEGA

Página de seguimiento del Proyecto OMEGA (IPPB). Este repositorio contiene **solo código**: ningún dato del cliente.

- **Pantallas y lógica:** este repositorio, publicado con GitHub Pages.
- **Datos y usuarios:** un proyecto de Supabase (base de datos + inicio de sesión). Las reglas de quién ve y edita qué están en la base (`supabase.sql`, que no se sube aquí).
- **Documentos:** siguen en SharePoint; la plataforma guarda solo el enlace.

## Archivos

| Archivo | Qué es | ¿Se edita? |
| --- | --- | --- |
| `index.html` | La plataforma: pantallas, reglas, gráficos, reporte Excel | Solo para cambios de funcionalidad |
| `conexion.js` | Todo lo que habla con Supabase: inicio de sesión, leer y guardar | Solo si cambia la forma de conectar |
| `config.json` | `supabaseUrl` y `supabaseKey` (la clave **publishable / anon**) | Sí, una vez |

La clave de `config.json` es pública por diseño: sin un usuario y contraseña válidos no permite ver nada, porque la base exige iniciar sesión y estar en la lista de Equipo. **Nunca** pongas aquí la clave `service_role` / `secret`.

## Soporte rápido

| Lo que ves | Causa probable | Qué hacer |
| --- | --- | --- |
| «Falta configurar la conexión» | `config.json` vacío | Completar `supabaseUrl` y `supabaseKey` |
| «Correo o contraseña incorrectos» | Usuario no creado o clave mal escrita | Supabase → Authentication → Users: revisar el usuario o enviar recuperación |
| «Todavía no tienes acceso» | Su correo no está en Equipo | Equipo → Dar acceso, con el correo exacto |
| «No tienes permiso para este cambio» | Su rol no lo permite (ej. consultor en otro equipo) | Revisar su rol y equipo en Equipo |
| «No hay conexión con el servidor» | Sin internet o proyecto pausado por 7 días sin uso | Supabase → el proyecto → Restore project |
| «No se pudo guardar» | Corte momentáneo | Pulsar Reintentar |
