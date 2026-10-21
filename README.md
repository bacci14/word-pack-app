# Word Pack

Página para editar varios documentos de Word a la vez:

1. **Buscar y reemplazar** palabras en todos los documentos (cuerpo, tablas, cuadros de texto, encabezados y pies).
2. **Portada y encabezado, uno por uno**: para cada documento tildás si querés cambiarlos; la app muestra las fechas que encontró (por ejemplo `22/10/2026`) y escribís la nueva, o la elegís del calendario y la app respeta el formato.
3. **Descarga** un `.zip` con los Word editados, listos para pasar a PDF con Microsoft Word.

Los archivos se procesan **dentro del navegador**: no se suben a GitHub ni a ningún servidor.

## Publicarla en GitHub Pages

1. Entrá a <https://github.com/new>, poné un nombre (por ejemplo `word-pack`) y creá el repositorio.
2. En el repositorio: **Add file → Upload files**, arrastrá todo el contenido de esta carpeta (`index.html`, `app.js`, la carpeta `lib` y este `README.md`; el archivo oculto `.nojekyll` es opcional) y tocá **Commit changes**.
3. Andá a **Settings → Pages**. En *Build and deployment* elegí **Deploy from a branch**, rama `main`, carpeta `/ (root)` y **Save**.
4. En uno o dos minutos aparece arriba el link, del estilo `https://TU-USUARIO.github.io/word-pack/`. Guardalo en favoritos.

### Sobre la privacidad

Con una cuenta gratuita, una página de GitHub Pages es pública: cualquiera que tenga el link la puede abrir (aunque el repositorio sea privado, eso requiere un plan pago y la página igual queda pública). En este caso no es un problema, porque la página está vacía: tus documentos nunca salen de tu compu, así que nadie más puede ver tu trabajo. La página tiene además una etiqueta para que Google no la indexe.

## Pasar los Word a PDF

Después de **Procesar y descargar**, abrí el `.zip` descargado y entrá a la carpeta.

- **Mac:** seleccioná los Word (Cmd+A), clic derecho → **Acciones rápidas → Convertir a PDF**. También funciona con clic derecho sobre la carpeta entera. Esta opción se instala una sola vez en `~/Library/Services/Convertir a PDF.workflow`.
- **Windows:** doble clic en `Convertir a PDF (Windows).bat`.

Word exporta cada documento y deja los PDF en la misma carpeta. La primera vez, la Mac puede preguntar si se permite controlar Microsoft Word (respondé **Permitir**) y Word puede pedir acceso a la carpeta (**Seleccionar** / **Conceder acceso**).

## Notas

- El buscar y reemplazar no distingue mayúsculas, salvo que tildes la opción.
- La lista de reemplazos queda guardada en el navegador para la próxima vez.
- Con *Aplicar también al nombre del archivo*, los reemplazos se aplican al nombre de los Word descargados.
- Si un documento no tiene la fecha en la portada o el encabezado, igual podés editar esos textos desde *Otros textos de portada y encabezado*.
