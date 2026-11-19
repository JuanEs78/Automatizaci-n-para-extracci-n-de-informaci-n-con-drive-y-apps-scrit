// ─────────────────────────────────────────────
// CONFIGURACIÓN GLOBAL
// ─────────────────────────────────────────────

// ID de la carpeta de Google Drive donde están los contratos
// Se obtiene de la URL de la carpeta: drive.google.com/drive/folders/ESTE_ES_EL_ID
const FOLDER_ID = "";

// Nombre que tendrá la hoja dentro del Google Sheets resultado
const SHEET_NAME = "Datos";


// ─────────────────────────────────────────────
// FUNCIÓN 1: Crear el archivo Google Sheets de resultados
// ─────────────────────────────────────────────
function crearSpreadsheet() {
  // Captura la fecha y hora actual del sistema
  const ahora = new Date();

  // Formatea la fecha como "28-04-2026 08:18:04" para usarla en el nombre del archivo
  // Session.getScriptTimeZone() usa la zona horaria configurada en el proyecto (Colombia = America/Bogota)
  const nombre = `Extracción ${Utilities.formatDate(ahora, Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm:ss")}`;

  // Crea el Sheets con ese nombre
  // IMPORTANTE: SpreadsheetApp.create() se usa porque el script es independiente
  // (no está vinculado a ningún Sheets existente).
  // Antes se usaba getActiveSpreadsheet() pero devolvía null porque no hay hoja activa.
  const ss = SpreadsheetApp.create(nombre);

  // Obtiene el ID del Sheets recién creado para poder moverlo
  const fileId = ss.getId();

  // Obtiene el token de autorización del usuario que ejecuta el script
  // Es como el "carnet de acceso" que le permite al script hablar con la API de Google
  const token = ScriptApp.getOAuthToken();

  // Mueve el Sheets a la carpeta correcta usando la Drive API v3
  // Por defecto Google crea el archivo en la raíz de "Mi Unidad", no en la carpeta deseada
  // addParents: agrega la carpeta destino como ubicación del archivo
  // removeParents=root: lo quita de la raíz
  UrlFetchApp.fetch(
    `https://www.googleapis.com/drive/v3/files/${fileId}?addParents=${FOLDER_ID}&removeParents=root`,
    {
      method: "PATCH", // PATCH = actualizar solo algunos datos del archivo (no reemplazarlo completo)
      headers: { Authorization: "Bearer " + token }, // Envía el token para autenticarse
      muteHttpExceptions: true // Si hay error HTTP, no lanza excepción, lo maneja el código
    }
  );

  // Devuelve el objeto Sheets para que otras funciones puedan escribir en él
  return ss;
}


// ─────────────────────────────────────────────
// FUNCIÓN 2: Convertir un archivo .docx a Google Doc temporalmente
// ─────────────────────────────────────────────
function convertirDocxAGoogleDoc(file) {
  // Obtiene el token de autorización
  const token = ScriptApp.getOAuthToken();

  // getBlob() obtiene el archivo como objeto binario (los bytes del archivo Word)
  const blob = file.getBlob();
  const bytes = blob.getBytes();

  // Validación: si el archivo está vacío (0 bytes), no intenta convertirlo
  // Archivos vacíos causaban errores "Empty response" en la API
  if (bytes.length === 0) {
    Logger.log(`⚠️ Archivo vacío: ${file.getName()}`);
    return null; // Retorna null para que la función principal lo salte
  }

  // Metadatos que le dicen a la API de Drive cómo crear el archivo convertido:
  // - name: mismo nombre que el .docx original
  // - mimeType: "application/vnd.google-apps.document" es el tipo MIME de Google Docs
  //   (un tipo MIME identifica el formato de un archivo, como "application/pdf" para PDFs)
  const metadata = JSON.stringify({
    name: file.getName(),
    mimeType: "application/vnd.google-apps.document"
  });

  // El "boundary" es un separador inventado para dividir el cuerpo de la petición en partes
  // Se usa en peticiones "multipart" (que envían varios tipos de datos juntos)
  const boundary = "-------314159265358979323846";
  const delimiter = "\r\n--" + boundary + "\r\n";   // Separador entre partes
  const closeDelim = "\r\n--" + boundary + "--";    // Cierre final del multipart

  // Construye el cuerpo de la petición con dos partes:
  // Parte 1: los metadatos en JSON (nombre y tipo del archivo destino)
  // Parte 2: el contenido del .docx codificado en base64
  // Base64 es una forma de convertir bytes binarios a texto para poder enviarlos por HTTP
  const body = delimiter +
    "Content-Type: application/json\r\n\r\n" + metadata +
    delimiter +
    "Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document\r\n" +
    "Content-Transfer-Encoding: base64\r\n\r\n" +
    Utilities.base64Encode(bytes) + // Convierte los bytes del Word a texto base64
    closeDelim;

  // Sistema de reintentos: intenta hasta 3 veces si falla por límite de quota
  const maxIntentos = 3;
  const esperaMs = [5000, 15000, 30000]; // Tiempos de espera: 5s, 15s, 30s

  for (let intento = 0; intento < maxIntentos; intento++) {

    // Envía la petición a la API de Drive para subir y convertir el archivo
    const response = UrlFetchApp.fetch(
      "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart",
      {
        method: "POST", // POST = crear un recurso nuevo en el servidor
        headers: {
          Authorization: "Bearer " + token,
          "Content-Type": "multipart/related; boundary=" + boundary
        },
        payload: body,          // El cuerpo construido arriba con los metadatos y el archivo
        muteHttpExceptions: true
      }
    );

    // Convierte la respuesta JSON a objeto JavaScript
    const result = JSON.parse(response.getContentText());

    // Si la respuesta incluye un ID, la conversión fue exitosa → retorna el ID
    if (result.id) {
      return result.id;
    }

    // Verifica si el error es por quota de bandwidth (límite de datos transferidos)
    // Esto ocurre cuando se suben varios archivos seguidos y Google frena las transferencias
    const esQuota = response.getContentText().includes("Bandwidth quota") ||
                    response.getContentText().includes("quota");

    if (esQuota && intento < maxIntentos - 1) {
      // Si es error de quota y aún quedan intentos, espera y reintenta
      Logger.log(`⏳ Quota alcanzada en ${file.getName()}, esperando ${esperaMs[intento] / 1000}s... (intento ${intento + 1}/${maxIntentos})`);
      Utilities.sleep(esperaMs[intento]); // Pausa la ejecución los milisegundos indicados
    } else {
      // Si no es quota o ya no quedan intentos, registra el error y abandona
      Logger.log(`❌ Error en conversión de ${file.getName()}: ${JSON.stringify(result)}`);
      return null;
    }
  }

  return null; // Si los 3 intentos fallaron, retorna null
}


// ─────────────────────────────────────────────
// FUNCIÓN 3: Extraer un valor cuando etiqueta y valor están en la misma celda
// Ejemplo: celda contiene "CONTRATISTA: Juan García"
// ─────────────────────────────────────────────
function extraerValorInline(texto, etiqueta) {
  // Busca la posición donde aparece la etiqueta dentro del texto
  // toUpperCase() en ambos para que la búsqueda no distinga mayúsculas/minúsculas
  const idx = texto.toUpperCase().indexOf(etiqueta.toUpperCase());

  // Si no encontró la etiqueta, retorna null
  if (idx === -1) return null;

  // Toma todo el texto que viene DESPUÉS de la etiqueta
  const despues = texto.substring(idx + etiqueta.length).trim();

  // Si empieza con ":", lo quita y retorna el resto
  // Ejemplo: ": Juan García" → "Juan García"
  return despues.startsWith(":") ? despues.substring(1).trim() : despues;
}


// ─────────────────────────────────────────────
// FUNCIÓN 4: Recorre una tabla y extrae los 6 campos buscados
// ─────────────────────────────────────────────
function extraerCamposDeTabla(tabla) {
  // Variables donde se guardarán los valores encontrados
  // Empiezan vacías; solo se llenan si se encuentra la etiqueta correspondiente
  let contratista = "";
  let documentoIdentidad = "";
  let telefonoEmail = "";
  let objeto = "";
  let valor = "";
  let duracion = "";

  // Recorre cada fila de la tabla
  for (let r = 0; r < tabla.getNumRows(); r++) {
    const row = tabla.getRow(r);
    const numCeldas = row.getNumCells();

    // Función interna: recibe etiqueta y valor de celdas separadas (formato 2 columnas)
    // Se declara aquí dentro para poder acceder a las variables de arriba directamente
    const procesarEtiquetaValor = (etiqueta, val) => {
      const e = etiqueta.trim().toUpperCase(); // Normaliza la etiqueta
      const v = val.trim();                    // Limpia espacios del valor

      // Compara la etiqueta y asigna el valor a la variable correcta
      // El "if (!contratista)" evita sobreescribir un valor ya encontrado
      if (e.includes("CONTRATISTA")) { if (!contratista) contratista = v; }
      else if (e.includes("DOCUMENTO DE IDENTIDAD")) { if (!documentoIdentidad) documentoIdentidad = v; }
      else if (e.includes("TELÉFONO") || e.includes("TELEFONO")) { if (!telefonoEmail) telefonoEmail = v; }
      else if (e.includes("OBJETO")) { if (!objeto) objeto = v; }
      else if (e.includes("VALOR")) { if (!valor) valor = v; }
      else if (e.includes("DURACIÓN") || e.includes("DURACION")) { if (!duracion) duracion = v; }
    };

    // CASO A: Tabla con 2 o más columnas → celda 0 = etiqueta, celda 1 = valor
    // Ejemplo: | CONTRATISTA: | Juan García |
    if (numCeldas >= 2) {
      procesarEtiquetaValor(row.getCell(0).getText(), row.getCell(1).getText());
    }

    // CASO B: Revisa cada celda por si contiene etiqueta y valor juntos en una sola celda
    // Ejemplo: | CONTRATISTA: Juan García |
    // Esto cubre tablas de 1 columna Y también casos raros en tablas de 2 columnas
    for (let c = 0; c < numCeldas; c++) {
      const textoCelda = row.getCell(c).getText();
      const upper = textoCelda.toUpperCase();

      // Para cada campo, verifica si la celda contiene la etiqueta y aún no se encontró el valor
      if (!contratista && upper.includes("CONTRATISTA:"))
        contratista = extraerValorInline(textoCelda, "CONTRATISTA:") || "";

      if (!documentoIdentidad && upper.includes("DOCUMENTO DE IDENTIDAD:"))
        documentoIdentidad = extraerValorInline(textoCelda, "DOCUMENTO DE IDENTIDAD:") || "";

      // Para teléfono acepta varias variantes del nombre del campo
      if (!telefonoEmail && (upper.includes("TELÉFONO Y EMAIL:") || upper.includes("TELEFONO Y EMAIL:") ||
                              upper.includes("TELÉFONO:") || upper.includes("TELEFONO:"))) {
        // Detecta cuál variante encontró y usa esa como etiqueta de búsqueda
        telefonoEmail = extraerValorInline(textoCelda,
          upper.includes("TELÉFONO Y EMAIL:") ? "TELÉFONO Y EMAIL:" :
          upper.includes("TELEFONO Y EMAIL:") ? "TELEFONO Y EMAIL:" :
          upper.includes("TELÉFONO:") ? "TELÉFONO:" : "TELEFONO:") || "";
      }

      if (!objeto && upper.includes("OBJETO:"))
        objeto = extraerValorInline(textoCelda, "OBJETO:") || "";

      if (!valor && upper.includes("VALOR:"))
        valor = extraerValorInline(textoCelda, "VALOR:") || "";

      // Para duración acepta con y sin tilde
      if (!duracion && (upper.includes("DURACIÓN:") || upper.includes("DURACION:"))) {
        duracion = extraerValorInline(textoCelda,
          upper.includes("DURACIÓN:") ? "DURACIÓN:" : "DURACION:") || "";
      }
    }
  }

  // Retorna los 6 campos encontrados como un objeto
  return { contratista, documentoIdentidad, telefonoEmail, objeto, valor, duracion };
}


// ─────────────────────────────────────────────
// FUNCIÓN 5: Función principal — coordina todo el proceso
// ─────────────────────────────────────────────
function extraerDatosDeDocumentos() {
  // Abre la carpeta de Drive por su ID
  const folder = DriveApp.getFolderById(FOLDER_ID);

  // Obtiene un iterador con todos los archivos de la carpeta
  // Un iterador no carga todos los archivos a la vez, los entrega uno por uno (más eficiente)
  const files = folder.getFiles();

  // Crea el Sheets de resultados con fecha y hora en el nombre
  const ss = crearSpreadsheet();
  const sheet = ss.getSheets()[0]; // Toma la primera hoja del Sheets
  sheet.setName(SHEET_NAME);       // La renombra como "Datos"

  // Escribe la fila de encabezados en negrita
  sheet.getRange(1, 1, 1, 7).setValues([
    ["Nombre del Documento", "Contratista", "Documento de Identidad", "Teléfono y Email", "Objeto", "Valor", "Duración"]
  ]);
  sheet.getRange(1, 1, 1, 7).setFontWeight("bold");

  // Contador de fila: empieza en 2 porque la fila 1 son los encabezados
  let fila = 2;

  // Recorre cada archivo de la carpeta
  while (files.hasNext()) {
    const file = files.next();
    const mimeType = file.getMimeType(); // Tipo de archivo (Word, Google Doc, imagen, etc.)

    // Verifica si es un archivo .docx
    const esDocx = mimeType === "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
    // Verifica si es un Google Doc nativo
    const esGoogleDoc = mimeType === MimeType.GOOGLE_DOCS;

    // Si no es ninguno de los dos tipos, lo ignora y pasa al siguiente
    if (!esDocx && !esGoogleDoc) continue;

    let docId = null;      // ID del documento que se va a leer
    let tempCreado = false; // Bandera: indica si se creó un temporal que hay que borrar

    try {
      if (esGoogleDoc) {
        // Google Doc nativo: se usa directamente sin conversión
        docId = file.getId();
      } else {
        // .docx: hay que convertirlo primero a Google Doc temporal
        docId = convertirDocxAGoogleDoc(file);
        if (!docId) continue; // Si la conversión falló, salta este archivo
        tempCreado = true;    // Marca que hay un temporal para borrar después
        Utilities.sleep(3000); // Espera 3 segundos para que Drive termine de procesar la conversión
      }
    } catch (e) {
      Logger.log(`❌ Error convirtiendo ${file.getName()}: ${e.message}`);
      continue; // Si hay error inesperado, salta al siguiente archivo
    }

    // Variables para guardar los datos extraídos de este documento
    let contratista = "";
    let documentoIdentidad = "";
    let telefonoEmail = "";
    let objeto = "";
    let valor = "";
    let duracion = "";

    try {
      // Abre el documento (ya sea el original Google Doc o el temporal convertido)
      const doc = DocumentApp.openById(docId);
      const body = doc.getBody(); // Obtiene el cuerpo del documento (todo el contenido)

      // Recorre cada elemento del documento (párrafos, tablas, imágenes, etc.)
      for (let i = 0; i < body.getNumChildren(); i++) {
        const element = body.getChild(i);

        // Solo procesa tablas, ignora párrafos e imágenes
        if (element.getType() !== DocumentApp.ElementType.TABLE) continue;

        // Llama a extraerCamposDeTabla con esta tabla
        const resultado = extraerCamposDeTabla(element.asTable());

        // Solo sobreescribe si encontró algo (no borra datos ya encontrados en tablas anteriores)
        if (resultado.contratista) contratista = resultado.contratista;
        if (resultado.documentoIdentidad) documentoIdentidad = resultado.documentoIdentidad;
        if (resultado.telefonoEmail) telefonoEmail = resultado.telefonoEmail;
        if (resultado.objeto) objeto = resultado.objeto;
        if (resultado.valor) valor = resultado.valor;
        if (resultado.duracion) duracion = resultado.duracion;

        // Optimización: si ya encontró los 6 campos, no sigue revisando más tablas
        if (contratista && documentoIdentidad && telefonoEmail && objeto && valor && duracion) break;
      }
    } catch (e) {
      Logger.log(`❌ Error leyendo doc ${file.getName()}: ${e.message}`);
      // No usa "continue" aquí para que igual intente borrar el temporal si existe
    }

    // Borra el Google Doc temporal si fue creado
    if (tempCreado && docId) {
      try {
        DriveApp.getFileById(docId).setTrashed(true); // Manda a la papelera
      } catch (e) {
        Logger.log(`⚠️ No se pudo borrar temporal: ${e.message}`);
      }
    }

    // Solo escribe la fila si encontró al menos un campo con dato
    // Esto evita filas vacías de documentos sin la tabla requerida
    if (contratista || documentoIdentidad || telefonoEmail || objeto || valor || duracion) {
      sheet.getRange(fila, 1, 1, 7).setValues([
        [file.getName(), contratista, documentoIdentidad, telefonoEmail, objeto, valor, duracion]
      ]);
      fila++; // Avanza a la siguiente fila para el próximo documento
      Logger.log(`✅ ${file.getName()} | Contratista: ${contratista} | Doc: ${documentoIdentidad} | Tel/Email: ${telefonoEmail} | Objeto: ${objeto} | Valor: ${valor} | Duración: ${duracion}`);
    } else {
      Logger.log(`⏭️ Saltado (sin datos): ${file.getName()}`);
    }
  }

  // Mensajes finales en el log
  Logger.log(`\n🏁 Proceso completado. ${fila - 2} documentos con datos procesados.`);
  Logger.log(`📄 Sheets creado: ${ss.getName()} | URL: ${ss.getUrl()}`);
}
