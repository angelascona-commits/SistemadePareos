
document.getElementById('btnProcesar').addEventListener('click', procesarArchivos);

function limpiarTexto(texto) {
    if (texto === undefined || texto === null) return "";
    let t = String(texto).toUpperCase();
    t = t.replace(/[\n\r]/g, ' '); 
    if (t.endsWith('.0')) t = t.slice(0, -2); 
    t = t.normalize("NFD").replace(/[\u0300-\u036f]/g, ""); 
    t = t.replace(/\s+/g, ' ').trim(); 
    return t;
}

function estandarizarFila(fila) {
    let filaLimpia = {};
    for (let key in fila) {
        let nuevaKey = key.trim().toUpperCase();
        filaLimpia[nuevaKey] = fila[key];
    }
    return filaLimpia;
}

function extraerDatosDinamicos(worksheet) {
    const matrizDatos = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
    let filaCabecera = 0;
    
    for (let i = 0; i < matrizDatos.length; i++) {
        let textoFila = matrizDatos[i].join(" ").toUpperCase();
        if (textoFila.includes("NOMBRE COMERCIAL") || textoFila.includes("SNOMBRE_COMERCIAL")) {
            filaCabecera = i;
            break;
        }
    }
    
    return XLSX.utils.sheet_to_json(worksheet, { range: filaCabecera, defval: "" });
}

async function obtenerDatosGoogleSheets(url) {
    const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (!match) throw new Error("URL de Google Sheets no válida.");
    
    const id = match[1];
    const urlCsv = `https://docs.google.com/spreadsheets/d/${id}/export?format=csv`;
    
    const respuesta = await fetch(urlCsv);
    if (!respuesta.ok) throw new Error("No se pudo leer el Google Sheet. Verifica los permisos.");
    
    const csvTexto = await respuesta.text();
    const workbook = XLSX.read(csvTexto, {type: 'string'});
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    
    return extraerDatosDinamicos(worksheet);
}

function leerExcelLocal(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {type: 'array'});
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            const json = extraerDatosDinamicos(worksheet); 
            resolve(json);
        };
        reader.onerror = (error) => reject(error);
        reader.readAsArrayBuffer(file);
    });
}

async function procesarArchivos() {
    const urlSheet = document.getElementById('urlGoogleSheet').value.trim();
    const fileNuevo = document.getElementById('archivoNuevo').files[0];
    const btn = document.getElementById('btnProcesar');
    const msj = document.getElementById('mensaje');

    if (!urlSheet || !fileNuevo) {
        msj.innerText = "⚠️ Faltan datos (URL o Archivo).";
        return;
    }

    btn.disabled = true;
    btn.innerText = "Comparando Nombres Comerciales...";
    msj.innerText = "";

    try {
        let dataBaseBruta = await obtenerDatosGoogleSheets(urlSheet);
        let dataNuevaBruta = await leerExcelLocal(fileNuevo);

        let baseMap = new Map();
        let nuevaMap = new Map();

        dataBaseBruta.forEach(filaBruta => {
            let fila = estandarizarFila(filaBruta);
            let nombreBD = limpiarTexto(fila['SNOMBRE_COMERCIAL']);
            let distritoBD = limpiarTexto(fila['SDISTRITO']); 

            if (nombreBD) {
                let huella = `${nombreBD}|${distritoBD}`;
                baseMap.set(huella, fila);
            }
        });

        dataNuevaBruta.forEach(filaBruta => {
            let fila = estandarizarFila(filaBruta);
            let nombreNuevo = limpiarTexto(fila['NOMBRE COMERCIAL']);
            let distritoNuevo = limpiarTexto(fila['DISTRITO']);

            if (nombreNuevo && nombreNuevo !== "LIMA Y CALLAO" && nombreNuevo !== "PROVINCIAS") {
                let huella = `${nombreNuevo}|${distritoNuevo}`;
                nuevaMap.set(huella, filaBruta);
            }
        });

        let mantiene = [];
        let agregados = [];
        let eliminados = [];

        nuevaMap.forEach((filaOriginal, huella) => {
            let filaFinal = { ...filaOriginal };
            if (baseMap.has(huella)) {
                filaFinal['ESTADO'] = 'MANTIENE';
                mantiene.push(filaFinal);
            } else {
                filaFinal['ESTADO'] = 'AGREGADO';
                agregados.push(filaFinal);
            }
        });

        baseMap.forEach((filaBD, huella) => {
            if (!nuevaMap.has(huella)) { 
                let filaEliminada = {
                    'NOMBRE COMERCIAL': filaBD['SNOMBRE_COMERCIAL'],
                    'DISTRITO': filaBD['SDISTRITO'],
                    'DIRECCION': filaBD['SDIRECCION'],
                    'TELEFONO': filaBD['STELEFONO'],
                    'ESTADO': 'ELIMINADO'
                };
                eliminados.push(filaEliminada);
            }
        });

        let resultadoFinal = [...mantiene, ...agregados, ...eliminados];

        if (resultadoFinal.length === 0) throw new Error("No se encontraron datos.");

        const ws = XLSX.utils.json_to_sheet(resultadoFinal);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Resultados");
        XLSX.writeFile(wb, "Reporte_Actualizado.xlsx");

        msj.style.color = "#27ae60";
        msj.innerText = ` Mantiene: ${mantiene.length} | Agregados: ${agregados.length} | Eliminados: ${eliminados.length}`;

    } catch (error) {
        console.error(error);
        msj.style.color = "#c0392b";
        msj.innerText = " Error: " + error.message;
    } finally {
        btn.disabled = false;
        btn.innerText = "Comparar y Descargar Resultados";
    }
}