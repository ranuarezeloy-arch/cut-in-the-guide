// ============================================================
// SLICE IMAGE BY HORIZONTAL & VERTICAL GUIDES (GRID CUT)
// ============================================================
#target photoshop

function main() {
    if (app.documents.length === 0) {
        alert("First, open the image you want to crop.");
        return;
    }

    var docOriginal = app.activeDocument;
    var docInfo = validateDocument(docOriginal);

    buildUI(docOriginal, docInfo);
}


function readGuides(doc) {
    var unidadReglaOriginal = app.preferences.rulerUnits;
    app.preferences.rulerUnits = Units.PIXELS;

    var guiasH = [];
    var guiasV = [];

    for (var g = 0; g < doc.guides.length; g++) {
        var guia = doc.guides[g];
        var val = (typeof guia.coordinate === "object") ? guia.coordinate.value : parseFloat(guia.coordinate);

        if (guia.direction === Direction.HORIZONTAL) {
            guiasH.push(val);
        } else if (guia.direction === Direction.VERTICAL) {
            guiasV.push(val);
        }
    }

    app.preferences.rulerUnits = unidadReglaOriginal;

    return { guiasH: guiasH, guiasV: guiasV };
}


function calculateAxisCuts(posiciones, limiteMax) {
    posiciones.sort(function (a, b) { return a - b; });

    var puntos = [0];
    for (var i = 0; i < posiciones.length; i++) {
        var pos = Math.round(posiciones[i]);
        if (pos > 0 && pos < limiteMax) {
            // Evitar duplicados exactos
            if (puntos[puntos.length - 1] !== pos) {
                puntos.push(pos);
            }
        }
    }
    if (puntos[puntos.length - 1] !== Math.round(limiteMax)) {
        puntos.push(Math.round(limiteMax));
    }

    return puntos;
}


function validateDocument(doc) {
    var unidadReglaOriginal = app.preferences.rulerUnits;
    app.preferences.rulerUnits = Units.PIXELS;

    var anchoTotal = doc.width.value;
    var altoTotal = doc.height.value;

    app.preferences.rulerUnits = unidadReglaOriginal;

    var guiasInfo = readGuides(doc);
    
    var cortesY = calculateAxisCuts(guiasInfo.guiasH, altoTotal);
    var cortesX = calculateAxisCuts(guiasInfo.guiasV, anchoTotal);

    var numFilas = cortesY.length - 1;
    var numColumnas = cortesX.length - 1;
    var TotalPieces = numFilas * numColumnas;

    return {
        nombre: doc.name,
        ancho: anchoTotal,
        alto: altoTotal,
        modoColor: getColorModeText(doc.mode),
        resolucion: doc.resolution,
        numGuiasH: guiasInfo.guiasH.length,
        numGuiasV: guiasInfo.guiasV.length,
        cortesY: cortesY,
        cortesX: cortesX,
        numPiezas: TotalPieces,
        valido: (doc.guides.length > 0)
    };
}

function getColorModeText(mode) {
    switch (mode) {
        case DocumentMode.RGB: return "RGB";
        case DocumentMode.CMYK: return "CMYK";
        case DocumentMode.GRAYSCALE: return "Grayscale";
        case DocumentMode.BITMAP: return "Bitmap";
        case DocumentMode.LAB: return "Lab Color";
        case DocumentMode.INDEXEDCOLOR: return "Indexed Color";
        case DocumentMode.MULTICHANNEL: return "Multichannel";
        case DocumentMode.DUOTONE: return "Duotone";
        default: return "Unknown";
    }
}

function generateFilename(nombreOriginal, incluirFechaHora, numeroParte, extension) {
    var numeroTexto = ("00" + numeroParte).slice(-3);

    if (incluirFechaHora) {
        var ahora = new Date();
        var fechaHora =
            ahora.getFullYear() +
            ("0" + (ahora.getMonth() + 1)).slice(-2) +
            ("0" + ahora.getDate()).slice(-2) + "_" +
            ("0" + ahora.getHours()).slice(-2) +
            ("0" + ahora.getMinutes()).slice(-2);

        return nombreOriginal + "_" + fechaHora + "_parte_" + numeroTexto + extension;
    }

    return nombreOriginal + "_parte_" + numeroTexto + extension;
}

function getUniqueFile(carpeta, nombreArchivo) {
    var archivo = new File(carpeta + "/" + nombreArchivo);
    if (!archivo.exists) return archivo;

    var punto = nombreArchivo.lastIndexOf(".");
    var base = (punto > -1) ? nombreArchivo.substring(0, punto) : nombreArchivo;
    var ext = (punto > -1) ? nombreArchivo.substring(punto) : "";

    var contador = 1;
    var nuevoArchivo;
    do {
        nuevoArchivo = new File(carpeta + "/" + base + "_dup" + contador + ext);
        contador++;
    } while (nuevoArchivo.exists);

    return nuevoArchivo;
}

function savePNG(doc, archivo) {
    var opcionesPNG = new PNGSaveOptions();
    opcionesPNG.compression = 6;
    opcionesPNG.interlaced = false;
    doc.saveAs(archivo, opcionesPNG, true, Extension.LOWERCASE);
}

function saveJPG(doc, archivo, calidad) {
    var opcionesJPG = new JPEGSaveOptions();
    opcionesJPG.quality = calidad;
    doc.saveAs(archivo, opcionesJPG, true, Extension.LOWERCASE);
}

function openURL(url) {
    try {
        if ($.os.indexOf("Windows") !== -1) {
            system("start " + url);
        } else {
            system("open " + url);
        }
    } catch (e) {
        var f = new File(Folder.temp + "/temp_donations.url");
        f.open("w");
        f.writeln("[InternetShortcut]");
        f.writeln("URL=" + url);
        f.close();
        f.execute();
    }
}


function exportPieces(docOriginal, docInfo, config) {
    var nombreOriginal = docOriginal.name.replace(/\.[^\.]+$/, "");
    var extension = (config.formatoSalida === "JPG") ? ".jpg" : ".png";

    var unidadReglaOriginal = app.preferences.rulerUnits;
    app.preferences.rulerUnits = Units.PIXELS;

    var cortesY = docInfo.cortesY;
    var cortesX = docInfo.cortesX;

    var numeroParte = 1;
    var exportadas = 0;
    var errores = [];

    
    for (var r = 0; r < cortesY.length - 1; r++) {
        var yInicio = cortesY[r];
        var yFin = cortesY[r + 1];

        if (yFin - yInicio <= 0) continue;

       
        for (var c = 0; c < cortesX.length - 1; c++) {
            var xInicio = cortesX[c];
            var xFin = cortesX[c + 1];

            if (xFin - xInicio <= 0) continue;

            var docTemp = null;
            try {
                docTemp = docOriginal.duplicate();
                app.activeDocument = docTemp;

               
                docTemp.crop([
                    UnitValue(xInicio, "px"),
                    UnitValue(yInicio, "px"),
                    UnitValue(xFin, "px"),
                    UnitValue(yFin, "px")
                ]);

                docTemp.flatten();

                var nombreArchivo = generateFilename(nombreOriginal, config.incluirFechaHora, numeroParte, extension);

                var archivoSalida = config.sobrescribir
                    ? new File(config.carpetaSalida + "/" + nombreArchivo)
                    : getUniqueFile(config.carpetaSalida, nombreArchivo);

                if (config.formatoSalida === "JPG") {
                    saveJPG(docTemp, archivoSalida, config.calidadJPG);
                } else {
                    savePNG(docTemp, archivoSalida);
                }

                docTemp.close(SaveOptions.DONOTSAVECHANGES);
                exportadas++;
            } catch (errorCorte) {
                errores.push("Parte " + numeroParte + " (R:" + r + " C:" + c + "): " + errorCorte.toString());
                try { if (docTemp) docTemp.close(SaveOptions.DONOTSAVECHANGES); } catch (e2) {}
            }

            numeroParte++;
        }
    }

    app.preferences.rulerUnits = unidadReglaOriginal;

    return { exportadas: exportadas, errores: errores };
}

function buildUI(docOriginal, docInfo) {

    var urlDonacion = "https://ko-fi.com/eloyjosueranuarez"; 

    var config = {
        formatoSalida: "PNG",
        calidadJPG: 12,
        carpetaSalida: null,
        incluirFechaHora: true,
        sobrescribir: false,
        mantenerAbierto: true,
        abrirCarpetaAlFinalizar: true
    };

    var win = new Window("dialog", "Slice Image by Guides (2D Grid)");
    win.orientation = "column";
    win.alignChildren = "fill";
    win.spacing = 10;
    win.margins = 16;
    win.preferredSize.width = 460;


    var pnlInfo = win.add("panel", undefined, "Document Information");
    pnlInfo.orientation = "column";
    pnlInfo.alignChildren = "left";
    pnlInfo.margins = [12, 16, 12, 12];
    pnlInfo.spacing = 4;

    pnlInfo.add("statictext", undefined, "File: " + docInfo.nombre);
    pnlInfo.add("statictext", undefined, "Dimensions: " + docInfo.ancho + " x " + docInfo.alto + " px");
    pnlInfo.add("statictext", undefined, "Horizontal guides: " + docInfo.numGuiasH + " | Vertical guides: " + docInfo.numGuiasV);
    pnlInfo.add("statictext", undefined, "Total pieces to export: " + docInfo.numPiezas);

    if (!docInfo.valido) {
        var lblAdvertencia = pnlInfo.add("statictext", undefined,
            "No guides found. Add at least one guide before exporting.", { multiline: true });
        lblAdvertencia.graphics.foregroundColor = lblAdvertencia.graphics.newPen(
            lblAdvertencia.graphics.PenType.SOLID_COLOR, [0.8, 0.2, 0.2], 1);
    }


    var pnlFolder = win.add("panel", undefined, "Output Folder");
    pnlFolder.orientation = "row";
    pnlFolder.alignChildren = "center";
    pnlFolder.margins = [12, 16, 12, 12];
    pnlFolder.spacing = 8;

    var txtFolder = pnlFolder.add("edittext", undefined, "");
    txtFolder.enabled = false;
    txtFolder.preferredSize.width = 300;

    var btnBrowse = pnlFolder.add("button", undefined, "Browse...");
    btnBrowse.onClick = function () {
        var carpeta = Folder.selectDialog("Select output folder");
        if (carpeta !== null) {
            config.carpetaSalida = carpeta;
            txtFolder.text = carpeta.fsName;
            btnStart.enabled = docInfo.valido;
        }
    };

   
    var pnlFormat = win.add("panel", undefined, "Output Format");
    pnlFormat.orientation = "row";
    pnlFormat.alignChildren = "center";
    pnlFormat.margins = [12, 16, 12, 12];
    pnlFormat.spacing = 16;

    var rbPNG = pnlFormat.add("radiobutton", undefined, "PNG");
    var rbJPG = pnlFormat.add("radiobutton", undefined, "JPG");
    rbPNG.value = true;

  
    var pnlQuality = win.add("panel", undefined, "JPEG Quality");
    pnlQuality.orientation = "row";
    pnlQuality.alignChildren = "center";
    pnlQuality.margins = [12, 16, 12, 12];
    pnlQuality.spacing = 10;

    var sliderQuality = pnlQuality.add("slider", undefined, 12, 1, 12);
    sliderQuality.preferredSize.width = 260;

    var lblQualityValue = pnlQuality.add("statictext", undefined, "12");
    lblQualityValue.preferredSize.width = 24;

    sliderQuality.onChanging = function () {
        sliderQuality.value = Math.round(sliderQuality.value);
        lblQualityValue.text = String(sliderQuality.value);
        config.calidadJPG = sliderQuality.value;
        actualizarVistaPrevia();
    };

    function actualizarVisibilidadCalidad() {
        if (rbJPG.value) {
            config.formatoSalida = "JPG";
            pnlQuality.visible = true;
            sliderQuality.enabled = true;
        } else {
            config.formatoSalida = "PNG";
            pnlQuality.visible = false;
            sliderQuality.enabled = false;
        }
        win.layout.layout(true);
        actualizarVistaPrevia();
    }

    rbPNG.onClick = actualizarVisibilidadCalidad;
    rbJPG.onClick = actualizarVisibilidadCalidad;


    var pnlPreview = win.add("panel", undefined, "Filename Preview");
    pnlPreview.orientation = "column";
    pnlPreview.alignChildren = "left";
    pnlPreview.margins = [12, 16, 12, 12];

    var lblPreview = pnlPreview.add("statictext", undefined, "");
    lblPreview.preferredSize.width = 420;

    function actualizarVistaPrevia() {
        var nombreOriginal = docOriginal.name.replace(/\.[^\.]+$/, "");
        var extension = (config.formatoSalida === "JPG") ? ".jpg" : ".png";
        lblPreview.text = generateFilename(nombreOriginal, chkFechaHora.value, 1, extension);
    }


    var pnlOptions = win.add("panel", undefined, "Additional Options");
    pnlOptions.orientation = "column";
    pnlOptions.alignChildren = "left";
    pnlOptions.margins = [12, 16, 12, 12];
    pnlOptions.spacing = 4;

    var chkAbrirCarpeta = pnlOptions.add("checkbox", undefined, "Open the output folder when finished");
    chkAbrirCarpeta.value = true;

    var chkFechaHora = pnlOptions.add("checkbox", undefined, "Include date and time in the filename");
    chkFechaHora.value = true;
    chkFechaHora.onClick = function () {
        config.incluirFechaHora = chkFechaHora.value;
        actualizarVistaPrevia();
    };

    var chkSobrescribir = pnlOptions.add("checkbox", undefined, "Overwrite existing files");
    chkSobrescribir.value = false;

    var chkMantenerAbierto = pnlOptions.add("checkbox", undefined, "Keep the original document open");
    chkMantenerAbierto.value = true;


    var configButtons = win.add("group");
    configButtons.orientation = "row";
    configButtons.alignment = "fill";
    configButtons.spacing = 8;

    var btnDonate = configButtons.add("button", undefined, "Donate");
    btnDonate.alignment = ["left", "center"];
    btnDonate.onClick = function () {
        openURL(urlDonacion);
    };

    var rightButtons = configButtons.add("group");
    rightButtons.alignment = ["right", "center"];
    rightButtons.spacing = 8;

    var btnCancel = rightButtons.add("button", undefined, "Cancel", { name: "cancel" });
    var btnStart = rightButtons.add("button", undefined, "Start Export", { name: "ok" });
    btnStart.enabled = false;

    actualizarVisibilidadCalidad();

    btnCancel.onClick = function () { win.close(); };

    btnStart.onClick = function () {
        if (!docInfo.valido || config.carpetaSalida === null) return;

        config.calidadJPG = sliderQuality.value;
        config.incluirFechaHora = chkFechaHora.value;
        config.sobrescribir = chkSobrescribir.value;
        config.mantenerAbierto = chkMantenerAbierto.value;
        config.abrirCarpetaAlFinalizar = chkAbrirCarpeta.value;

        win.close();

        var resultado = exportPieces(docOriginal, docInfo, config);

        if (!config.mantenerAbierto) {
            try {
                app.activeDocument = docOriginal;
                docOriginal.close(SaveOptions.DONOTSAVECHANGES);
            } catch (eClose) {}
        }

        if (config.abrirCarpetaAlFinalizar) {
            try { config.carpetaSalida.execute(); } catch (eOpen) {}
        }

        var msj = "Process completed successfully.\n\n" +
                  "Exported Parts: " + resultado.exportadas + " / " + docInfo.numPiezas + "\n" +
                  "Folder: " + config.carpetaSalida.fsName;

        if (resultado.errores.length > 0) {
            msj += "\n\nERRORS FOUND:\n" + resultado.errores.join("\n");
        }

        alert(msj);
    };

    win.center();
    win.show();
}

main();
