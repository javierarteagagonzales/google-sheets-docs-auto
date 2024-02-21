function Listar_Archivos_Epub2() {
    try {
        var hoja = SpreadsheetApp.getActive().getSheetByName("hoja2");

        if (!hoja) {
            throw new Error("La hoja con el nombre 'hoja2' no se encontró.");
        }

        var cabecera = hoja.getRange("A1:B1");
        var Datoscabecera = cabecera.setValues([["Nombre", "url-archivo"]]);

        var folder = DriveApp.getFolderById('1lH2Q7cNnoo958R5A3MJVsgS3-vS-4_PD');
        var Mis_Archivos = folder.searchFiles('mimeType = "application/epub+zip" ');

        while (Mis_Archivos.hasNext()) {
            var Archivo = Mis_Archivos.next();
            var parentFolder = Archivo.getParents().next();

            hoja.appendRow([
                Archivo.getName(),
                parentFolder.getUrl(),
            ]);
        }
    } catch (error) {
        console.error("Error en el script:", error.message);
    }
}