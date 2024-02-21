function Listar_Archivos_OE_nivelEpub() {
    try {
        var hoja = SpreadsheetApp.getActive().getSheetByName("hoja1");

        if (!hoja) {
            throw new Error("La hoja con el nombre 'hoja1' no se encontró.");
        }

        var cabecera = hoja.getRange("A1:G1");
        var Datoscabecera = cabecera.setValues([["Nombre", "ID-archivo", "tipo", "url-archivo", "Folder", "Url-folder", "Tamaño"]]);

        var folder = DriveApp.getFolderById('1lH2Q7cNnoo958R5A3MJVsgS3-vS-4_PD');
        var Mis_Archivos = folder.searchFiles('mimeType = "application/epub+zip" ');

        while (Mis_Archivos.hasNext()) {
            var Archivo = Mis_Archivos.next();
            var parentFolder = Archivo.getParents().next();

            hoja.appendRow([
                Archivo.getName(),
                Archivo.getId(),
                Archivo.getMimeType(),
                Archivo.getUrl(),
                parentFolder.getName(),
                parentFolder.getUrl(),
                (Archivo.getSize() * 0.000977).toFixed() + " KB"
            ]);
        }
    } catch (error) {
        console.error("Error en el script:", error.message);
    }
}