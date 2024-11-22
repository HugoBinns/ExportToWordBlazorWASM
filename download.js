// Script para descargar el archivo generado
function saveAsFile(fileName, byteBase64) {
    var link = document.createElement('a');
    // Especificar el tipo MIME adecuado para Word
    link.download = fileName;
    link.href = 'data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,' + byteBase64;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}