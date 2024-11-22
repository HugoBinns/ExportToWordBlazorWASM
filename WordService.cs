using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using ExportToWord.Services.IServices;
using System.IO;

namespace ExportToWord.Services
{
    public class WordService : IWordService
    {
        public byte[] GenerateWordDocument(byte[] templateBytes, Dictionary<string, string> parameters)
        {
            using (MemoryStream stream = new MemoryStream(templateBytes))
            {
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(stream, true))
                {
                    // Modificar el contenido principal del documento
                    var body = wordDoc.MainDocumentPart.Document.Body;

                    foreach (var param in parameters)
                    {
                        // Verificar si la llave existe antes de intentar reemplazarla
                        if (!body.Descendants<Text>().Any(t => t.Text.Contains(param.Key)))
                            continue; // Si no existe, se salta al siguiente parámetro

                        // Reemplazar texto en el contenido principal
                        foreach (var text in body.Descendants<Text>().Where(t => t.Text.Contains(param.Key)))
                        {
                            text.Text = text.Text.Replace(param.Key, param.Value);
                        }
                    }

                    // Modificar encabezados
                    var headers = wordDoc.MainDocumentPart.HeaderParts;
                    foreach (var header in headers)
                    {
                        foreach (var param in parameters)
                        {
                            if (!header.RootElement.Descendants<Text>().Any(t => t.Text.Contains(param.Key)))
                                continue;

                            foreach (var text in header.RootElement.Descendants<Text>().Where(t => t.Text.Contains(param.Key)))
                            {
                                text.Text = text.Text.Replace(param.Key, param.Value);
                            }
                        }
                    }

                    // Modificar pies de página
                    var footers = wordDoc.MainDocumentPart.FooterParts;
                    foreach (var footer in footers)
                    {
                        foreach (var param in parameters)
                        {
                            if (!footer.RootElement.Descendants<Text>().Any(t => t.Text.Contains(param.Key)))
                                continue;

                            foreach (var text in footer.RootElement.Descendants<Text>().Where(t => t.Text.Contains(param.Key)))
                            {
                                text.Text = text.Text.Replace(param.Key, param.Value);
                            }
                        }
                    }

                    // Guardar los cambios en el documento
                    wordDoc.MainDocumentPart.Document.Save();
                }

                // Retornar el archivo modificado como un arreglo de bytes
                return stream.ToArray();
            }
        }
    }
}
