using System.Diagnostics;

namespace ModuloWeb.MANAGER
{
    /// <summary>
    /// Convierte un archivo .xlsx a .pdf usando LibreOffice headless.
    /// Requiere LibreOffice instalado: https://www.libreoffice.org/download/download/
    /// </summary>
    public class PdfConverterService
    {
        // Rutas típicas de LibreOffice en Windows
        private static readonly string[] _rutasCandidatas = new[]
        {
            @"C:\Program Files\LibreOffice\program\soffice.exe",
            @"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
            @"C:\LibreOffice\program\soffice.exe",
        };

        public static string? EncontrarLibreOffice()
        {
            foreach (var ruta in _rutasCandidatas)
                if (File.Exists(ruta)) return ruta;
            return null;
        }

        /// <summary>
        /// Convierte xlsxPath a PDF en la misma carpeta.
        /// Devuelve la ruta del PDF generado.
        /// </summary>
        public static string ConvertirAPdf(string xlsxPath)
        {
            string? soffice = EncontrarLibreOffice();
            if (soffice == null)
                throw new InvalidOperationException(
                    "LibreOffice no encontrado. Instálalo desde https://www.libreoffice.org/download/download/ " +
                    "y vuelve a intentarlo.");

            string carpeta = Path.GetDirectoryName(xlsxPath)!;

            var psi = new ProcessStartInfo
            {
                FileName               = soffice,
                Arguments              = $"--headless --convert-to pdf \"{xlsxPath}\" --outdir \"{carpeta}\"",
                RedirectStandardOutput = true,
                RedirectStandardError  = true,
                UseShellExecute        = false,
                CreateNoWindow         = true
            };

            using var proc = Process.Start(psi)
                ?? throw new Exception("No se pudo iniciar LibreOffice.");

            proc.WaitForExit(60_000); // máximo 60 s

            string pdfPath = Path.ChangeExtension(xlsxPath, ".pdf");
            if (!File.Exists(pdfPath))
                throw new Exception(
                    $"LibreOffice no generó el PDF. Error: {proc.StandardError.ReadToEnd()}");

            return pdfPath;
        }
    }
}