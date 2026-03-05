using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Upload;
using Google.Apis.Util.Store;

namespace ModuloWeb.MANAGER
{
    /// <summary>
    /// Sube archivos a Google Drive usando OAuth2 con tu cuenta de Google personal.
    ///
    /// CONFIGURACIÓN (solo una vez):
    /// 1. https://console.cloud.google.com → tu proyecto → APIs y servicios → Credenciales
    /// 2. "+ Crear credencial" → "ID de cliente OAuth"
    /// 3. Tipo: "Aplicación de escritorio" → Crear
    /// 4. Descargar JSON → guardar como ModuloWeb1/Credenciales/oauth-client.json
    /// 5. La primera vez que corras la app, abrirá el navegador para que autorices con tu cuenta de Google
    /// 6. El token queda guardado localmente y no vuelve a pedir autorización
    /// </summary>
    public class GoogleDriveService
    {
        private const string FOLDER_ID = "1aTN1zSyKDh2_9f37Ytx8mpBX0mWs8qgS";
        private static readonly string[] Scopes = { DriveService.ScopeConstants.DriveFile };

        private readonly string _rutaCredenciales;
        private readonly string _rutaToken;

        public GoogleDriveService(string rutaCredenciales)
        {
            _rutaCredenciales = rutaCredenciales;
            // El token se guarda junto a las credenciales
            _rutaToken = Path.Combine(
                Path.GetDirectoryName(rutaCredenciales)!, "token_drive");
        }

        public async Task<string> SubirPdfAsync(string rutaPdf, string nombreArchivo)
        {
            if (!File.Exists(_rutaCredenciales))
                throw new FileNotFoundException(
                    $"Credenciales OAuth no encontradas en: {_rutaCredenciales}\n" +
                    "Sigue las instrucciones en GoogleDriveService.cs para configurarlo.");

            UserCredential credential;
            using (var stream = new FileStream(_rutaCredenciales, FileMode.Open, FileAccess.Read))
            {
                // La primera vez abre el navegador para autorizar; luego usa el token guardado
                credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.FromStream(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(_rutaToken, true)
                );
            }

            var service = new DriveService(new BaseClientService.Initializer
            {
                HttpClientInitializer = credential,
                ApplicationName       = "OrdenesCompra"
            });

            var metadata = new Google.Apis.Drive.v3.Data.File
            {
                Name    = nombreArchivo,
                Parents = new[] { FOLDER_ID }
            };

            using var contenido = new FileStream(rutaPdf, FileMode.Open, FileAccess.Read);
            var request = service.Files.Create(metadata, contenido, "application/pdf");
            request.Fields = "id, name, webViewLink";

            var resultado = await request.UploadAsync();
            if (resultado.Status != UploadStatus.Completed)
                throw new Exception($"Error subiendo a Drive: {resultado.Exception?.Message}");

            return request.ResponseBody?.WebViewLink ?? "";
        }
    }
}