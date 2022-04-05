using Microsoft.Extensions.Configuration;
using NextSIT.Utility;
using System;
using System.Threading.Tasks;

namespace ProcesarConosce
{
    internal class Program
    {
        private static IConfigurationRoot Configuration { get; set; }
        static void Main(string[] args)
        {
            var anioProceso = "";
            var conexionBd = "";
            var servidorSmtp = "";
            var puertoSmtp = "";
            var usuarioSmtp = "";
            var claveSmtp = "";
            var deSmtp = "";
            var paraSmtp = "";
            var hostSftp = "";
            var puertoSftp = "";
            var usuarioSftp = "";
            var claveSftp = "";
            var rutaOrigen = "";
            var rutaDestino = "";

            try
            {
                anioProceso = string.IsNullOrEmpty(args[0]) ? "" : args[0];
                conexionBd = string.IsNullOrEmpty(args[1]) ? "" : args[1];
                servidorSmtp = string.IsNullOrEmpty(args[2]) ? "" : args[2];
                puertoSmtp = string.IsNullOrEmpty(args[3]) ? "" : args[3];
                usuarioSmtp = string.IsNullOrEmpty(args[4]) ? "" : args[4];
                claveSmtp = string.IsNullOrEmpty(args[5]) ? "" : args[5];
                deSmtp = string.IsNullOrEmpty(args[6]) ? "" : args[6];
                paraSmtp = string.IsNullOrEmpty(args[7]) ? "" : args[7];
                hostSftp = string.IsNullOrEmpty(args[8]) ? "" : args[8];
                puertoSftp = string.IsNullOrEmpty(args[9]) ? "" : args[9];
                usuarioSftp = string.IsNullOrEmpty(args[10]) ? "" : args[10];
                claveSftp = string.IsNullOrEmpty(args[11]) ? "" : args[11];
                rutaOrigen = string.IsNullOrEmpty(args[12]) ? "" : args[12];
                rutaDestino = string.IsNullOrEmpty(args[13]) ? "" : args[13];
            }
            catch (Exception)
            {
                Console.WriteLine($"Algunos parametros no han sido transferidos a la consola, se utilizaran los valores por defecto");
            }

            IConfigurationBuilder builder = new ConfigurationBuilder()
                .AddJsonFile("appsettings.json")
                .AddEnvironmentVariables();

            Configuration = builder.Build();

            var mailConfiguration = new Mail
            {
                Servidor = !string.IsNullOrEmpty(servidorSmtp) ? servidorSmtp : Configuration.GetSection("Servidor").Value.ToString(),
                Puerto = int.Parse(!string.IsNullOrEmpty(puertoSmtp) ? puertoSmtp : Configuration.GetSection("Puerto").Value.ToString()),
                Usuario = !string.IsNullOrEmpty(usuarioSmtp) ? usuarioSmtp : Configuration.GetSection("Usuario").Value.ToString(),
                Clave = !string.IsNullOrEmpty(claveSmtp) ? claveSmtp : Configuration.GetSection("Clave").Value.ToString(),
                De = !string.IsNullOrEmpty(deSmtp) ? deSmtp : Configuration.GetSection("De").Value.ToString(),
                Para = !string.IsNullOrEmpty(paraSmtp) ? paraSmtp : Configuration.GetSection("Para").Value.ToString()
            };

            var sftprequest = new FileManager.SftpRequest()
            {
                Host = !string.IsNullOrEmpty(hostSftp) ? hostSftp : Configuration.GetSection("HostSftp").Value.ToString(),
                Port = int.Parse(!string.IsNullOrEmpty(puertoSftp) ? puertoSftp : Configuration.GetSection("PuertoSftp").Value.ToString()),
                Username = !string.IsNullOrEmpty(usuarioSftp) ? usuarioSftp : Configuration.GetSection("UsuarioSftp").Value.ToString(),
                Password = !string.IsNullOrEmpty(claveSftp) ? claveSftp : Configuration.GetSection("ClaveSftp").Value.ToString(),
                SourceRoute = !string.IsNullOrEmpty(rutaOrigen) ? rutaOrigen : Configuration.GetSection("RutaOrigen").Value.ToString(),
                DestinationRoute = !string.IsNullOrEmpty(rutaDestino) ? rutaDestino : Configuration.GetSection("RutaDestino").Value.ToString()
            };
            EjecutarProceso(
                !string.IsNullOrEmpty(conexionBd) ? conexionBd : Configuration.GetConnectionString("ConexionPcm"),
                mailConfiguration, 
                sftprequest,
                !string.IsNullOrEmpty(anioProceso) ? anioProceso : Configuration.GetSection("AnioCarga").Value.ToString()
                )
                .GetAwaiter()
                .GetResult();
        }


        static async Task EjecutarProceso(string conexion, Mail mail, FileManager.SftpRequest sftp, string anioProceso)
        {
            var mensajeRespuesta = @"<h3>Proceso de extraccion de reportes desde CONOSCE</h3><h4>Procesos conformes</h4>mensaje_ok<h4>Procesos Errados</h4>mensaje_error";

            var repositorio = new Repositorio(conexion);
            try
            {
                Console.WriteLine($"--------------------------------------------------------------------------------");
                Console.WriteLine($"    Proceso de extraccion de reportes desde CONOSCE para el anio configurado");
                Console.WriteLine($"---------------------------------------------------------------------------------");

                var rutaBaseOrigen = sftp.SourceRoute;
                var rutaBaseDestino = sftp.DestinationRoute;

                Console.WriteLine($"Descargando los reportes de contratos de CONOSCE");
                sftp.SourceRoute = $"{rutaBaseOrigen}conosce_contratos_s3/";
                sftp.DestinationRoute = $"{rutaBaseDestino}conosce_contratos_s3\\";

                var haDescargadoContratos = true;//repositorio.RecuperarReportesConosce(sftp);

                if (!haDescargadoContratos)
                {
                    Console.WriteLine("Los reportes de contratos no han podido ser descargados correctamente de CONOSCE.");
                }
                else
                {
                    Console.WriteLine("Los reportes de contratos han sido descargados correctamente de CONOSCE.");
                    Console.WriteLine("Cargando reportes descargados a la tabla de contratos.");
                    var haProcesadoDatosContratos = await repositorio.CargarDataTablaContrato($"{sftp.DestinationRoute}{anioProceso}\\", anioProceso);
                    Console.WriteLine(haProcesadoDatosContratos ? "Los contratos se cargaron correctamente en la tabla de contratos." : "No se cargó la información correctamente en la tabla de contratos.");
                }

                Console.WriteLine($"Descargando los reportes de cronogramas de CONOSCE");
                sftp.SourceRoute = $"{rutaBaseOrigen}conosce_cronograma/";
                sftp.DestinationRoute = $"{rutaBaseDestino}conosce_cronograma\\";

                var haDescargadoCronogramas = true;//repositorio.RecuperarReportesConosce(sftp);

                if (!haDescargadoCronogramas)
                {
                    Console.WriteLine("Los reportes de cronogramas no han podido ser descargados correctamente de CONOSCE.");
                }
                else
                {

                    Console.WriteLine($"Los reportes de cronogramas han sido descargados correctamente de CONOSCE");
                    Console.WriteLine("Cargando reportes descargados a la tabla de cronogramas.");
                    var haProcesadoDatosCronogramas = await repositorio.CargarDataTablaCronograma($"{sftp.DestinationRoute}{anioProceso}\\", anioProceso);
                    Console.WriteLine(haProcesadoDatosCronogramas ? "Los cronogramas se cargaron correctamente en la tabla de cronogramas." : "No se cargó la información correctamente en la tabla de cronogramas.");
                }

                Console.WriteLine($"Descargando los reportes especiales de CONOSCE");
                sftp.SourceRoute = $"{rutaBaseOrigen}reporte_especial/";
                sftp.DestinationRoute = $"{rutaBaseDestino}reporte_especial\\";

                var haDescargadoReporteEspecial = true;//repositorio.RecuperarReportesConosce(sftp);

                if (!haDescargadoReporteEspecial)
                {
                    Console.WriteLine("Los reportes especiales no han podido ser descargados correctamente de CONOSCE.");
                }
                else
                {
                    Console.WriteLine($"Los reportes especiales han sido descargados correctamente de CONOSCE");
                    Console.WriteLine("Cargando reportes descargados a la tabla de snip.");
                    var haProcesadoDatosSnip = await repositorio.CargarDataTablaSnip(sftp.DestinationRoute, anioProceso);
                    Console.WriteLine(haProcesadoDatosSnip ? "Los snip se cargaron correctamente en la tabla de snip." : "No se cargó la información correctamente en la tabla de snip.");
                }


                var detalle = $"El proceso de extracción de reportes desde CONOSCE se ha completado correctamente, puede revisar los archivos en el servidor o consultar directamente desde la base de datos y/o Web del Mapa Inversiones</p>";
                mensajeRespuesta = mensajeRespuesta.Replace("mensaje_respuesta", detalle);
                repositorio.SendMail(mail, "Proceso de Carga Masiva de Datos de Proyectos", mensajeRespuesta);

                return;
            }
            catch (Exception exception)
            {
                var detalle = $"Ocurrió un problema durante el proceso de extracción de reportes. Detalle del error : {exception.Message}";
                mensajeRespuesta = mensajeRespuesta.Replace("mensaje_respuesta", detalle);
                repositorio.SendMail(mail, "Proceso de Carga Masiva de Datos de Proyectos", mensajeRespuesta);

                Console.WriteLine(exception.Message);
                throw;
            }
        }

    }
}
