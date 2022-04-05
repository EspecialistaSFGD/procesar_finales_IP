using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;
using MimeKit.Text;
using NextSIT.Utility;
using System;
using Renci.SshNet;
using System.Collections.Generic;
using Renci.SshNet.Sftp;
using System.IO;
using System.Linq;
using System.Data;
using System.Threading.Tasks;
using System.Data.SqlClient;
using Dapper;

namespace ProcesarConosce
{
    public class Repositorio
    {
        private readonly string Conexion = "";
        private readonly int TiempoEsperaCargadoMasivo;
        private readonly int BatchSize;

        public Repositorio(string conexion)
        {
            Conexion = conexion;
            TiempoEsperaCargadoMasivo = 10000;
            BatchSize = 50000;
        }

        //Paso 1.- Recupera reportes desde el SFTP de CONOSCE
        public bool RecuperarReportesConosce(FileManager.SftpRequest sftpRequest)
        {
            try
            {
                Console.WriteLine($"Directorio en la ubicacion {sftpRequest.DestinationRoute} generado correctamente");

                using (var client = new SftpClient(sftpRequest.Host, sftpRequest.Port ?? 22, sftpRequest.Username, sftpRequest.Password))
                {
                    try
                    {
                        client.Connect();
                        Console.WriteLine($"Inicio de descarga de archivos y directorios del directorio {sftpRequest.SourceRoute}");
                        DescargarDirectorio(client, sftpRequest.SourceRoute, sftpRequest.DestinationRoute);
                        client.Disconnect();
                    }
                    catch (Exception exception)
                    {
                        throw new Exception(exception.Message);
                    }
                }
                return true;
            }
            catch (Exception exception)
            {
                Console.WriteLine($"Ocurrio un problema al recuperar los reportes de conosce. Detalle del error => { exception.Message }");
                return false;
            }
        }

        //Pase 2.- Prepara datos de contratos desde los reportes descargados de CONOSCE
        public async Task<bool> CargarDataTablaContrato(string urlBase, string anioProceso)
        {
            try
            {

                var directorio = new DirectoryInfo(urlBase);
                var dataTableContrato = new DataTable();
                var indice = 0;
                foreach (var archivo in directorio.GetFiles($"CONOSCE_CONTRATOS{anioProceso}_?.xlsx"))
                {
                    if (indice == 0)
                    {
                        dataTableContrato = ObtenerDataTableDesdeExcel(archivo.FullName, Tuple.Create(2, 1), true);
                        var columna = dataTableContrato.Columns.Add("ANIO", Type.GetType("System.Int32"), anioProceso);
                        columna.SetOrdinal(0);
                    }
                    else
                    {
                        var dataTableContratoDetalle = new DataTable();
                        dataTableContratoDetalle = ObtenerDataTableDesdeExcel(archivo.FullName, Tuple.Create(2, 1), true);
                        var column = dataTableContratoDetalle.Columns.Add("ANIO", Type.GetType("System.Int32"), anioProceso);
                        column.SetOrdinal(0);

                        dataTableContrato.Merge(dataTableContratoDetalle);

                    }
                    indice++;
                }

                var contratosHanSidoEliminados = await EliminarContratos();

                if (!contratosHanSidoEliminados)
                {
                    throw new Exception("Los contratos existentes en la tabla no han podido ser eliminados.");
                }

                var contratosHanSidoInsertados = InsercionMasivaContratos(dataTableContrato);

                if (!contratosHanSidoInsertados)
                {
                    throw new Exception("Los contratos extraidos del reporte de CONOSCE no han podido ser registrados.");
                }

                return true;
            }
            catch (Exception exception)
            {
                throw new Exception($"No se ha completado el proceso de carga masiva de contratos: Detalle del error => {exception.Message}" ,exception);
            }

        }

        //Paso 3.- Eliminar contratos para el anio en proceso
        public async Task<bool> EliminarContratos()
        {
            using var conexionSql = new SqlConnection(Conexion);
            try
            {
                conexionSql.Open();

                var respuesta = await conexionSql.QueryAsync("dbo.05A_EliminarContratoConosce", commandType: CommandType.StoredProcedure, commandTimeout: 1200);
                return true;
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception.Message);
                throw;
            }
            finally
            {
                conexionSql.Close();
            }

        }

        //Paso 4.- Registrar Contratos para el anio de proceso de forma masiva
        public bool InsercionMasivaContratos(DataTable valores)
        {
            using var conexionSql = new SqlConnection(Conexion);
            conexionSql.Open();

            using SqlBulkCopy bulkCopy = new(conexionSql);
            bulkCopy.BulkCopyTimeout = TiempoEsperaCargadoMasivo;
            bulkCopy.BatchSize = BatchSize;
            bulkCopy.DestinationTableName = "dbo.CONOSCE_CONTRATOS";

            try
            {
                bulkCopy.WriteToServer(valores);
                return true;
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception.Message);
                throw;

            }
            finally
            {
                conexionSql.Close();
            }
        }

        //Pase 5.- Prepara datos de cronogramas desde los reportes descargados de CONOSCE
        public async Task<bool> CargarDataTablaCronograma(string urlBase, string anioProceso)
        {
            try
            {

                var directorio = new DirectoryInfo(urlBase);
                var dataTableCronogramas = new DataTable();
                var indice = 0;
                foreach (var archivo in directorio.GetFiles($"CONOSCE_CRONOGRAMA{anioProceso}_?.xlsx"))
                {
                    if (indice == 0)
                    {
                        dataTableCronogramas = ObtenerDataTableDesdeExcel(archivo.FullName, Tuple.Create(2, 1), true);
                        var columna = dataTableCronogramas.Columns.Add("ANIO", Type.GetType("System.Int32"), anioProceso);
                        columna.SetOrdinal(0);
                    }
                    else
                    {
                        var dataTableCronogramasDetalle = new DataTable();
                        dataTableCronogramasDetalle = ObtenerDataTableDesdeExcel(archivo.FullName, Tuple.Create(2, 1), true);
                        var column = dataTableCronogramasDetalle.Columns.Add("ANIO", Type.GetType("System.Int32"), anioProceso);
                        column.SetOrdinal(0);

                        dataTableCronogramas.Merge(dataTableCronogramasDetalle);

                    }
                    indice++;
                }

                var cronogramasHanSidoEliminados = await EliminarCronogramas();

                if (!cronogramasHanSidoEliminados)
                {
                    throw new Exception("Los cronogramas existentes en la tabla no han podido ser eliminados.");
                }

                var cronogramasHanSidoRegistrados = InsercionMasivaCronogramas(dataTableCronogramas);

                if (!cronogramasHanSidoRegistrados)
                {
                    throw new Exception("Los cronogramas extraidos del reporte de CONOSCE no han podido ser registrados.");
                }

                return true;
            }
            catch (Exception exception)
            {
                throw new Exception($"No se ha completado el proceso de carga masiva de cronogramas: Detalle del error => {exception.Message}", exception);
            }

        }

        //Paso 6.- Eliminar cronogramas de convocatoria para el anio en proceso
        public async Task<bool> EliminarCronogramas()
        {
            using var conexionSql = new SqlConnection(Conexion);
            try
            {
                conexionSql.Open();

                var respuesta = await conexionSql.QueryAsync("dbo.05B_EliminarCronogramaConosce", commandType: CommandType.StoredProcedure, commandTimeout: 1200);
                return true;
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception.Message);
                throw;
            }
            finally
            {
                conexionSql.Close();
            }

        }
       
        //Paso 7.- Registrar Cronogramas para el anio de proceso de forma masiva
        public bool InsercionMasivaCronogramas(DataTable valores)
        {
            using var conexionSql = new SqlConnection(Conexion);
            conexionSql.Open();

            using SqlBulkCopy bulkCopy = new(conexionSql);
            bulkCopy.BulkCopyTimeout = TiempoEsperaCargadoMasivo;
            bulkCopy.BatchSize = BatchSize;
            bulkCopy.DestinationTableName = "dbo.CONOSCE_CRONOGRAMA";

            try
            {
                bulkCopy.WriteToServer(valores);
                return true;
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception.Message);
                throw;

            }
            finally
            {
                conexionSql.Close();
            }
        }

        //Pase 8.- Prepara datos de cronogramas desde los reportes descargados de CONOSCE
        public async Task<bool> CargarDataTablaSnip(string urlBase, string anioProceso)
        {
            try
            {
                var dataTableSnip = new DataTable();
                dataTableSnip = ObtenerDataTableDesdeExcel($"{urlBase}SNIP SEACE.xlsx", Tuple.Create(1, 1), true);

                var snipHanSidoEliminados = await EliminarSnip();

                if (!snipHanSidoEliminados)
                {
                    throw new Exception("Los snip existentes en la tabla no han podido ser eliminados.");
                }

                var snipHanSidoInsertados = InsercionMasivaSnip(dataTableSnip);

                if (!snipHanSidoInsertados)
                {
                    throw new Exception("Los snip extraidos del reporte especial de CONOSCE no han podido ser registrados.");
                }

                return true;
            }
            catch (Exception exception)
            {
                throw new Exception($"No se ha completado el proceso de carga masiva de snip: Detalle del error => {exception.Message}", exception);
            }

        }

        //Paso 9.- Eliminar snip para el anio en proceso
        public async Task<bool> EliminarSnip()
        {
            using var conexionSql = new SqlConnection(Conexion);
            try
            {
                conexionSql.Open();

                var respuesta = await conexionSql.QueryAsync("dbo.05C_EliminarSnipConosce", commandType: CommandType.StoredProcedure, commandTimeout: 1200);
                return true;
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception.Message);
                throw;
            }
            finally
            {
                conexionSql.Close();
            }

        }

        //Paso 10.- Registrar SNIP de forma masiva
        public bool InsercionMasivaSnip(DataTable valores)
        {
            using var conexionSql = new SqlConnection(Conexion);
            conexionSql.Open();

            using SqlBulkCopy bulkCopy = new(conexionSql);
            bulkCopy.BulkCopyTimeout = TiempoEsperaCargadoMasivo;
            bulkCopy.BatchSize = BatchSize;
            bulkCopy.DestinationTableName = "dbo.SNIP_SEACE";

            try
            {
                bulkCopy.WriteToServer(valores);
                return true;
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception.Message);
                throw;

            }
            finally
            {
                conexionSql.Close();
            }
        }

        //Paso 11.- Descargar directorio especifico desde el SFTP
        public void DescargarDirectorio(SftpClient clienteSftp, string rutaOrigen, string rutaDestino)
        {
            Directory.CreateDirectory(rutaDestino);
            IEnumerable<SftpFile> archivos = clienteSftp.ListDirectory(rutaOrigen);
            Console.WriteLine($"Listado de archivos en {rutaOrigen}:\n {string.Join('\n',archivos.Select(x => $"  *  {x.Name}"))}");
            foreach (SftpFile archivoSftp in archivos)
            {
                if ((archivoSftp.Name != ".") && (archivoSftp.Name != ".."))
                {
                    string archivoOrigen = rutaOrigen + archivoSftp.Name;
                    string archivoDestino = rutaDestino + archivoSftp.Name;
                    if (archivoSftp.IsDirectory)
                    {
                        archivoOrigen += "/";
                        archivoDestino += "\\";
                        DescargarDirectorio(clienteSftp, archivoOrigen, archivoDestino);
                    }
                    else
                    {
                        using (Stream fileStream = File.Create(archivoDestino))
                        {
                            clienteSftp.DownloadFile(archivoOrigen, fileStream);
                        }
                    }
                }
            }
        }

        //Paso 12.- Convertir excel a DataTable
        public DataTable ObtenerDataTableDesdeExcel(string rutaArchivo, Tuple<int, int> posicionInicioLibro, bool tieneCabecera = true)
        {
            using (var excelPackage = new OfficeOpenXml.ExcelPackage())
            {
                using (var stream = File.OpenRead(rutaArchivo))
                {
                    excelPackage.Load(stream);
                }
                var hojaCalculo = excelPackage.Workbook.Worksheets.First();
                DataTable tabla = new DataTable();
                foreach (var primeraFilaCelda in hojaCalculo.Cells[posicionInicioLibro.Item1, posicionInicioLibro.Item2, posicionInicioLibro.Item1, hojaCalculo.Dimension.End.Column])
                {
                    tabla.Columns.Add(tieneCabecera ? primeraFilaCelda.Text : string.Format("Column {0}", primeraFilaCelda.Start.Column));
                }
                var filaInicial = tieneCabecera ? (posicionInicioLibro.Item1 + 1) : posicionInicioLibro.Item1;
                for (int numeroFila = filaInicial; numeroFila <= hojaCalculo.Dimension.End.Row; numeroFila++)
                {
                    var fila = hojaCalculo.Cells[numeroFila, posicionInicioLibro.Item2, numeroFila, hojaCalculo.Dimension.End.Column];
                    DataRow row = tabla.Rows.Add();
                    foreach (var celda in fila)
                    {
                        row[celda.Start.Column - 1] = celda.Text;
                    }
                }
                return tabla;
            }
        }

        //Paso 13.- Enviar mail por concepto de error o éxito
        public void SendMail(Mail configuracion, string asunto, string mensaje)
        {
            try
            {
                // create message
                var email = new MimeMessage();
                email.Sender = MailboxAddress.Parse(configuracion.De);
                string[] destinatarios = configuracion.Para.Split(";");

                foreach (string destinatario in destinatarios) email.To.Add(MailboxAddress.Parse(destinatario));
                email.Subject = asunto;//"Notificaciones Mapa Inversiones - Sincronizacion de Datos del MEF";
                email.Body = new TextPart(TextFormat.Html) { Text = mensaje };

                // send email
                using var smtp = new SmtpClient();
                smtp.Connect(configuracion.Servidor, configuracion.Puerto, SecureSocketOptions.StartTls);
                smtp.Authenticate(configuracion.De, configuracion.Clave);
                smtp.Send(email);
                smtp.Disconnect(true);
            }
            catch (Exception exception)
            {
                Console.WriteLine($"Ocurrio un problema al enviar la notificacion de la carga fallida. Detalle del error => { exception.Message }");
            }
        }

    }
}
