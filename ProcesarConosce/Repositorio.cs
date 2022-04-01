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

namespace ProcesarConosce
{
    public class Repositorio
    {
        private readonly string Conexion = "";
        private readonly FileManager fileManager;
        private readonly TypeConvertionManager typeConvertionsManager;

        public Repositorio(string conexion)
        {
            Conexion = conexion;
            fileManager = FileManager.GetNewFileManager();
            typeConvertionsManager = TypeConvertionManager.GetNewTypeConvertionManager();
        }

        public bool RecuperarReportesConosce(FileManager.SftpRequest sftpRequest)
        {
            try
            {
                Directory.CreateDirectory(sftpRequest.DestinationRoute);

                using (var client = new SftpClient(sftpRequest.Host, sftpRequest.Port ?? 22, sftpRequest.Username, sftpRequest.Password))
                {
                    try
                    {
                        client.Connect();
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


        public void DescargarDirectorio(SftpClient clienteSftp, string rutaOrigen, string rutaDestino)
        {
            Directory.CreateDirectory(rutaDestino);
            IEnumerable<SftpFile> archivos = clienteSftp.ListDirectory(rutaOrigen);
            foreach (SftpFile archivoSftp in archivos)
            {
                if ((archivoSftp.Name != ".") && (archivoSftp.Name != ".."))
                {
                    string archivoOrigen = rutaOrigen + "/" + archivoSftp.Name;
                    string archivoDestino = Path.Combine(rutaDestino, archivoSftp.Name);
                    if (archivoSftp.IsDirectory)
                    {
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

        //Paso 8.- Enviar mail por concepto de error o éxito
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
