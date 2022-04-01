using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;
using MimeKit.Text;
using NextSIT.Utility;
using System;
using SshNet;

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
                return true;
            }
            catch (Exception exception)
            {
                Console.WriteLine($"Ocurrio un problema al recuperar los reportes de conosce. Detalle del error => { exception.Message }");
                return false;
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
