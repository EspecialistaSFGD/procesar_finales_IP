﻿using Microsoft.Extensions.Configuration;
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
            try
            {
                conexionBd = string.IsNullOrEmpty(args[0]) ? "" : args[0];
                servidorSmtp = string.IsNullOrEmpty(args[1]) ? "" : args[1];
                puertoSmtp = string.IsNullOrEmpty(args[2]) ? "" : args[2];
                usuarioSmtp = string.IsNullOrEmpty(args[3]) ? "" : args[3];
                claveSmtp = string.IsNullOrEmpty(args[4]) ? "" : args[4];
                deSmtp = string.IsNullOrEmpty(args[5]) ? "" : args[5];
                paraSmtp = string.IsNullOrEmpty(args[6]) ? "" : args[6];
                hostSftp = string.IsNullOrEmpty(args[7]) ? "" : args[7];
                puertoSftp = string.IsNullOrEmpty(args[8]) ? "" : args[8];
                usuarioSftp = string.IsNullOrEmpty(args[9]) ? "" : args[9];
                claveSftp = string.IsNullOrEmpty(args[10]) ? "" : args[10];
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
                Password = !string.IsNullOrEmpty(claveSftp) ? claveSftp : Configuration.GetSection("ClaveSftp").Value.ToString()
            };
            EjecutarProceso(conexionBd, mailConfiguration, sftprequest)
                .GetAwaiter()
                .GetResult();
        }


        static async Task EjecutarProceso(string conexion, Mail mail, FileManager.SftpRequest sftp)
        {
            try
            {
                Console.WriteLine($"--------------------------------------------------------------------------------");
                Console.WriteLine($"    Proceso de extraccion de reportes desde CONOSCE para el anio configurado");
                Console.WriteLine($"---------------------------------------------------------------------------------");
                var mensajeRespuesta = @"<h3>Proceso de extraccion de reportes desde CONOSCE</h3><p>mensaje_respuesta</p>";
                var proxyManager = ProxyManager.GetNewProxyManager();
                var typeConvertionsManager = TypeConvertionManager.GetNewTypeConvertionManager();
                var fileManager = FileManager.GetNewFileManager();
                var request = new ProxyManager.Request();
                var repositorio = new Repositorio(conexion);

                return;
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception.Message);
                throw;
            }
        }

    }
}