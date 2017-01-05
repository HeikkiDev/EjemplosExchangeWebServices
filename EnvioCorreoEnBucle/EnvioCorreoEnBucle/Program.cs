using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace EnvioCorreoEnBucle
{
    class Program
    {
        static string ExceptionMessage { get; set; }

        static void Main(string[] args)
        {
            Console.WriteLine("\n Enviando 200 emails...");

            ExchangeService _service = new ExchangeService();
            List<EmailMessage> listaMensajes = new List<EmailMessage>();

            try
            {
                string password = "YOUR_PASSWORD";
                string accentureAccount = "ACCENTURE_ACCOUNT";
                string sendTo = "TO_Recipient";

                //Provide the account user name
                _service.Credentials = new WebCredentials(accentureAccount, password);
                _service.Url = new Uri("https://email.o365.accenture.com/EWS/Exchange.asmx");

                // Enviar 200 emails
                for (int i = 1; i <= 200; i++)
                {
                    EmailMessage message = new EmailMessage(_service);
                    message.Subject = "Correo número " + i;
                    message.Body = "Este es el Correo de prueba de rendimiento número " + i;
                    message.ToRecipients.Add(sendTo);
                    listaMensajes.Add(message);
                }

                foreach (var item in listaMensajes)
                {
                    Thread t = new Thread(new ParameterizedThreadStart(EnviarCorreoAsync));
                    t.Start(item);
                }
                
            }
            catch (Exception ex)
            {
                _service = null;
                ExceptionMessage = ex.Message;
            }

            Console.WriteLine("\n\n Envío finalizado");
            Console.ReadKey();
        }

        private static void EnviarCorreoAsync(object obj)
        {
            EmailMessage email = (EmailMessage)obj;
            email.Send();
        }

    }
}
