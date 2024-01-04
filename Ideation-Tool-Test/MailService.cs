using System.Net;
using System.Net.Mail;

namespace Ideation_Tool_Test;

public class MailService
{
    public static void SendEmail(string to, string subject, string body)
    {
        try
        {
            var fromAddress = new MailAddress("mailserviceabbott@gmail.com", "Abbott HQ"); 
            var toAddress = new MailAddress(to); 
            const string fromPassword = "ppcc wynp fhbi nexc";

            var smtp = new SmtpClient 
            {
                Host = "smtp.gmail.com", // Vervang door de SMTP server van je provider
                Port = 587, // Standaard SMTP poort, kan variëren afhankelijk van je provider
                EnableSsl = true, // Afhankelijk van je provider
                DeliveryMethod = SmtpDeliveryMethod.Network, 
                UseDefaultCredentials = false, 
                Credentials = new NetworkCredential(fromAddress.Address, fromPassword)
            };

            using var message = new MailMessage(fromAddress, toAddress)
            {
                Subject = subject, 
                Body = body,
                IsBodyHtml = true
            };
            
            smtp.Send(message); 
            Console.WriteLine("E-mail succesvol verzonden.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Fout bij het verzenden van e-mail: " + ex.Message);
        }
    }

}