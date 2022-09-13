using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Watermark
{
    class ErrorHandler
    {
        private SmtpClient smtpClient = new SmtpClient("mail.cftrail.cz");

        /// <summary>
        /// Posle info o chybe na IT
        /// </summary>
        /// <param name="method"></param>
        /// <param name="errorMessage"></param>
        public void SendError(string method, string errorMessage)
        {
            MailMessage mail = new MailMessage("intra@interfracht.cz", "it@interfracht.cz", String.Format("Excel.Watermark: SERVISNÍ INFO (#Chyba ve funkci {0}) #EXCEL.WATERMARK_ERROR", method), errorMessage);
            mail.IsBodyHtml = false;
            smtpClient.Send(mail);
        }
    }
}
