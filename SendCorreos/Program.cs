using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Mail;
using System.Net.Mime;
using System.Net;
using System.IO;
using System.Security.Claims;
using FluentValidation;

public class CorreoValidator : AbstractValidator<string>
{
    public CorreoValidator()
    {
        RuleFor(correo => correo.Trim())
            .EmailAddress().WithMessage("Correo inválido")
            .Must(correo => !correo.Contains("..")).WithMessage("Correo no debe contener '..'")
            .Must(correo => !correo.Contains(" ")).WithMessage("Correo no debe contener espacios");
    }
}


namespace SendCorreos
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string connectionString = ConfigurationSettings.AppSettings["conexion_BD_Ambulatorio"];
            SqlConnection cnn = new SqlConnection(connectionString);
            //string strsql = " select cast([nsecuencia] as varchar(20)) as nsecuencia , [identificador], isnull([correo],'') as correo, [fecha], [hora], [direccion], [sucursal], [ubicacion], [subject], [body], isnull([nombre], '') as nombre, tipo, sistema from  [dbo].[CORREOS_SEND]  where ( tipo = 'CNF' or tipo = 'ANU'  ) and (Cestado  = 0)";
            //string strsql = "select cast([nsecuencia] as varchar(20)) as nsecuencia, [identificador], isnull([correo],'') as correo, [fecha], [hora], [direccion], [sucursal], [ubicacion], [subject], [body], isnull([nombre], '') as nombre, [subtipoobj], tipo, sistema from [dbo].[CORREOS_SEND] where (tipo = 'CNF' or tipo = 'ANU') and (Cestado = 0)";
            string strsql = "select cast([nsecuencia] as varchar(20)) as nsecuencia, [identificador], isnull([correo],'') as correo, [fecha], [hora], [direccion], [sucursal], [ubicacion], [subject], [body], isnull([nombre], '') as nombre, [subtipoobj], tipo, sistema, isnull(cod_cita, '') as cod_cita from [dbo].[CORREOS_SEND] where (tipo = 'CNF' or tipo = 'ANU') and (Cestado = 0)";

            SqlDataAdapter da = new SqlDataAdapter(strsql, cnn);
            DataSet ds = new DataSet();
            da.Fill(ds);

            var validator = new CorreoValidator();

            if (ds != null && ds.Tables[0].Rows.Count > 0)
            {
                for (int j = 0; j < ds.Tables[0].Rows.Count; j++)
                {
                    string correo = ds.Tables[0].Rows[j].Field<string>("correo").ToString().ToLower().Trim();
                    correo = correo.Replace(",", ".");


                    var validationResult = validator.Validate(correo);
                    if (!validationResult.IsValid)
                    {
                        correo = "agenda@clini.cl";
                    }

                    if (correo.Trim() != "")
                    {
                        string nsecuencia = ds.Tables[0].Rows[j].Field<string>("nsecuencia").ToString();
                        string identificador = ds.Tables[0].Rows[j].Field<string>("identificador").ToString();
                        string fecha = ds.Tables[0].Rows[j].Field<string>("fecha").ToString();
                        string hora = ds.Tables[0].Rows[j].Field<string>("hora").ToString();
                        string direccion = ds.Tables[0].Rows[j].Field<string>("direccion").ToString();
                        //string sucursal = ds.Tables[0].Rows[j].Field<string>("sucursal").ToString(); //
                        string sucursal = ds.Tables[0].Rows[j].Field<string>("sucursal")?.ToString() ?? string.Empty;
                        string ubicacion = ds.Tables[0].Rows[j].Field<string>("ubicacion")?.ToString() ?? string.Empty;
                        string subject = ds.Tables[0].Rows[j].Field<string>("subject").ToString();
                        string body = ds.Tables[0].Rows[j].Field<string>("body").ToString();
                        string nombre = ds.Tables[0].Rows[j].Field<string>("nombre").ToString();
                        string tipo = ds.Tables[0].Rows[j].Field<string>("tipo").ToString();
                        string sistema = ds.Tables[0].Rows[j].Field<string>("sistema").ToString();
                        string subtipoobj = Convert.ToString(ds.Tables[0].Rows[j]["subtipoobj"]);
                        string codCita = ds.Tables[0].Rows[j].Field<string>("cod_cita").ToString();

                        SincronizarCitaYCorreosSend(cnn, nsecuencia, ref sistema, ref subtipoobj, ref codCita, identificador);

                        if (tipo == "CNF")
                        {
                            string retorno_correo = EnvioCorreoC(subject, correo, nombre, fecha, hora, direccion, sucursal, ubicacion, sistema, subtipoobj, codCita);
                            string ret_upd = UpdateEstadoEnvio(nsecuencia);
                        }

                        if (tipo == "ANU")
                        {
                            string retorno_correo = EnvioCorreoAnula("Anulación Cita", correo, fecha, hora, sucursal, ubicacion, sistema);
                            string ret_upd = UpdateEstadoEnvio(nsecuencia);
                        }
                    }
                }
            }
        }





        private static string EnvioCorreoC(string subject, string tcorreo, string tnombre, string fecha, string hora, string direccion, string sucursal, string ubicacion, string sistema, string subtipoobj, string codCita)
        {
            string senderEmail = System.Configuration.ConfigurationSettings.AppSettings.Get("mail_usuario");
            string senderPassword = System.Configuration.ConfigurationSettings.AppSettings.Get("mail_clave");

            string smtpServer = "in.smtpok.com";
            int smtpPort = 587;
            string smtpUser = "s182692_3";
            string smtpPassword = "7awmO.S4Xv";
            //string senderEmail = "agenda@clini.cl";
            string token = GenerarToken(codCita);

            // Recipient's email address
            string recipientEmail = tcorreo;

            // Create the email message
            MailMessage mail = new MailMessage(senderEmail, recipientEmail);
            mail.Subject = subject;

            string imagePath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory) + @"img\logo.png";
            string imageFileName = "logo.png";

            string imagePath2 = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory) + @"img\facebook@2x.png";
            string imageFileName2 = "facebook@2x.png";

            string imagePath3 = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory) + @"img\linkedin@2x.png";
            string imageFileName3 = "linkedin@2x.png";

            string imagePath4 = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory) + @"img\instagram@2x.png";
            string imageFileName4 = "instagram@2x.png";

            string imagePath5 = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory) + @"img\youtube@2x.png";
            string imageFileName5 = "youtube@2x.png";

            string imagePath6 = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory) + @"img\tac_scanner.png";
            string imageFileName6 = "tac_scanner.png";

            string imagePath7 = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory) + @"img\resonancia.png";
            string imageFileName7 = "resonancia.png";

            string imagePath8 = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory) + @"img\ecografia4d5d.png";
            string imageFileName8 = "ecografia4d5d.png";

            string imagePath9 = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory) + @"img\mamografia.png";
            string imageFileName9 = "mamografia.png";

            string imagePath10 = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory) + @"img\holter_presion.png";
            string imageFileName10 = "holter_presion.png";

            string imagePath11 = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory) + @"img\holter_ritmo.png";
            string imageFileName11 = "holter_ritmo.png";

            string imagePath12 = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory) + @"img\ecocardiograma.png";
            string imageFileName12 = "ecocardiograma.png";

            string imagePath13 = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory) + @"img\electrocardiograma.png";
            string imageFileName13 = "electrocardiograma.png";

            string htmlimagen = $"<img src=\"cid:{imageFileName}\" alt=\"Embedded Image\" style=\"display: block; height: auto; border: 0; width: 100%;\" width=\"180\" height=\"auto\">";
            string htmlimagen2 = $"<img src=\"cid:{imageFileName2}\" alt=\"Embedded Image\" style=\"display: block; height: auto; border: 0; width: 100%;\" width=\"32\" height=\"auto\">";
            string htmlimagen3 = $"<img src=\"cid:{imageFileName3}\" alt=\"Embedded Image\" style=\"display: block; height: auto; border: 0; width: 100%;\" width=\"32\" height=\"auto\" alt=\"Linkedin\" title=\"linkedin\">";
            string htmlimagen4 = $"<img src=\"cid:{imageFileName4}\" alt=\"Embedded Image\" style=\"display: block; height: auto; border: 0; width: 100%;\" width=\"32\" height=\"auto\" alt=\"Instagram\" title=\"instagram\">";
            string htmlimagen5 = $"<img src=\"cid:{imageFileName5}\" alt=\"Embedded Image\" style=\"display: block; height: auto; border: 0; width: 100%;\" width=\"32\" height=\"auto\" alt=\"YouTube\" title=\"YouTube\">";
            string htmlimagen6 = $"<img src=\"cid:{imageFileName6}\" alt=\"Embedded Image\" style=\"display: block; height: auto; border: 0; width: 100%;\" width=\"180\" height=\"auto\">";
            string htmlimagen7 = $"<img src=\"cid:{imageFileName7}\" alt=\"Embedded Image\" style=\"display: block; height: auto; border: 0; width: 100%;\" width=\"180\" height=\"auto\">";
            string htmlimagen8 = $"<img src=\"cid:{imageFileName8}\" alt=\"Embedded Image\" style=\"display: block; height: auto; border: 0; width: 100%;\" width=\"180\" height=\"auto\">";
            string htmlimagen9 = $"<img src=\"cid:{imageFileName9}\" alt=\"Embedded Image\" style=\"display: block; height: auto; border: 0; width: 100%;\" width=\"180\" height=\"auto\">";
            string htmlimagen10 = $"<img src=\"cid:{imageFileName10}\" alt=\"Embedded Image\" style=\"display: block; height: auto; border: 0; width: 100%;\" width=\"180\" height=\"auto\">";
            string htmlimagen11 = $"<img src=\"cid:{imageFileName11}\" alt=\"Embedded Image\" style=\"display: block; height: auto; border: 0; width: 100%;\" width=\"180\" height=\"auto\">";
            string htmlimagen12 = $"<img src=\"cid:{imageFileName12}\" alt=\"Embedded Image\" style=\"display: block; height: auto; border: 0; width: 100%;\" width=\"180\" height=\"auto\">";
            string htmlimagen13 = $"<img src=\"cid:{imageFileName13}\" alt=\"Embedded Image\" style=\"display: block; height: auto; border: 0; width: 100%;\" width=\"180\" height=\"auto\">";



            //string htmlBody = File.ReadAllText(System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory) + @"html\confirmacionImagen.html");
            string htmlBody = File.ReadAllText(System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory) + @"html\confirmacionImagen.html");

            htmlBody = htmlBody.Replace("@cod_cita", codCita);
            htmlBody = htmlBody.Replace("@token", token);

            if (sistema == "73" || sistema == "74" || sistema == "72" || sistema == "75")
            {
                //htmlBody = File.ReadAllText(System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory) + @"html\confirmacionImagenB.html");
                htmlBody = File.ReadAllText(System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory) + @"html\confirmacionImagenB.html");

                htmlBody = htmlBody.Replace("@@imagen6", "");
                htmlBody = htmlBody.Replace("@@imagen7", "");
                htmlBody = htmlBody.Replace("@@imagen8", "");
                htmlBody = htmlBody.Replace("@@imagen9", "");
                htmlBody = htmlBody.Replace("@@imagen10", "");
                htmlBody = htmlBody.Replace("@@imagen11", "");
                htmlBody = htmlBody.Replace("@@imagen12", "");
                htmlBody = htmlBody.Replace("@@imagen13", "");
            }

            htmlBody = htmlBody.Replace("@@imagen1", htmlimagen);
            htmlBody = htmlBody.Replace("@@imagen2", htmlimagen2);
            htmlBody = htmlBody.Replace("@@imagen3", htmlimagen3);
            htmlBody = htmlBody.Replace("@@imagen4", htmlimagen4);
            htmlBody = htmlBody.Replace("@@imagen5", htmlimagen5);
            htmlBody = htmlBody.Replace("@@imagen6", "");
            htmlBody = htmlBody.Replace("@@imagen7", "");
            htmlBody = htmlBody.Replace("@@imagen8", "");
            htmlBody = htmlBody.Replace("@@imagen9", "");
            htmlBody = htmlBody.Replace("@@imagen10", "");
            htmlBody = htmlBody.Replace("@@imagen11", "");
            htmlBody = htmlBody.Replace("@@imagen12", "");
            htmlBody = htmlBody.Replace("@@imagen13", "");

            if (sistema == "62" || sistema == "1")
            {
                htmlBody = htmlBody.Replace("@@textobjage", "Toma de Muestra");
                htmlBody = htmlBody.Replace("@@telefono", "<p style=\"margin: 0;\">Si desea realizar cambios a su hora puede contactarnos al<span style=\"color: #dd3488;\"><a href=\"tel:22 994 4003\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a></span></p>");
                htmlBody = htmlBody.Replace("@@imagen6", "");
                htmlBody = htmlBody.Replace("@@imagen7", "");
                htmlBody = htmlBody.Replace("@@imagen8", "");
                htmlBody = htmlBody.Replace("@@imagen9", "");
                htmlBody = htmlBody.Replace("@@imagen10", "");
                htmlBody = htmlBody.Replace("@@imagen11", "");
                htmlBody = htmlBody.Replace("@@imagen12", "");
                htmlBody = htmlBody.Replace("@@imagen13", "");
            }
            if (sistema == "73")
            {
                htmlBody = htmlBody.Replace("@@textobjage", "Cirugía Refractiva");
                htmlBody = htmlBody.Replace("@@telefono", "<p style=\"margin: 0;\">Si desea realizar cambios a su hora puede contactarnos al<span style=\"color: #dd3488;\"><a href=\"tel:22 994 4003\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a></span></p>");
                htmlBody = htmlBody.Replace("@@imagen6", "");
                htmlBody = htmlBody.Replace("@@imagen7", "");
                htmlBody = htmlBody.Replace("@@imagen8", "");
                htmlBody = htmlBody.Replace("@@imagen9", "");
                htmlBody = htmlBody.Replace("@@imagen10", "");
                htmlBody = htmlBody.Replace("@@imagen11", "");
                htmlBody = htmlBody.Replace("@@imagen12", "");
                htmlBody = htmlBody.Replace("@@imagen13", "");
            }

            if (sistema == "74")
            {
                htmlBody = htmlBody.Replace("@@textobjage", "Evaluación Gratuita Cirugía Lasik");
                htmlBody = htmlBody.Replace("@@telefono", "<p style=\"margin: 0;\">Si desea realizar cambios a su hora puede contactarnos al<span style=\"color: #dd3488;\"><a href=\"tel:22 994 4003\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a></span></p>");
                htmlBody = htmlBody.Replace("@@imagen6", "");
                htmlBody = htmlBody.Replace("@@imagen7", "");
                htmlBody = htmlBody.Replace("@@imagen8", "");
                htmlBody = htmlBody.Replace("@@imagen9", "");
                htmlBody = htmlBody.Replace("@@imagen10", "");
                htmlBody = htmlBody.Replace("@@imagen11", "");
                htmlBody = htmlBody.Replace("@@imagen12", "");
                htmlBody = htmlBody.Replace("@@imagen13", "");
            }
            if (sistema == "72")
            {
                htmlBody = htmlBody.Replace("@@textobjage", "Control Cirugía Lasik");
                htmlBody = htmlBody.Replace("@@telefono", "<p style=\"margin: 0;\">Si desea realizar cambios a su hora puede contactarnos al<span style=\"color: #dd3488;\"><a href=\"tel:22 994 4003\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a></span></p>");
                htmlBody = htmlBody.Replace("@@imagen6", "");
                htmlBody = htmlBody.Replace("@@imagen7", "");
                htmlBody = htmlBody.Replace("@@imagen8", "");
                htmlBody = htmlBody.Replace("@@imagen9", "");
                htmlBody = htmlBody.Replace("@@imagen10", "");
                htmlBody = htmlBody.Replace("@@imagen11", "");
                htmlBody = htmlBody.Replace("@@imagen12", "");
                htmlBody = htmlBody.Replace("@@imagen13", "");
            }
            if (sistema == "75")
            {
                htmlBody = htmlBody.Replace("@@textobjage", "Operativo Preventivo");
                htmlBody = htmlBody.Replace("@@textoIsapre", ""); /*Si eres paciente Isapre, 'Bono Clini' "*/
                htmlBody = htmlBody.Replace("@@telefono", "");
                htmlBody = htmlBody.Replace("@@textoOperativo", "Para participar debes llegar al operativo con 9 a 16 horas de ayuno.");
                htmlBody = htmlBody.Replace("@@imagen6", "");
                htmlBody = htmlBody.Replace("@@imagen7", "");
                htmlBody = htmlBody.Replace("@@imagen8", "");
                htmlBody = htmlBody.Replace("@@imagen9", "");
                htmlBody = htmlBody.Replace("@@imagen10", "");
                htmlBody = htmlBody.Replace("@@imagen11", "");
                htmlBody = htmlBody.Replace("@@imagen12", "");
                htmlBody = htmlBody.Replace("@@imagen13", "");
            }

            //if (sistema == "64")
            //{
            //    htmlBody = htmlBody.Replace("@@textobjage", "Resonancia");
            //    htmlBody = htmlBody.Replace("@@telefono", "<p style=\"margin: 0;\">Si desea realizar cambios a su hora puede contactarnos al<span style=\"color: #dd3488;\"><a href=\"tel:22 994 4003\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a></span></p>");

            //}
            //if (sistema == "65")
            //{
            //    htmlBody = htmlBody.Replace("@@textobjage", "Mamografía");
            //    htmlBody = htmlBody.Replace("@@telefono", "<p style=\"margin: 0;\">Si desea realizar cambios a su hora puede contactarnos al<span style=\"color: #dd3488;\"><a href=\"tel:22 994 4003\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a></span></p>");
            //}

            if (sistema == "66")
            {
                htmlBody = htmlBody.Replace("@@textobjage", "Ecotomografía");
                htmlBody = htmlBody.Replace("@@telefono", "<p style=\"margin: 0;\">Si desea realizar cambios a su hora puede contactarnos al<span style=\"color: #dd3488;\"><a href=\"tel:22 994 4003\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a></span></p>");
            }
            //if (sistema == "70")
            //{
            //    htmlBody = htmlBody.Replace("@@textobjage", "Scanner");
            //    htmlBody = htmlBody.Replace("@@telefono", "<p style=\"margin: 0;\">Si desea realizar cambios a su hora puede contactarnos al<span style=\"color: #dd3488;\"><a href=\"tel:22 994 4003\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a></span></p>");
            //}
            if (sistema == "71")
            {
                htmlBody = htmlBody.Replace("@@textobjage", "Pet CT");
                htmlBody = htmlBody.Replace("@@telefono", "<p style=\"margin: 0;\">\r\n  Si desea realizar cambios a su hora puede contactarnos al\r\n  <span style=\"color: #dd3488;\">\r\n    <a href=\"tel:22 994 4003\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a>\r\n  </span>\r\n</p>\r\n<p style=\"margin: 0;\">\r\n  Además, puede revisar las <a href=\"https://www.clini.cl/indicaciones-pet\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">indicaciones previas a su examen aquí</a>.\r\n</p>");
                mail.Bcc.Add("macarena.fierro@clini.cl");
                mail.Bcc.Add("sandra.gutierrez@clini.cl");
            }

            if (sistema == "82" && subtipoobj == "42")
            {
                htmlBody = htmlBody.Replace("@@textobjage", "Ecogine Transvaginal");
                htmlBody = htmlBody.Replace("@@telefono", "<p style=\"margin: 0;\">Si desea realizar cambios a su hora puede contactarnos al<span style=\"color: #dd3488;\"><a href=\"tel:22 994 4003\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a></span></p>");
            }

            if (sistema == "82" && subtipoobj == "43")
            {
                htmlBody = htmlBody.Replace("@@textobjage", "Ecografía 4D y 5D");
                htmlBody = htmlBody.Replace("@@telefono", "<p style=\"margin: 0;\">Si desea realizar cambios a su hora puede contactarnos al<span style=\"color: #dd3488;\"><a href=\"tel:22 994 4003\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a></span></p>");
            }

            if (sistema == "82" && subtipoobj == "44")
            {
                htmlBody = htmlBody.Replace("@@textobjage", "Ecografía doppler 11 a 14 semanas");
                htmlBody = htmlBody.Replace("@@telefono", "<p style=\"margin: 0;\">Si desea realizar cambios a su hora puede contactarnos al<span style=\"color: #dd3488;\"><a href=\"tel:22 994 4003\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a></span></p>");
            }

            if (sistema == "82" && subtipoobj == "45")
            {
                htmlBody = htmlBody.Replace("@@textobjage", "Ecografía doppler 20 a 24 semanas");
                htmlBody = htmlBody.Replace("@@telefono", "<p style=\"margin: 0;\">Si desea realizar cambios a su hora puede contactarnos al<span style=\"color: #dd3488;\"><a href=\"tel:22 994 4003\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a></span></p>");
            }

            if (sistema == "76")
            {
                htmlBody = htmlBody.Replace("@@textobjage", "Holter de Ritmo");
                htmlBody = htmlBody.Replace("@@telefono", "<p style=\"margin: 0;\">Si desea realizar cambios a su hora puede contactarnos al<span style=\"color: #dd3488;\"><a href=\"tel:22 994 4003\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a></span></p>");
            }

            if (sistema == "77")
            {
                htmlBody = htmlBody.Replace("@@textobjage", "Holter de Presión");
                htmlBody = htmlBody.Replace("@@telefono", "<p style=\"margin: 0;\">Si desea realizar cambios a su hora puede contactarnos al<span style=\"color: #dd3488;\"><a href=\"tel:22 994 4003\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a></span></p>");
            }

            if (sistema == "78")
            {
                htmlBody = htmlBody.Replace("@@textobjage", "Electrocardiograma");
                htmlBody = htmlBody.Replace("@@telefono", "<p style=\"margin: 0;\">Si desea realizar cambios a su hora puede contactarnos al<span style=\"color: #dd3488;\"><a href=\"tel:22 994 4003\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a></span></p>");
            }

            if (sistema == "79")
            {
                htmlBody = htmlBody.Replace("@@textobjage", "Ecocardiograma");
                htmlBody = htmlBody.Replace("@@telefono", "<p style=\"margin: 0;\">Si desea realizar cambios a su hora puede contactarnos al<span style=\"color: #dd3488;\"><a href=\"tel:22 994 4003\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a></span></p>");
            }

            else
            {
                htmlBody = htmlBody.Replace("@@textoIsapre", ""); // Remover el marcador si no es sistema 75 
                htmlBody = htmlBody.Replace("@@textoOperativo", "");
            }

            if (sistema == "81" && subtipoobj == "60")
            {
                htmlBody = htmlBody.Replace("@@textobjage", "Ecografía Mamaria");
                htmlBody = htmlBody.Replace("@@telefono", "<p style=\"margin: 0;\">\r\n  Si desea realizar cambios a su hora puede contactarnos al\r\n  <span style=\"color: #dd3488;\">\r\n    <a href=\"tel:22 994 4003\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a>\r\n  </span>\r\n</p>\r\n<p style=\"margin: 0;\">\r\n  Además, puede revisar las <a href=\"https://www.clini.cl/indicaciones\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">indicaciones previas a su examen aquí</a>.\r\n</p>");
            }

            if (sistema == "81" && subtipoobj == "61")
            {
                htmlBody = htmlBody.Replace("@@textobjage", "Ecografía Abdominal, Tiroides, Renal");
                htmlBody = htmlBody.Replace("@@telefono", "<p style=\"margin: 0;\">\r\n  Si desea realizar cambios a su hora puede contactarnos al\r\n  <span style=\"color: #dd3488;\">\r\n    <a href=\"tel:22 994 4003\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a>\r\n  </span>\r\n</p>\r\n<p style=\"margin: 0;\">\r\n  Además, puede revisar las <a href=\"https://www.clini.cl/indicaciones\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">indicaciones previas a su examen aquí</a>.\r\n</p>");
            }

            if (sistema == "81" && subtipoobj == "62")
            {
                htmlBody = htmlBody.Replace("@@textobjage", "Ecografía Pelvis");
                htmlBody = htmlBody.Replace("@@telefono", "<p style=\"margin: 0;\">\r\n  Si desea realizar cambios a su hora puede contactarnos al\r\n  <span style=\"color: #dd3488;\">\r\n    <a href=\"tel:22 994 4003\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a>\r\n  </span>\r\n</p>\r\n<p style=\"margin: 0;\">\r\n  Además, puede revisar las <a href=\"https://www.clini.cl/indicaciones\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">indicaciones previas a su examen aquí</a>.\r\n</p>");
            }

            if (sistema == "81" && subtipoobj == "63")
            {
                htmlBody = htmlBody.Replace("@@textobjage", "Ecografía Doppler Renal, Abdominal, Hepatica, Aortoiliaco");
                htmlBody = htmlBody.Replace("@@telefono", "<p style=\"margin: 0;\">\r\n  Si desea realizar cambios a su hora puede contactarnos al\r\n  <span style=\"color: #dd3488;\">\r\n    <a href=\"tel:22 994 4003\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a>\r\n  </span>\r\n</p>\r\n<p style=\"margin: 0;\">\r\n  Además, puede revisar las <a href=\"https://www.clini.cl/indicaciones\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">indicaciones previas a su examen aquí</a>.\r\n</p>");
            }

            if (sistema == "81" && subtipoobj == "64")
            {
                htmlBody = htmlBody.Replace("@@textobjage", "Ecografía Musculoesqueletica u Osteoarticular");
                htmlBody = htmlBody.Replace("@@telefono", "<p style=\"margin: 0;\">\r\n  Si desea realizar cambios a su hora puede contactarnos al\r\n  <span style=\"color: #dd3488;\">\r\n    <a href=\"tel:22 994 4003\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a>\r\n  </span>\r\n</p>\r\n<p style=\"margin: 0;\">\r\n  Además, puede revisar las <a href=\"https://www.clini.cl/indicaciones\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">indicaciones previas a su examen aquí</a>.\r\n</p>");
            }

            if (sistema == "81" && subtipoobj == "65")
            {
                htmlBody = htmlBody.Replace("@@textobjage", "Otras Ecografías");
                htmlBody = htmlBody.Replace("@@telefono", "<p style=\"margin: 0;\">\r\n  Si desea realizar cambios a su hora puede contactarnos al\r\n  <span style=\"color: #dd3488;\">\r\n    <a href=\"tel:22 994 4003\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a>\r\n  </span>\r\n</p>\r\n<p style=\"margin: 0;\">\r\n  Además, puede revisar las <a href=\"https://www.clini.cl/indicaciones\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">indicaciones previas a su examen aquí</a>.\r\n</p>");
            }

            if (sistema == "81" && subtipoobj == "66")
            {
                htmlBody = htmlBody.Replace("@@textobjage", "Ecografía Tiroides y Renal");
                htmlBody = htmlBody.Replace("@@telefono", "<p style=\"margin: 0;\">\r\n  Si desea realizar cambios a su hora puede contactarnos al\r\n  <span style=\"color: #dd3488;\">\r\n    <a href=\"tel:22 994 4003\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a>\r\n  </span>\r\n</p>\r\n<p style=\"margin: 0;\">\r\n  Además, puede revisar las <a href=\"https://www.clini.cl/indicaciones\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">indicaciones previas a su examen aquí</a>.\r\n</p>");
            }

            if (sistema == "81" && subtipoobj == "67")
            {
                htmlBody = htmlBody.Replace("@@textobjage", "Ecografía Abdominal");
                htmlBody = htmlBody.Replace("@@telefono", "<p style=\"margin: 0;\">\r\n  Si desea realizar cambios a su hora puede contactarnos al\r\n  <span style=\"color: #dd3488;\">\r\n    <a href=\"tel:22 994 4003\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a>\r\n  </span>\r\n</p>\r\n<p style=\"margin: 0;\">\r\n  Además, puede revisar las <a href=\"https://www.clini.cl/indicaciones\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">indicaciones previas a su examen aquí</a>.\r\n</p>");
            }

            if (sistema == "46" && subtipoobj == "101603")
            {
                htmlBody = htmlBody.Replace("@@textobjage", "Consulta Traumatología");
                htmlBody = htmlBody.Replace("@@telefono", "<p style=\"margin: 0;\">Si desea realizar cambios a su hora puede contactarnos al<span style=\"color: #dd3488;\"><a href=\"tel:22 994 4003\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a></span></p>");
            }

            if (sistema == "46" && subtipoobj == "101614")
            {
                htmlBody = htmlBody.Replace("@@textobjage", "Consulta Ginecología");
                htmlBody = htmlBody.Replace("@@telefono", "<p style=\"margin: 0;\">Si desea realizar cambios a su hora puede contactarnos al<span style=\"color: #dd3488;\"><a href=\"tel:22 994 4003\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a></span></p>");
            }

            if (sistema == "46" && subtipoobj == "101604")
            {
                htmlBody = htmlBody.Replace("@@textobjage", "Consulta Urología");
                htmlBody = htmlBody.Replace("@@telefono", "<p style=\"margin: 0;\">Si desea realizar cambios a su hora puede contactarnos al<span style=\"color: #dd3488;\"><a href=\"tel:22 994 4003\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a></span></p>");
            }

            if (sistema == "64" && subtipoobj == "800")
            {
                htmlBody = htmlBody.Replace("@@textobjage", "RM Cerebro / Cabeza / Cuello");

                if (sucursal.Trim().ToUpper() == "CENTRO DE SALUD | METRO TOBALABA")
                {
                    htmlBody = htmlBody.Replace("@@telefono",
                        "<p style=\"margin: 0;\">Si desea realizar cambios a su hora puede contactarnos al " +
                        "<span style=\"color: #dd3488;\"><a href=\"tel:22 994 4003\" target=\"_blank\" " +
                        "style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a></span></p>" +
                        "<p style=\"margin: 15px auto 15px auto; font-weight: bold;\">Evita filas y tiempos de espera. Completa desde casa el formulario obligatorio antes de tu examen.</p>" +
                        $"<p style=\"margin: 0;\"><a href=\"https://formularios.clini.cl/resonancia/{codCita}\" " +
                        "target=\"_blank\" style=\"background-color: #dd3488; color: white; padding: 10px 20px; " +
                        "text-decoration: none; border-radius: 200px; font-weight: bold; display: inline-block;\">¡Llena tu formulario ahora!</a></p>");
                }
                else
                {
                    htmlBody = htmlBody.Replace("@@telefono",
                        "<p style=\"margin: 0;\">Si desea realizar cambios a su hora puede contactarnos al " +
                        "<span style=\"color: #dd3488;\"><a href=\"tel:22 994 4003\" target=\"_blank\" " +
                        "style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a></span></p>" +
                        "<p style=\"margin: 10px 0 0 0;\">Encuesta Resonancia : " +
                        "<a href=\"https://www.clini.cl/admi/formulario-resonancias\" target=\"_blank\" style=\"color: #3366cc; text-decoration: underline;\">https://www.clini.cl/admi/formulario-resonancias</a></p>");
                }
            }


            if (sistema == "64" && subtipoobj == "801")
            {
                htmlBody = htmlBody.Replace("@@textobjage", "RM CUELLO");
                if (sucursal.Trim().ToUpper() == "CENTRO DE SALUD | METRO TOBALABA")
                {

                    htmlBody = htmlBody.Replace("@@telefono",
                        "<p style=\"margin: 0;\">Si desea realizar cambios a su hora puede contactarnos al " +
                        "<span style=\"color: #dd3488;\"><a href=\"tel:22 994 4003\" target=\"_blank\" " +
                        "style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a></span></p>" +
                        "<p style=\"margin: 15px auto 15px auto; font-weight: bold;\">Evita filas y tiempos de espera. Completa desde casa el formulario obligatorio antes de tu examen.</p>" +
                        $"<p style=\"margin: 0;\"><a href=\"https://formularios.clini.cl/resonancia/{codCita}\" " +
                        "target=\"_blank\" style=\"background-color: #dd3488; color: white; padding: 10px 20px; " +
                        "text-decoration: none; border-radius: 200px; font-weight: bold; display: inline-block;\">¡Llena tu formulario ahora!</a></p>");
                }
                else
                {
                    htmlBody = htmlBody.Replace("@@telefono", "<p style=\"margin: 0;\">\r\n  Si desea realizar cambios a su hora puede contactarnos al\r\n  <span style=\"color: #dd3488;\">\r\n    <a href=\"tel:22 994 4003\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a>\r\n  </span>\r\n</p>\r\n<p style=\"margin: 0;\">\r\n  Además, puede revisar las <a href=\"https://www.clini.cl/indicaciones\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">indicaciones previas a su examen aquí</a>.\r\n</p>");
                }


            }

            if (sistema == "64" && subtipoobj == "808")
            {
                htmlBody = htmlBody.Replace("@@textobjage", "RM COLUMNA TOTAL CON CONTRASTE");
                if (sucursal.Trim().ToUpper() == "CENTRO DE SALUD | METRO TOBALABA")
                {

                    htmlBody = htmlBody.Replace("@@telefono",
                        "<p style=\"margin: 0;\">Si desea realizar cambios a su hora puede contactarnos al " +
                        "<span style=\"color: #dd3488;\"><a href=\"tel:22 994 4003\" target=\"_blank\" " +
                        "style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a></span></p>" +
                        "<p style=\"margin: 15px auto 15px auto; font-weight: bold;\">Evita filas y tiempos de espera. Completa desde casa el formulario obligatorio antes de tu examen.</p>" +
                        $"<p style=\"margin: 0;\"><a href=\"https://formularios.clini.cl/resonancia/{codCita}\" " +
                        "target=\"_blank\" style=\"background-color: #dd3488; color: white; padding: 10px 20px; " +
                        "text-decoration: none; border-radius: 200px; font-weight: bold; display: inline-block;\">¡Llena tu formulario ahora!</a></p>");
                }
                else
                {
                    htmlBody = htmlBody.Replace("@@telefono", "<p style=\"margin: 0;\">\r\n  Si desea realizar cambios a su hora puede contactarnos al\r\n  <span style=\"color: #dd3488;\">\r\n    <a href=\"tel:22 994 4003\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a>\r\n  </span>\r\n</p>\r\n<p style=\"margin: 0;\">\r\n  Además, puede revisar las <a href=\"https://www.clini.cl/indicaciones\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">indicaciones previas a su examen aquí</a>.\r\n</p>");
                }


            }

            if (sistema == "64" && subtipoobj == "802")
            {
                htmlBody = htmlBody.Replace("@@textobjage", "RM ABDOMEN");
                if (sucursal.Trim().ToUpper() == "CENTRO DE SALUD | METRO TOBALABA")
                {

                    htmlBody = htmlBody.Replace("@@telefono",
                        "<p style=\"margin: 0;\">Si desea realizar cambios a su hora puede contactarnos al " +
                        "<span style=\"color: #dd3488;\"><a href=\"tel:22 994 4003\" target=\"_blank\" " +
                        "style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a></span></p>" +
                        "<p style=\"margin: 15px auto 15px auto; font-weight: bold;\">Evita filas y tiempos de espera. Completa desde casa el formulario obligatorio antes de tu examen.</p>" +
                        $"<p style=\"margin: 0;\"><a href=\"https://formularios.clini.cl/resonancia/{codCita}\" " +
                        "target=\"_blank\" style=\"background-color: #dd3488; color: white; padding: 10px 20px; " +
                        "text-decoration: none; border-radius: 200px; font-weight: bold; display: inline-block;\">¡Llena tu formulario ahora!</a></p>");
                }
                else
                {
                    htmlBody = htmlBody.Replace("@@telefono", "<p style=\"margin: 0;\">\r\n  Si desea realizar cambios a su hora puede contactarnos al\r\n  <span style=\"color: #dd3488;\">\r\n    <a href=\"tel:22 994 4003\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a>\r\n  </span>\r\n</p>\r\n<p style=\"margin: 0;\">\r\n  Además, puede revisar las <a href=\"https://www.clini.cl/indicaciones\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">indicaciones previas a su examen aquí</a>.\r\n</p>");
                }


            }

            if (sistema == "64" && subtipoobj == "803")
            {
                htmlBody = htmlBody.Replace("@@textobjage", "RM PELVIS");
                if (sucursal.Trim().ToUpper() == "CENTRO DE SALUD | METRO TOBALABA")
                {

                    htmlBody = htmlBody.Replace("@@telefono",
                        "<p style=\"margin: 0;\">Si desea realizar cambios a su hora puede contactarnos al " +
                        "<span style=\"color: #dd3488;\"><a href=\"tel:22 994 4003\" target=\"_blank\" " +
                        "style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a></span></p>" +
                        "<p style=\"margin: 15px auto 15px auto; font-weight: bold;\">Evita filas y tiempos de espera. Completa desde casa el formulario obligatorio antes de tu examen.</p>" +
                        $"<p style=\"margin: 0;\"><a href=\"https://formularios.clini.cl/resonancia/{codCita}\" " +
                        "target=\"_blank\" style=\"background-color: #dd3488; color: white; padding: 10px 20px; " +
                        "text-decoration: none; border-radius: 200px; font-weight: bold; display: inline-block;\">¡Llena tu formulario ahora!</a></p>");
                }
                else
                {
                    htmlBody = htmlBody.Replace("@@telefono", "<p style=\"margin: 0;\">\r\n  Si desea realizar cambios a su hora puede contactarnos al\r\n  <span style=\"color: #dd3488;\">\r\n    <a href=\"tel:22 994 4003\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a>\r\n  </span>\r\n</p>\r\n<p style=\"margin: 0;\">\r\n  Además, puede revisar las <a href=\"https://www.clini.cl/indicaciones\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">indicaciones previas a su examen aquí</a>.\r\n</p>");
                }


            }

            if (sistema == "64" && subtipoobj == "804")
            {
                htmlBody = htmlBody.Replace("@@textobjage", "RM RODILLA, CADERA, HOMBRO, MANO, MUÑECA");
                if (sucursal.Trim().ToUpper() == "CENTRO DE SALUD | METRO TOBALABA")
                {

                    htmlBody = htmlBody.Replace("@@telefono",
                        "<p style=\"margin: 0;\">Si desea realizar cambios a su hora puede contactarnos al " +
                        "<span style=\"color: #dd3488;\"><a href=\"tel:22 994 4003\" target=\"_blank\" " +
                        "style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a></span></p>" +
                        "<p style=\"margin: 15px auto 15px auto; font-weight: bold;\">Evita filas y tiempos de espera. Completa desde casa el formulario obligatorio antes de tu examen.</p>" +
                        $"<p style=\"margin: 0;\"><a href=\"https://formularios.clini.cl/resonancia/{codCita}\" " +
                        "target=\"_blank\" style=\"background-color: #dd3488; color: white; padding: 10px 20px; " +
                        "text-decoration: none; border-radius: 200px; font-weight: bold; display: inline-block;\">¡Llena tu formulario ahora!</a></p>");
                }
                else
                {
                    htmlBody = htmlBody.Replace("@@telefono", "<p style=\"margin: 0;\">\r\n  Si desea realizar cambios a su hora puede contactarnos al\r\n  <span style=\"color: #dd3488;\">\r\n    <a href=\"tel:22 994 4003\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a>\r\n  </span>\r\n</p>\r\n<p style=\"margin: 0;\">\r\n  Además, puede revisar las <a href=\"https://www.clini.cl/indicaciones\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">indicaciones previas a su examen aquí</a>.\r\n</p>");
                }


            }

            if (sistema == "64" && subtipoobj == "806")
            {
                htmlBody = htmlBody.Replace("@@textobjage", "RM AMBAS RODILLAS, CADERAS, HOMBROS, MANOS, MUÑECAS");
                if (sucursal.Trim().ToUpper() == "CENTRO DE SALUD | METRO TOBALABA")
                {

                    htmlBody = htmlBody.Replace("@@telefono",
                        "<p style=\"margin: 0;\">Si desea realizar cambios a su hora puede contactarnos al " +
                        "<span style=\"color: #dd3488;\"><a href=\"tel:22 994 4003\" target=\"_blank\" " +
                        "style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a></span></p>" +
                        "<p style=\"margin: 15px auto 15px auto; font-weight: bold;\">Evita filas y tiempos de espera. Completa desde casa el formulario obligatorio antes de tu examen.</p>" +
                        $"<p style=\"margin: 0;\"><a href=\"https://formularios.clini.cl/resonancia/{codCita}\" " +
                        "target=\"_blank\" style=\"background-color: #dd3488; color: white; padding: 10px 20px; " +
                        "text-decoration: none; border-radius: 200px; font-weight: bold; display: inline-block;\">¡Llena tu formulario ahora!</a></p>");
                }
                else
                {
                    htmlBody = htmlBody.Replace("@@telefono", "<p style=\"margin: 0;\">\r\n  Si desea realizar cambios a su hora puede contactarnos al\r\n  <span style=\"color: #dd3488;\">\r\n    <a href=\"tel:22 994 4003\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a>\r\n  </span>\r\n</p>\r\n<p style=\"margin: 0;\">\r\n  Además, puede revisar las <a href=\"https://www.clini.cl/indicaciones\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">indicaciones previas a su examen aquí</a>.\r\n</p>");
                }


            }

            if (sistema == "64" && subtipoobj == "805")
            {
                htmlBody = htmlBody.Replace("@@textobjage", "RM PIERNA, BRAZO, ANTEBRAZO");
                if (sucursal.Trim().ToUpper() == "CENTRO DE SALUD | METRO TOBALABA")
                {

                    htmlBody = htmlBody.Replace("@@telefono",
                        "<p style=\"margin: 0;\">Si desea realizar cambios a su hora puede contactarnos al " +
                        "<span style=\"color: #dd3488;\"><a href=\"tel:22 994 4003\" target=\"_blank\" " +
                        "style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a></span></p>" +
                        "<p style=\"margin: 15px auto 15px auto; font-weight: bold;\">Evita filas y tiempos de espera. Completa desde casa el formulario obligatorio antes de tu examen.</p>" +
                        $"<p style=\"margin: 0;\"><a href=\"https://formularios.clini.cl/resonancia/{codCita}\" " +
                        "target=\"_blank\" style=\"background-color: #dd3488; color: white; padding: 10px 20px; " +
                        "text-decoration: none; border-radius: 200px; font-weight: bold; display: inline-block;\">¡Llena tu formulario ahora!</a></p>");
                }
                else
                {
                    htmlBody = htmlBody.Replace("@@telefono", "<p style=\"margin: 0;\">\r\n  Si desea realizar cambios a su hora puede contactarnos al\r\n  <span style=\"color: #dd3488;\">\r\n    <a href=\"tel:22 994 4003\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a>\r\n  </span>\r\n</p>\r\n<p style=\"margin: 0;\">\r\n  Además, puede revisar las <a href=\"https://www.clini.cl/indicaciones\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">indicaciones previas a su examen aquí</a>.\r\n</p>");
                }


            }

            if (sistema == "64" && subtipoobj == "807")
            {
                htmlBody = htmlBody.Replace("@@textobjage", "RM AMBAS PIERNAS, BRAZOS, ANTEBRAZOS");
                if (sucursal.Trim().ToUpper() == "CENTRO DE SALUD | METRO TOBALABA")
                {

                    htmlBody = htmlBody.Replace("@@telefono",
                        "<p style=\"margin: 0;\">Si desea realizar cambios a su hora puede contactarnos al " +
                        "<span style=\"color: #dd3488;\"><a href=\"tel:22 994 4003\" target=\"_blank\" " +
                        "style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a></span></p>" +
                        "<p style=\"margin: 15px auto 15px auto; font-weight: bold;\">Evita filas y tiempos de espera. Completa desde casa el formulario obligatorio antes de tu examen.</p>" +
                        $"<p style=\"margin: 0;\"><a href=\"https://formularios.clini.cl/resonancia/{codCita}\" " +
                        "target=\"_blank\" style=\"background-color: #dd3488; color: white; padding: 10px 20px; " +
                        "text-decoration: none; border-radius: 200px; font-weight: bold; display: inline-block;\">¡Llena tu formulario ahora!</a></p>");
                }
                else
                {
                    htmlBody = htmlBody.Replace("@@telefono", "<p style=\"margin: 0;\">\r\n  Si desea realizar cambios a su hora puede contactarnos al\r\n  <span style=\"color: #dd3488;\">\r\n    <a href=\"tel:22 994 4003\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a>\r\n  </span>\r\n</p>\r\n<p style=\"margin: 0;\">\r\n  Además, puede revisar las <a href=\"https://www.clini.cl/indicaciones\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">indicaciones previas a su examen aquí</a>.\r\n</p>");
                }


            }

            if (sistema == "64" && subtipoobj == "81")
            {
                htmlBody = htmlBody.Replace("@@textobjage", "RM COLUMNA CERVICAL O DORSAL O LUMBAR");
                if (sucursal.Trim().ToUpper() == "CENTRO DE SALUD | METRO TOBALABA")
                {

                    htmlBody = htmlBody.Replace("@@telefono",
                        "<p style=\"margin: 0;\">Si desea realizar cambios a su hora puede contactarnos al " +
                        "<span style=\"color: #dd3488;\"><a href=\"tel:22 994 4003\" target=\"_blank\" " +
                        "style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a></span></p>" +
                        "<p style=\"margin: 15px auto 15px auto; font-weight: bold;\">Evita filas y tiempos de espera. Completa desde casa el formulario obligatorio antes de tu examen.</p>" +
                        $"<p style=\"margin: 0;\"><a href=\"https://formularios.clini.cl/resonancia/{codCita}\" " +
                        "target=\"_blank\" style=\"background-color: #dd3488; color: white; padding: 10px 20px; " +
                        "text-decoration: none; border-radius: 200px; font-weight: bold; display: inline-block;\">¡Llena tu formulario ahora!</a></p>");
                }
                else
                {
                    htmlBody = htmlBody.Replace("@@telefono", "<p style=\"margin: 0;\">\r\n  Si desea realizar cambios a su hora puede contactarnos al\r\n  <span style=\"color: #dd3488;\">\r\n    <a href=\"tel:22 994 4003\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a>\r\n  </span>\r\n</p>\r\n<p style=\"margin: 0;\">\r\n  Además, puede revisar las <a href=\"https://www.clini.cl/indicaciones\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">indicaciones previas a su examen aquí</a>.\r\n</p>");
                }


            }

            if (sistema == "64" && subtipoobj == "82")
            {
                htmlBody = htmlBody.Replace("@@textobjage", "RM COLUMNA TOTAL");
                if (sucursal.Trim().ToUpper() == "CENTRO DE SALUD | METRO TOBALABA")
                {

                    htmlBody = htmlBody.Replace("@@telefono",
                        "<p style=\"margin: 0;\">Si desea realizar cambios a su hora puede contactarnos al " +
                        "<span style=\"color: #dd3488;\"><a href=\"tel:22 994 4003\" target=\"_blank\" " +
                        "style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a></span></p>" +
                        "<p style=\"margin: 15px auto 15px auto; font-weight: bold;\">Evita filas y tiempos de espera. Completa desde casa el formulario obligatorio antes de tu examen.</p>" +
                        $"<p style=\"margin: 0;\"><a href=\"https://formularios.clini.cl/resonancia/{codCita}\" " +
                        "target=\"_blank\" style=\"background-color: #dd3488; color: white; padding: 10px 20px; " +
                        "text-decoration: none; border-radius: 200px; font-weight: bold; display: inline-block;\">¡Llena tu formulario ahora!</a></p>");
                }
                else
                {
                    htmlBody = htmlBody.Replace("@@telefono", "<p style=\"margin: 0;\">\r\n  Si desea realizar cambios a su hora puede contactarnos al\r\n  <span style=\"color: #dd3488;\">\r\n    <a href=\"tel:22 994 4003\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a>\r\n  </span>\r\n</p>\r\n<p style=\"margin: 0;\">\r\n  Además, puede revisar las <a href=\"https://www.clini.cl/indicaciones\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">indicaciones previas a su examen aquí</a>.\r\n</p>");
                }


            }



            if (sistema == "64" && subtipoobj == "83")
            {
                htmlBody = htmlBody.Replace("@@textobjage", "RM ABDOMEN Y PELVIS");
                if (sucursal.Trim().ToUpper() == "CENTRO DE SALUD | METRO TOBALABA")
                {

                    htmlBody = htmlBody.Replace("@@telefono",
                        "<p style=\"margin: 0;\">Si desea realizar cambios a su hora puede contactarnos al " +
                        "<span style=\"color: #dd3488;\"><a href=\"tel:22 994 4003\" target=\"_blank\" " +
                        "style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a></span></p>" +
                        "<p style=\"margin: 15px auto 15px auto; font-weight: bold;\">Evita filas y tiempos de espera. Completa desde casa el formulario obligatorio antes de tu examen.</p>" +
                        $"<p style=\"margin: 0;\"><a href=\"https://formularios.clini.cl/resonancia/{codCita}\" " +
                        "target=\"_blank\" style=\"background-color: #dd3488; color: white; padding: 10px 20px; " +
                        "text-decoration: none; border-radius: 200px; font-weight: bold; display: inline-block;\">¡Llena tu formulario ahora!</a></p>");
                }
                else
                {
                    htmlBody = htmlBody.Replace("@@telefono", "<p style=\"margin: 0;\">\r\n  Si desea realizar cambios a su hora puede contactarnos al\r\n  <span style=\"color: #dd3488;\">\r\n    <a href=\"tel:22 994 4003\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a>\r\n  </span>\r\n</p>\r\n<p style=\"margin: 0;\">\r\n  Además, puede revisar las <a href=\"https://www.clini.cl/indicaciones\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">indicaciones previas a su examen aquí</a>.\r\n</p>");
                }


            }

            if (sistema == "70" && subtipoobj == "90")
            {
                htmlBody = htmlBody.Replace("@@textobjage", "TAC Cerebro/Cabeza/Cuello");
                htmlBody = htmlBody.Replace("@@telefono", "<p style=\"margin: 0;\">Si desea realizar cambios a su hora puede contactarnos al<span style=\"color: #dd3488;\"><a href=\"tel:22 994 4003\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a></span></p>");

            }

            if (sistema == "70" && subtipoobj == "91")
            {
                htmlBody = htmlBody.Replace("@@textobjage", "TAC Columna cervical, dorsal o lumbar");
                htmlBody = htmlBody.Replace("@@telefono", "<p style=\"margin: 0;\">Si desea realizar cambios a su hora puede contactarnos al<span style=\"color: #dd3488;\"><a href=\"tel:22 994 4003\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a></span></p>");

            }

            if (sistema == "70" && subtipoobj == "92")
            {
                htmlBody = htmlBody.Replace("@@textobjage", "TAC Tórax/Abdomen/Pelvis");
                htmlBody = htmlBody.Replace("@@telefono", "<p style=\"margin: 0;\">Si desea realizar cambios a su hora puede contactarnos al<span style=\"color: #dd3488;\"><a href=\"tel:22 994 4003\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a></span></p>");

            }

            if (sistema == "70" && subtipoobj == "93")
            {
                htmlBody = htmlBody.Replace("@@textobjage", "TAC Extremidades (Rodilla, Cadera, Pierna, Muslo, Hombro, Brazo, Mano, Muñeca)");
                htmlBody = htmlBody.Replace("@@telefono", "<p style=\"margin: 0;\">Si desea realizar cambios a su hora puede contactarnos al<span style=\"color: #dd3488;\"><a href=\"tel:22 994 4003\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a></span></p>");

            }

            if (sistema == "70" && subtipoobj == "94")
            {
                htmlBody = htmlBody.Replace("@@textobjage", "Angio TAC (Cerebro, Cuello, Tórax, Abdomen, Pelvis)");
                htmlBody = htmlBody.Replace("@@telefono", "<p style=\"margin: 0;\">Si desea realizar cambios a su hora puede contactarnos al<span style=\"color: #dd3488;\"><a href=\"tel:22 994 4003\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a></span></p>");

            }

            if (sistema == "70" && subtipoobj == "95")
            {
                htmlBody = htmlBody.Replace("@@textobjage", "Otros TAC");
                htmlBody = htmlBody.Replace("@@telefono", "<p style=\"margin: 0;\">Si desea realizar cambios a su hora puede contactarnos al<span style=\"color: #dd3488;\"><a href=\"tel:22 994 4003\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a></span></p>");

            }

            if (sistema == "65" && subtipoobj == "401010")
            {
                htmlBody = htmlBody.Replace("@@textobjage", "Mamografía Bilateral (4 EXP.)");
                htmlBody = htmlBody.Replace("@@telefono", "<p style=\"margin: 0;\">Si desea realizar cambios a su hora puede contactarnos al<span style=\"color: #dd3488;\"><a href=\"tel:22 994 4003\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a></span></p>");

            }

            if (sistema == "65" && subtipoobj == "401110")
            {
                htmlBody = htmlBody.Replace("@@textobjage", "Mamografía Unilateral (2 EXP.)");
                htmlBody = htmlBody.Replace("@@telefono", "<p style=\"margin: 0;\">Si desea realizar cambios a su hora puede contactarnos al<span style=\"color: #dd3488;\"><a href=\"tel:22 994 4003\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a></span></p>");

            }

            if (sistema == "83")
            {
                htmlBody = htmlBody.Replace("@@textobjage", "Test de Esfuerzo");
                htmlBody = htmlBody.Replace("@@telefono", "<p style=\"margin: 0;\">Si desea realizar cambios a su hora puede contactarnos al<span style=\"color: #dd3488;\"><a href=\"tel:22 994 4003\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a></span></p>");
            }

            //if (sistema == "57")
            //{
            //    htmlBody = htmlBody.Replace("@@textobjage", "Radiografía");
            //    htmlBody = htmlBody.Replace("@@telefono", "<p style=\"margin: 0;\">Si desea realizar cambios a su hora puede contactarnos al<span style=\"color: #dd3488;\"><a href=\"tel:22 994 4003\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a></span></p>");
            //}

            if (sistema == "57" && subtipoobj == "70")
            {
                htmlBody = htmlBody.Replace("@@textobjage", "Radiografía General");
                htmlBody = htmlBody.Replace("@@telefono", "<p style=\"margin: 0;\">Si desea realizar cambios a su hora puede contactarnos al<span style=\"color: #dd3488;\"><a href=\"tel:22 994 4003\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a></span></p>");

            }

            if (sistema == "57" && subtipoobj == "72")
            {
                htmlBody = htmlBody.Replace("@@textobjage", "Radiografía Proyecciones");
                htmlBody = htmlBody.Replace("@@telefono", "<p style=\"margin: 0;\">Si desea realizar cambios a su hora puede contactarnos al<span style=\"color: #dd3488;\"><a href=\"tel:22 994 4003\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a></span></p>");

            }


            if (sistema == "55")
            {
                htmlBody = htmlBody.Replace("@@textobjage", "Kinesiología");
                htmlBody = htmlBody.Replace("@@telefono", "<p style=\"margin: 0;\">Si desea realizar cambios a su hora puede contactarnos al<span style=\"color: #dd3488;\"><a href=\"tel:22 994 4003\" target=\"_blank\" style=\"text-decoration: underline; color: #dd3488;\" rel=\"noopener\">22 994 4003.&nbsp;</a></span></p>");
            }

            htmlBody = htmlBody.Replace("@nombre", tnombre);
            htmlBody = htmlBody.Replace("@fecha", fecha);
            htmlBody = htmlBody.Replace("@hora", hora);
            htmlBody = htmlBody.Replace("@direccion", direccion);
            htmlBody = htmlBody.Replace("@sucursal", sucursal);
            htmlBody = htmlBody.Replace("@ubicacion", ubicacion);

            // Crear una vista alternativa para el cuerpo HTML
            AlternateView htmlView = AlternateView.CreateAlternateViewFromString(htmlBody, null, MediaTypeNames.Text.Html);

            // Cargar las imágenes en LinkedResources
            LinkedResource imageResource = new LinkedResource(imagePath, MediaTypeNames.Image.Jpeg) { ContentId = imageFileName, ContentType = { Name = imageFileName } };
            htmlView.LinkedResources.Add(imageResource);

            LinkedResource imageResource2 = new LinkedResource(imagePath2, MediaTypeNames.Image.Jpeg) { ContentId = imageFileName2, ContentType = { Name = imageFileName2 } };
            htmlView.LinkedResources.Add(imageResource2);

            LinkedResource imageResource3 = new LinkedResource(imagePath3, MediaTypeNames.Image.Jpeg) { ContentId = imageFileName3, ContentType = { Name = imageFileName3 } };
            htmlView.LinkedResources.Add(imageResource3);

            LinkedResource imageResource4 = new LinkedResource(imagePath4, MediaTypeNames.Image.Jpeg) { ContentId = imageFileName4, ContentType = { Name = imageFileName4 } };
            htmlView.LinkedResources.Add(imageResource4);

            LinkedResource imageResource5 = new LinkedResource(imagePath5, MediaTypeNames.Image.Jpeg) { ContentId = imageFileName5, ContentType = { Name = imageFileName5 } };
            htmlView.LinkedResources.Add(imageResource5);

            // Agregar las imágenes adicionales solo si el sistema es específico
            if (sistema == "70")
            {
                LinkedResource imageResource6 = new LinkedResource(imagePath6, MediaTypeNames.Image.Jpeg) { ContentId = imageFileName6, ContentType = { Name = imageFileName6 } };
                htmlView.LinkedResources.Add(imageResource6);
                htmlBody = htmlBody.Replace("@@imagen6", htmlimagen6);
            }
            else
            {
                htmlBody = htmlBody.Replace("@@imagen6", ""); // Elimina el marcador si no es sistema 70
            }

            if (sistema == "64")
            {
                LinkedResource imageResource7 = new LinkedResource(imagePath7, MediaTypeNames.Image.Jpeg) { ContentId = imageFileName7, ContentType = { Name = imageFileName7 } };
                htmlView.LinkedResources.Add(imageResource7);
                htmlBody = htmlBody.Replace("@@imagen7", htmlimagen7);
            }
            else
            {
                htmlBody = htmlBody.Replace("@@imagen7", ""); // Elimina el marcador si no es sistema 64
            }

            if (sistema == "82" && subtipoobj == "43")
            {
                LinkedResource imageResource8 = new LinkedResource(imagePath8, MediaTypeNames.Image.Jpeg) { ContentId = imageFileName8, ContentType = { Name = imageFileName8 } };
                htmlView.LinkedResources.Add(imageResource8);
                htmlBody = htmlBody.Replace("@@imagen8", htmlimagen8);
            }

            else
            {
                htmlBody = htmlBody.Replace("@@imagen8", ""); 
            }

            if (sistema == "65")
            {
                LinkedResource imageResource9 = new LinkedResource(imagePath9, MediaTypeNames.Image.Jpeg) { ContentId = imageFileName9, ContentType = { Name = imageFileName9 } };
                htmlView.LinkedResources.Add(imageResource9);
                htmlBody = htmlBody.Replace("@@imagen9", htmlimagen9);
            }
            else
            {
                htmlBody = htmlBody.Replace("@@imagen9", ""); 
            }

            if (sistema == "76")
            {
                LinkedResource imageResource11 = new LinkedResource(imagePath11, MediaTypeNames.Image.Jpeg) { ContentId = imageFileName11, ContentType = { Name = imageFileName11 } };
                htmlView.LinkedResources.Add(imageResource11);
                htmlBody = htmlBody.Replace("@@imagen11", htmlimagen11);
            }
            else
            {
                htmlBody = htmlBody.Replace("@@imagen11", ""); 
            }

            if (sistema == "77")
            {
                LinkedResource imageResource10 = new LinkedResource(imagePath10, MediaTypeNames.Image.Jpeg) { ContentId = imageFileName10, ContentType = { Name = imageFileName10 } };
                htmlView.LinkedResources.Add(imageResource10);
                htmlBody = htmlBody.Replace("@@imagen10", htmlimagen10);
            }
            else
            {
                htmlBody = htmlBody.Replace("@@imagen10", "");
            }

            if (sistema == "78")
            {
                LinkedResource imageResource13 = new LinkedResource(imagePath13, MediaTypeNames.Image.Jpeg) { ContentId = imageFileName13, ContentType = { Name = imageFileName13 } };
                htmlView.LinkedResources.Add(imageResource13);
                htmlBody = htmlBody.Replace("@@imagen13", htmlimagen13);
            }
            else
            {
                htmlBody = htmlBody.Replace("@@imagen13", "");
            }

            if (sistema == "79")
            {
                LinkedResource imageResource12 = new LinkedResource(imagePath12, MediaTypeNames.Image.Jpeg) { ContentId = imageFileName12, ContentType = { Name = imageFileName12 } };
                htmlView.LinkedResources.Add(imageResource12);
                htmlBody = htmlBody.Replace("@@imagen12", htmlimagen12);
            }
            else
            {
                htmlBody = htmlBody.Replace("@@imagen12", "");
            }


            // Añadir la vista alternativa al mensaje de correo electrónico
            mail.AlternateViews.Add(htmlView);

            // Configurar el cliente SMTP y enviar el correo
            SmtpClient smtpClient = new SmtpClient(smtpServer)
            {
                Port = smtpPort,
                Credentials = new NetworkCredential(smtpUser, smtpPassword),
                EnableSsl = true // Asegurarse de que SSL esté habilitado para mayor seguridad
            };

            try
            {
                smtpClient.Send(mail);
            }
            catch (Exception ex)
            {
                return ex.Message;
            }

            return "";
        }



        // ESTA PARTE ESTA BUENA,ESTE ES UNA BARRERA DE SEGURIDAD POR SI BORRO ALGO 1 2 3 4 5 6

        static string GenerarToken(string citNumero)
        {
            using (var sha256 = System.Security.Cryptography.SHA256.Create())
            {
                byte[] hashBytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(citNumero + "secreto_seguro"));
                return Convert.ToBase64String(hashBytes).Replace("+", "").Replace("/", "").Replace("=", "");
            }
        }


        private static string UpdateEstadoEnvio(string tsecuencia)
        {

            string connectionString = ConfigurationSettings.AppSettings["conexion_BD_Ambulatorio"];
            SqlConnection conexion = new SqlConnection(connectionString);
            conexion.Open();



            string cadena = "UPDATE [dbo].[CORREOS_SEND]  set cestado = 1, dfecestado = '@value' where nsecuencia =  " + tsecuencia;

            DateTime dateTimeVariable = DateTime.Now;
            cadena = cadena.Replace("@value", dateTimeVariable.ToString());


            SqlCommand comando = new SqlCommand(cadena, conexion);

            comando.ExecuteNonQuery();
            conexion.Close();

            return "OK";

        }




        private static string EnvioCorreoAnula(string subject, string tcorreo, string fecha, string hora, string sucursal, string ubicacion, string sistema)
        {



            string senderEmail = System.Configuration.ConfigurationSettings.AppSettings.Get("mail_usuario");
            string senderPassword = System.Configuration.ConfigurationSettings.AppSettings.Get("mail_clave");



            string smtpServer = "in.smtpok.com";
            int smtpPort = 587;
            string smtpUser = "s182692_3";
            string smtpPassword = "7awmO.S4Xv";

            //string senderEmail = "agenda@clini.cl";
            string recipientEmail = tcorreo;

            // Create the email message
            System.Net.Mail.MailMessage mail = new System.Net.Mail.MailMessage(senderEmail, recipientEmail);
            mail.Subject = subject;

            //////////// Create the HTML body with an embedded image
            //////////string htmlBody = "<html><body>";
            //////////htmlBody += "<p>Su clave para ingresar a portal Bono CLINI es:  @clave .</p>";
            //////////htmlBody += "<p>Ingrese al Portal BonoClini <a href=\"https://www.portal-bono.clini.cl/bonoclini\" style=\"color: #DD3F93;\">Aqui.</a></p>";

            string imagePath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory) + @"img\logo.png";
            //string imagePath = @"C:\inetpub\wwwroot\RegistroBonoCLINI\img\logo.png";
            string imageFileName = "logo.png";

            string imagePath2 = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory) + @"img\facebook@2x.png";
            string imageFileName2 = "facebook@2x.png";


            string imagePath3 = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory) + @"img\linkedin@2x.png";
            string imageFileName3 = "linkedin@2x.png";


            string imagePath4 = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory) + @"img\instagram@2x.png";
            string imageFileName4 = "instagram@2x.png";

            string imagePath5 = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory) + @"img\youtube@2x.png";
            string imageFileName5 = "youtube@2x.png";







            string htmlimagen = $"<img src=\"cid:{imageFileName}\" alt=\"Embedded Image\"  style=\"display: block; height: auto; border: 0; width: 100%;\" width=\"180\" height=\"auto\">";
            string htmlimagen2 = $"<img src=\"cid:{imageFileName2}\" alt=\"Embedded Image\"  style=\"display: block; height: auto; border: 0; width: 100%;\" width=\"32\" height=\"auto\" alt=\"Facebook\" title=\"facebook\">";
            string htmlimagen3 = $"<img src=\"cid:{imageFileName3}\" alt=\"Embedded Image\"  style=\"display: block; height: auto; border: 0; width: 100%;\" width=\"32\" height=\"auto\" alt=\"Linkedin\" title=\"linkedin\">";
            string htmlimagen4 = $"<img src=\"cid:{imageFileName4}\" alt=\"Embedded Image\"  style=\"display: block; height: auto; border: 0; width: 100%;\" width=\"32\" height=\"auto\" alt=\"Instagram\" title=\"instagram\">";
            string htmlimagen5 = $"<img src=\"cid:{imageFileName5}\" alt=\"Embedded Image\"  style=\"display: block; height: auto; border: 0; width: 100%;\" width=\"32\" height=\"auto\" alt=\"YouTube\" title=\"YouTube\">";



            string htmlBody = File.ReadAllText(System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory) + @"html\anulacionImagen.html");
            if (sistema == "73" || sistema == "74" || sistema == "72" || sistema == "75")
            {

                htmlBody = File.ReadAllText(System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory) + @"html\anulacionImagenB.html");
            }


            htmlBody = htmlBody.Replace("@@imagen1", htmlimagen);
            htmlBody = htmlBody.Replace("@@imagen2", htmlimagen2);
            htmlBody = htmlBody.Replace("@@imagen3", htmlimagen3);
            htmlBody = htmlBody.Replace("@@imagen4", htmlimagen4);
            htmlBody = htmlBody.Replace("@@imagen5", htmlimagen5);




            htmlBody = htmlBody.Replace("@fecha", fecha);
            htmlBody = htmlBody.Replace("@hora", hora);
            htmlBody = htmlBody.Replace("@sucursal", sucursal);
            htmlBody = htmlBody.Replace("@ubicacion", ubicacion);





            // Create an alternative view for the HTML body
            AlternateView htmlView = AlternateView.CreateAlternateViewFromString(htmlBody, null, MediaTypeNames.Text.Html);

            // 1 Load the image into a LinkedResource
            LinkedResource imageResource = new LinkedResource(imagePath, MediaTypeNames.Image.Jpeg);
            imageResource.ContentId = imageFileName;
            // 1 Add the LinkedResource to the alternate view
            htmlView.LinkedResources.Add(imageResource);


            // 2 Load the image into a LinkedResource
            LinkedResource imageResource2 = new LinkedResource(imagePath2, MediaTypeNames.Image.Jpeg);
            imageResource2.ContentId = imageFileName2;
            // 2 Add the LinkedResource to the alternate view
            htmlView.LinkedResources.Add(imageResource2);


            // 3 Load the image into a LinkedResource
            LinkedResource imageResource3 = new LinkedResource(imagePath3, MediaTypeNames.Image.Jpeg);
            imageResource3.ContentId = imageFileName3;
            // 3 Add the LinkedResource to the alternate view
            htmlView.LinkedResources.Add(imageResource3);


            // 4 Load the image into a LinkedResource
            LinkedResource imageResource4 = new LinkedResource(imagePath4, MediaTypeNames.Image.Jpeg);
            imageResource4.ContentId = imageFileName4;
            // 4 Add the LinkedResource to the alternate view
            htmlView.LinkedResources.Add(imageResource4);



            // 5 Load the image into a LinkedResource
            LinkedResource imageResource5 = new LinkedResource(imagePath5, MediaTypeNames.Image.Jpeg);
            imageResource5.ContentId = imageFileName5;
            // 5 Add the LinkedResource to the alternate view
            htmlView.LinkedResources.Add(imageResource5);






            // Add the alternate view to the email message
            mail.AlternateViews.Add(htmlView);

            // Create and configure the SMTP client
            //SmtpClient smtpClient = new SmtpClient("smtp.gmail.com");
            //smtpClient.Port = 587; // For Gmail, use port 587
            //smtpClient.Credentials = new NetworkCredential(senderEmail, senderPassword);
            //smtpClient.EnableSsl = true;

            SmtpClient smtpClient = new SmtpClient(smtpServer)
            {
                Port = smtpPort,
                Credentials = new NetworkCredential(smtpUser, smtpPassword),
                EnableSsl = true
            };

            try
            {
                // Send the email
                smtpClient.Send(mail);
                ////Console.WriteLine("Email sent successfully!");
                ///

            }
            catch (Exception ex)
            {


            }

            return "";



        }

        private static void SincronizarCitaYCorreosSend(
            SqlConnection cnn,
            string nsecuencia,
            ref string sistema,
            ref string subtipoobj,
            ref string codCita,
            string identificador)
        {
            try
            {
                // 1) Número de cita
                string numeroCita = !string.IsNullOrWhiteSpace(codCita) ? codCita : identificador;
                if (string.IsNullOrWhiteSpace(numeroCita))
                    return;

                bool cerroConn = false;
                if (cnn.State != ConnectionState.Open)
                {
                    cnn.Open();
                    cerroConn = true;
                }

                // 2) Traer TOBJ_CODIGO y PRS_CODIGO
                string tobj = null;
                string prs = null;

                using (var cmd = new SqlCommand(@"
SELECT TOP 1 TOBJ_CODIGO, PRS_CODIGO
FROM [Bd_Ambulatorio].[dbo].[CITA_MEDICA]
WHERE CIT_NUMERO = @nro;", cnn))
                {
                    cmd.Parameters.AddWithValue("@nro", numeroCita);

                    using (var rd = cmd.ExecuteReader())
                    {
                        if (rd.Read())
                        {
                            object oTobj = rd["TOBJ_CODIGO"];
                            object oPrs = rd["PRS_CODIGO"];
                            tobj = (oTobj == DBNull.Value) ? null : Convert.ToString(oTobj);
                            prs = (oPrs == DBNull.Value) ? null : Convert.ToString(oPrs);
                        }
                    }
                }

                // 3) Asegurar cod_cita en CORREOS_SEND
                using (var upCod = new SqlCommand(@"
UPDATE [dbo].[CORREOS_SEND]
SET cod_cita = @nro
WHERE nsecuencia = @nsec;", cnn))
                {
                    upCod.Parameters.AddWithValue("@nro", numeroCita);
                    upCod.Parameters.AddWithValue("@nsec", nsecuencia);
                    upCod.ExecuteNonQuery();
                }
                codCita = numeroCita;

                // 4) Si hay TOBJ, comparar/actualizar sistema y subtipoobj
                if (!string.IsNullOrWhiteSpace(tobj))
                {
                    string sistemaTrim = (sistema ?? "").Trim();
                    if (!sistemaTrim.Equals(tobj.Trim(), StringComparison.OrdinalIgnoreCase))
                    {
                        using (var upSis = new SqlCommand(@"
UPDATE [dbo].[CORREOS_SEND]
SET sistema = @tobj, subtipoobj = @prs
WHERE nsecuencia = @nsec;", cnn))
                        {
                            upSis.Parameters.AddWithValue("@tobj", tobj);
                            upSis.Parameters.AddWithValue("@prs", (prs != null) ? (object)prs : DBNull.Value);
                            upSis.Parameters.AddWithValue("@nsec", nsecuencia);
                            upSis.ExecuteNonQuery();
                        }

                        // Reflejar para el envío inmediato
                        sistema = tobj;
                        subtipoobj = prs ?? "";
                    }
                }

                if (cerroConn) cnn.Close();
            }
            catch
            {
                // No cortar el flujo de envío si falla esta sincronización.
            }
        }




    }
}
