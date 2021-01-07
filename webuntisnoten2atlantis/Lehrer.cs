// Published under the terms of GPLv3 Stefan Bäumer 2020.

using System;
using System.Collections.Generic;
using Microsoft.Exchange.WebServices.Data;

namespace webuntisnoten2atlantis
{
    internal class Lehrer
    {
        public Lehrer()
        {
        }

        public string Kuerzel { get; internal set; }
        public string Mail { get; internal set; }

        internal void ToOutlook(ExchangeService service, Termin konferenz)
        {
            try
            {
                Appointment appointment = new Appointment(service);
                appointment.Subject = "Zeugniskonferenz " + konferenz.Klasse;                
                appointment.Body = "<b>Zeugniskonferenz " + konferenz.Klasse + "</b></br>" +
                    "<ul>" +                    
                    "<li>Einen Gesamtüberblick finden Sie <a href='https://teams.microsoft.com/l/file/223700AB-6912-4F12-B7E9-C536B9E3401C?tenantId=bde93bf2-f69b-4968-8d34-68e9231b31be&fileType=xlsx&objectUrl=https%3A%2F%2Fbkborken.sharepoint.com%2Fsites%2FKollegium2%2FFreigegebene%20Dokumente%2FGeneral%2F03%20Schulleitung%2F3.04%20Termine%2F2020-21%2F14%20Zeugniskonferenzen%20HZ%2FZeugniskonferenzen%20HZ.xlsx&baseUrl=https%3A%2F%2Fbkborken.sharepoint.com%2Fsites%2FKollegium2&serviceName=teams&threadId=19:f1e523f3fca441b6bafdaab82d52bdd1@thread.tacv2&groupId=f50b2866-8ad0-4022-a0b1-8caf6cfafd0f'>hier</a>." +
                    "<li>Die Eintragungen sind als GESAMTNOTE in Webuntis vorzunehmen. " +
                     "Wie das genau geht, wird in einem <a href='https://teams.microsoft.com/l/channel/19%3A801ed4cf882f45d4b3e55b6b4050ac43%40thread.tacv2/tab%3A%3A23c4312b-e90b-4c8f-b1f1-dcfc390e6f2b?groupId=f50b2866-8ad0-4022-a0b1-8caf6cfafd0f&tenantId=bde93bf2-f69b-4968-8d34-68e9231b31be'>Video</a> erklärt. " +
                     "Wichtig ist, dass die Teilleistungen (sonstige Leistungen und schriftlichen Arbeiten) in Webuntis eingetragen wurden. Ansonsten werden keine Gesamtnoten übertragen." +
                    "<li>Die gesamte Ablauforganistion zum Nachlesen gibt es <a href='https://bkborken.sharepoint.com/:w:/s/Kollegium2/EUo-RDRcCWZAuHn4m6euDf0BSdnk6ddRx6X9TvpX7dYPrA?e=FHUx6C'>hier</a>." +
                    "</ul>";
                appointment.Start = konferenz.Uhrzeit;
                appointment.End = appointment.Start.AddMinutes(10);
                appointment.Location = konferenz.Raum;
                appointment.RequiredAttendees.Add(this.Mail);
                appointment.Categories.Add("ZKHZ");
                appointment.ReminderMinutesBeforeStart = 240;

                string leh = "";
                foreach (var l in konferenz.Lehrers)
                {
                    appointment.RequiredAttendees.Add(l.Mail);
                    leh = leh + l.Kuerzel + ",";
                }

                Console.Write("Meeting " + konferenz.Uhrzeit.ToShortDateString() + "(" + konferenz.Uhrzeit.ToShortTimeString() + "-" + konferenz.Uhrzeit.AddMinutes(10).ToShortTimeString() + ") " + konferenz.Raum + " " + konferenz.Klasse + "(" + leh.TrimEnd(',') + ")" + " in Outlook anlegen?");
                
                var meeting = Console.ReadKey();
                Console.WriteLine("");

                if (meeting.Key == ConsoleKey.J || meeting.Key == ConsoleKey.Enter)
                {
                    appointment.Save(SendInvitationsMode.SendToNone);
                    Global.PrintMessage("Meeting " + konferenz.Uhrzeit.ToShortDateString() + "(" + konferenz.Uhrzeit.ToShortTimeString() + "-" + konferenz.Uhrzeit.AddMinutes(10).ToShortTimeString() + ") " + konferenz.Raum + " " + konferenz.Klasse + "(" + leh.TrimEnd(',') + ")" + " in Outlook angelegt.");
                }
            }
            catch (System.Exception ex)
            {                
            }


            //service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, this.Mail);

            //Appointments appointmentsIst = new Appointments(this.Mail, bildungsgang, service);

            //Appointments appointmentsSoll = new Appointments(zeit, raum, klassenImBildungsgang, bildungsgang, service);

            //appointmentsIst.DeleteAppointments(appointmentsSoll, zeit, raum, klassenImBildungsgang, bildungsgang);

            //appointmentsSoll.AddAppointments(appointmentsIst, this, service, zeit, raum, klassenImBildungsgang, bildungsgang);
        }
    }
}