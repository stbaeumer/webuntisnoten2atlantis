// Published under the terms of GPLv3 Stefan Bäumer 2021.

using Microsoft.Exchange.WebServices.Data;
using System;

namespace webuntisnoten2atlantis
{
    internal class Termin
    {
        public Termin()
        {
        }

        public string Bildungsgang { get; internal set; }
        public string Klasse { get; internal set; }
        public DateTime Uhrzeit { get; internal set; }
        public string Raum { get; internal set; }
        public Lehrers Lehrers { get; internal set; }
        public dynamic Bereich { get; internal set; }

        internal void ToOutlook(ExchangeService service, string mail)
        {
            try
            {
                Appointment appointment = new Appointment(service);
                appointment.Subject = "Zeugniskonferenzen: " + Klasse;
                appointment.Body = "<b>Zeugniskonferenzen: " + Bildungsgang.TrimEnd(',') + "</b></br>" +
                    "<p>Sie erhalten diese Einladung zur Teams-Zeugniskonferenz, weil Sie in den Klassen " + Klasse.TrimEnd(',') + " unterrichten.</p>" +
                    "<p>Einen Gesamtüberblick mit allen Uhrzeiten und den wichtigsten Informationen zur Noteneintragung, Eintragung von Fehlzeiten, zu Zeugnissen und Zeugniskonferenzen finden Sie <a href='https://teams.microsoft.com/l/file/2819D9B8-1129-418E-83BD-A4B14C1F8422?tenantId=bde93bf2-f69b-4968-8d34-68e9231b31be&fileType=pdf&objectUrl=https%3A%2F%2Fbkborken.sharepoint.com%2Fsites%2FKollegium2%2FFreigegebene%20Dokumente%2FGeneral%2F03%20Schulleitung%2F3.04%20Termine%2F2020-21%2F14%20Zeugniskonferenzen%20HZ%2FZeugniskonferenzen%20HZ.pdf&baseUrl=https%3A%2F%2Fbkborken.sharepoint.com%2Fsites%2FKollegium2&serviceName=teams&threadId=19:f1e523f3fca441b6bafdaab82d52bdd1@thread.tacv2&groupId=f50b2866-8ad0-4022-a0b1-8caf6cfafd0f'>hier</a>. " +
                    "Dort sind auch die Hyperlinks zu den Teams-Zeugniskonferenzen eingebaut. " +
                    "Es wird nicht für jede Klasse eine eigene Teams-Besprechung geben. " +
                    "Stattdessen gibt es am Dienstag 3 Besprechungen, die alle um 15 Uhr starten. Jede Besprechung entspricht einem bisherigen Raum, den man bei Bedarf jederzeit betritt oder auch wieder verlässt. " +
                    "Sie können also nach Bedarf zwischen den 3 Teams-Besprechungen wechseln, indem Sie eine Besprechung verlassen und dann z. B. in dem <a href='https://teams.microsoft.com/l/file/2819D9B8-1129-418E-83BD-A4B14C1F8422?tenantId=bde93bf2-f69b-4968-8d34-68e9231b31be&fileType=pdf&objectUrl=https%3A%2F%2Fbkborken.sharepoint.com%2Fsites%2FKollegium2%2FFreigegebene%20Dokumente%2FGeneral%2F03%20Schulleitung%2F3.04%20Termine%2F2020-21%2F14%20Zeugniskonferenzen%20HZ%2FZeugniskonferenzen%20HZ.pdf&baseUrl=https%3A%2F%2Fbkborken.sharepoint.com%2Fsites%2FKollegium2&serviceName=teams&threadId=19:f1e523f3fca441b6bafdaab82d52bdd1@thread.tacv2&groupId=f50b2866-8ad0-4022-a0b1-8caf6cfafd0f'>Gesamtüberblick</a> auf den Hyperlink in bei einer parallel stattfinden Besprechung auf 'teilnehmen' klicken.</p>";
                    
                appointment.Start = Uhrzeit;
                appointment.End = appointment.Start.AddMinutes(10);
                appointment.Location = Raum;
                appointment.RequiredAttendees.Add(mail);
                appointment.RequiredAttendees.Add("ursula.moritz@berufskolleg-borken.de");
                appointment.RequiredAttendees.Add("klaus.lienenklaus@berufskolleg-borken.de");
                appointment.RequiredAttendees.Add("wolfgang.leuering@berufskolleg-borken.de");
                appointment.Categories.Add("Zeugniskonferenz");
                appointment.ReminderMinutesBeforeStart = 240;

                string leh = "";
                foreach (var l in Lehrers)
                {
                    appointment.RequiredAttendees.Add(l.Mail);
                    leh = leh + l.Kuerzel + ",";
                }

                Console.Write("Zeugniskonferenz " + Uhrzeit.ToShortDateString() + "(" + Uhrzeit.ToShortTimeString() + "-" + Uhrzeit.AddMinutes(10).ToShortTimeString() + ") " + Raum + " " + Klasse + "(" + leh.TrimEnd(',') + ")" + " in Outlook anlegen?");

                var meeting = Console.ReadKey();
                Console.WriteLine("");

                if (meeting.Key == ConsoleKey.J || meeting.Key == ConsoleKey.Enter)
                {
                    appointment.Save(SendInvitationsMode.SendToNone);
                    Global.PrintMessage(Global.Output.Count, "Zeugniskonferenz " + Uhrzeit.ToShortDateString() + "(" + Uhrzeit.ToShortTimeString() + "-" + Uhrzeit.AddMinutes(10).ToShortTimeString() + ") " + Raum + " " + Klasse + "(" + leh.TrimEnd(',') + ")" + " in Outlook angelegt.");
                }
            }
            catch (System.Exception)
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