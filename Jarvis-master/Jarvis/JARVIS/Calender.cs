using System;
using Microsoft.Office.Interop.Outlook;
using System.Threading;

namespace JARVIS
{
    public class Calender
    {
        Speak jarvis;
        Application outLookApp = new Application();
        public void CalenderAppointments()
        {
            Console.WriteLine("Checking Calender");
            jarvis = new Speak();
            try
            {
                DateTime start = DateTime.Now;
                DateTime end = start.AddDays(7);
                int index = 1;
                bool appointmentFound = false;

                NameSpace outlookNameSpace = outLookApp.GetNamespace("MAPI");
                MAPIFolder calender = outlookNameSpace.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
                Items calenderItems = GetAppointmentsInRange(calender, start, end);
                if (calenderItems != null)
                {
                    Console.WriteLine("Appointments Found");
                    //set bool here change in foreach ask after foreach for change
                    foreach (AppointmentItem appointment in calenderItems)
                    {
                        Console.WriteLine("Appointment " + index.ToString());
                        string[] appointmentContent = { index.ToString(), appointment.Subject, appointment.Start.Date.ToString("d"), appointment.Start.TimeOfDay.ToString(), appointment.Location };
                        Thread sayAppointment = new Thread(new ThreadStart(() => jarvis.sayAppointment(appointmentContent)));
                        sayAppointment.IsBackground = true;
                        sayAppointment.Start();
                        Thread.Sleep(8000);
                        index++;
                        appointmentFound = true;
                    }

                    if (!appointmentFound)
                    {
                        Console.WriteLine("Nothing found");
                        NoAppointments();
                    }
                }
                else
                {
                    Console.WriteLine("Nothing found");
                    NoAppointments();

                }
            }
            catch
            {
                NoAppointments();
            }

        }

        public Items GetAppointmentsInRange(MAPIFolder folder, DateTime start, DateTime end)
        {
            string filter = "[Start] >= '" + start.ToString("g") + "' AND [End] <= '" + end.ToString("g") + "'";

            try
            {
                Items calItems = folder.Items;
                calItems.IncludeRecurrences = true;
                calItems.Sort("[Start]", Type.Missing);
                Items restictItems = calItems.Restrict(filter);
                if (restictItems.Count > 0)
                {
                    return restictItems;
                }
                else
                {
                    return null;
                }
            }
            catch(System.Exception ex)
            {
                Console.WriteLine(ex.ToString());
                return null;
            }

        }

        private void NoAppointments()
        {
            Thread noAppointment = new Thread(new ThreadStart(() => jarvis.sayAppointmentError()))
            {
                IsBackground = true
            };
            noAppointment.Start();
        }
    }
}
