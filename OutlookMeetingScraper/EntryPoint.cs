using System;
using System.Net;
using Microsoft.Exchange.WebServices.Data;

namespace OutlookMeetingScraper
{
    public class EntryPoint
    {
        public static void Main(string[] args)
        {
            ServicePointManager.ServerCertificateValidationCallback = EwsExample.CertificateValidationCallBack;

            System.Console.WriteLine("Enter your email address");
            var userName = System.Console.ReadLine();

            System.Console.WriteLine("Enter your password.  I won't steal it honest");
            var password = Console.ReadPassword();

            var service = new ExchangeService(ExchangeVersion.Exchange2007_SP1)
            {
                Credentials = new WebCredentials(userName, password),
                TraceEnabled = false,
                TraceFlags = TraceFlags.None
            };

            service.AutodiscoverUrl(userName, EwsExample.RedirectionUrlValidationCallback);    
      
            new ExchangeRetriever(service).MeetingStatistics();

            System.Console.ReadKey();
        }
    }

    public class ExchangeRetriever
    {
        private readonly ExchangeService m_ExchangeService;        

        public ExchangeRetriever(ExchangeService exchangeService)
        {
            m_ExchangeService = exchangeService;
        }

        public void MeetingStatistics()
        {
            DateTime startDate = DateTime.Now - TimeSpan.FromDays(10);
            DateTime endDate = DateTime.Now;

            // Initialize the calendar folder object with only the folder ID. 
            var calendar = CalendarFolder.Bind(m_ExchangeService, WellKnownFolderName.Calendar, new PropertySet());

            // Set the start and end time and number of appointments to retrieve.
            var cView = new CalendarView(startDate, endDate)
            {
                PropertySet = new PropertySet(
                    ItemSchema.Subject,
                    AppointmentSchema.Start,
                    AppointmentSchema.Duration
                )
            };

            // Retrieve a collection of appointments by using the calendar view.
            var meetings = calendar.FindAppointments(cView);

            System.Console.WriteLine("You've had {0} meetings since {1}", meetings.TotalCount, startDate);         
            TimeSpan totalDuration = new TimeSpan();

            foreach (var meeting in meetings)
            {
                if (meeting.MyResponseType == MeetingResponseType.Accept)
                    totalDuration += meeting.Duration;
            }

            System.Console.WriteLine("This has consumed {0} of your life", totalDuration.TotalHours.ToString("#.##"));

        }
    }
}
