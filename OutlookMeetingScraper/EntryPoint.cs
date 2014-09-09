using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
// ReSharper disable once PossibleNullReferenceException
            var userName = System.Console.ReadLine().Trim();

            System.Console.WriteLine("Enter your password.  I won't steal it honest");
            var password = Console.ReadPassword();

            var service = new ExchangeService(ExchangeVersion.Exchange2007_SP1)
            {
                Credentials = new WebCredentials(userName, password),
                TraceEnabled = false,
                TraceFlags = TraceFlags.None
            };

            service.AutodiscoverUrl(userName, EwsExample.RedirectionUrlValidationCallback);    
      
            var meetings = new ExchangeRetriever(service).MeetingStatistics(DateTime.Now - TimeSpan.FromDays(365 * 3), DateTime.Now);

            using (var output = File.CreateText("c:/temp/accepted_meetings.csv"))
            {
                output.WriteLine("Meeting name,Meeting Date,Duration (hours)");
                foreach (var meeting in meetings)
                {
                    output.WriteLine("{0},{1},{2}", meeting.Name, meeting.StartTime, meeting.Duration.TotalHours);
                }
            }
        }
    }

    public class ExchangeRetriever
    {
        private readonly ExchangeService m_ExchangeService;        

        public ExchangeRetriever(ExchangeService exchangeService)
        {
            m_ExchangeService = exchangeService;
        }

        public List<Meeting> MeetingStatistics(DateTime startDate, DateTime endDate)
        {
            var allMeetings = new List<Meeting>();

            while (startDate < endDate)
            {
                System.Console.WriteLine("Retrieving for {0}", startDate);

                var calendar = CalendarFolder.Bind(m_ExchangeService, WellKnownFolderName.Calendar, new PropertySet());

                // Set the start and end time and number of appointments to retrieve.
                var cView = new CalendarView(startDate, startDate + TimeSpan.FromDays(28))
                {
                    PropertySet = new PropertySet(
                        ItemSchema.Subject,
                        AppointmentSchema.Start,
                        AppointmentSchema.Duration,
                        AppointmentSchema.MyResponseType
                    )
                };

                // Retrieve a collection of appointments by using the calendar view.
                var meetings = calendar.FindAppointments(cView);

                allMeetings.AddRange(
                    meetings.Where(x => x.MyResponseType == MeetingResponseType.Accept).Select(meeting => new Meeting(meeting.Start, meeting.Duration, meeting.Subject))
                );

                startDate += TimeSpan.FromDays(28);
            }

            return allMeetings;
        }
    }
}
