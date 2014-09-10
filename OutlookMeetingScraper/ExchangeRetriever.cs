using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Exchange.WebServices.Data;

namespace OutlookMeetingScraper
{
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

            while (startDate.DayOfWeek != DayOfWeek.Monday)
            {
                startDate -= TimeSpan.FromDays(1);
            }

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
                        AppointmentSchema.MyResponseType,
                        AppointmentSchema.IsAllDayEvent
                        )
                };

                // Retrieve a collection of appointments by using the calendar view.
                var meetings = calendar.FindAppointments(cView);

                

                allMeetings.AddRange(
                    meetings.Where(x => x.MyResponseType == MeetingResponseType.Accept).Select(meeting => 
                        new Meeting(meeting.Start, meeting.Duration > TimeSpan.FromHours(8) ? TimeSpan.FromHours(8) : meeting.Duration, meeting.Subject)
                        )
                    );

                startDate += TimeSpan.FromDays(28);
            }

            return allMeetings;
        }
    }
}