using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using Microsoft.Exchange.WebServices.Data;
using Newtonsoft.Json;

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

            var startDate = DateTime.Now - TimeSpan.FromDays(365*6);
            while (startDate.DayOfWeek != DayOfWeek.Monday)
            {
                startDate -= TimeSpan.FromDays(1);
            }

            List<Meeting> meetings = null;
            if (File.Exists(userName))
            {
                meetings = JsonConvert.DeserializeObject<List<Meeting>>(File.ReadAllText(userName));
            }
            else
            {
                System.Console.WriteLine(
                    "Looks like I need to connect to Exchange. Enter your password.  I won't steal it. Honest");
                var password = Console.ReadPassword();

                var service = new ExchangeService(ExchangeVersion.Exchange2007_SP1)
                {
                    Credentials = new WebCredentials(userName, password),
                    TraceEnabled = false,
                    TraceFlags = TraceFlags.None
                };

                service.AutodiscoverUrl(userName, EwsExample.RedirectionUrlValidationCallback);

                meetings = new ExchangeRetriever(service).MeetingStatistics(startDate, DateTime.Now);
                File.WriteAllText(userName, JsonConvert.SerializeObject(meetings));
            }

            using (var streamWriter = File.CreateText("c:/temp/accepted_meetings.csv"))
            {
                AcceptedMeetings(meetings, streamWriter);
            }

            using (var output = File.CreateText("c:/temp/accepted_meetings_per_week.csv"))
            {
                TotalMeetingDuration(meetings, startDate, output);
            }

            using (var output = File.CreateText("c:/temp/contiguous_time_per_day.csv"))
            {
                ContiguousTimePerDay(output, meetings);
            }
        }

        private static void ContiguousTimePerDay(StreamWriter output, IEnumerable<Meeting> meetings)
        {
            output.WriteLine("Date,Total Free Time (hours)");

            foreach (var meetingsGroupedByDay in meetings.GroupBy(x => x.StartTime.Date))
            {
                // Assume that days run from 9am to 5pm
                var hours = new List<int> {9, 10, 11, 12, 13, 14, 15, 16, 17};

                foreach (var meeting in meetingsGroupedByDay)
                {
                    for (var hourToRemove = meeting.StartTime.Hour;
                        hourToRemove <= (meeting.StartTime + meeting.Duration).Hour;
                        hourToRemove++)
                    {
                        hours.Remove(hourToRemove);
                    }
                }

                // Now find the longest contiguous sequence
                var currentCount = 0;
                var max = 0;
                for (var i = 0; i < hours.Count - 1; ++i)
                {
                    if (hours[i] + 1 == hours[i + 1])
                    {
                        max++;
                    }
                    else
                    {
                        currentCount = Math.Max(max, currentCount);
                        max = 0;
                    }
                }
                currentCount = Math.Max(max, currentCount);

                output.WriteLine("{0},{1}", meetingsGroupedByDay.Key.Date, currentCount);
            }
        }

        private static void TotalMeetingDuration(IEnumerable<Meeting> meetings, DateTime startDate, StreamWriter output)
        {
            output.WriteLine("Week Number,Total Meeting Count,Duration (hours)");
            var week = 0;

            foreach (
                var meetingsGroupedByWeek in
                    meetings.GroupBy(x => ((int) (Math.Floor((x.StartTime - startDate).TotalDays)/7.0))))
            {
                var total = meetingsGroupedByWeek.Aggregate(
                    AggregatedMeeting.Zero,
                    (current, meetingInAWeek) =>
                        new AggregatedMeeting(current.TotalMeetings + 1, current.Duration + meetingInAWeek.Duration));

                output.WriteLine("{0},{1},{2}", week++, total.TotalMeetings, total.Duration.TotalHours);
            }
        }

        private static void AcceptedMeetings(IEnumerable<Meeting> meetings, StreamWriter output)
        {
            output.WriteLine("Meeting name,Meeting Date,Duration (hours)");
            foreach (var meeting in meetings)
            {
                output.WriteLine("{0},{1},{2}", meeting.Name.Replace(',', ' '), meeting.StartTime,
                    meeting.Duration.TotalHours);
            }
        }
    }
}
