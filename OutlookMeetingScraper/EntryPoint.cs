using System;
using System.Collections.Generic;
using System.IO;
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
                System.Console.WriteLine("Looks like I need to connect to Exchange. Enter your password.  I won't steal it. Honest");
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

            using (var streamWriter = File.CreateText("accepted_meetings.csv"))
            {
                Reports.AcceptedMeetings(meetings, streamWriter);
            }

            using (var output = File.CreateText("accepted_meetings_per_week.csv"))
            {
                Reports.TotalMeetingDuration(meetings, startDate, output);
            }

            using (var output = File.CreateText("contiguous_time_per_day.csv"))
            {
                Reports.ContiguousTimePerDay(output, meetings);
            }
        }
    }
}
