using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace OutlookMeetingScraper
{
    public static class Reports
    {
        public static void ContiguousTimePerDay(StreamWriter output, IEnumerable<Meeting> meetings)
        {
            output.WriteLine("Date,Total Free Time (hours)");

            foreach (var meetingsGroupedByDay in meetings.GroupBy(x => x.StartTime.Date))
            {
                var timePerDay = ContiguousFreeTimePerDay(meetingsGroupedByDay);
                output.WriteLine("{0},{1}", meetingsGroupedByDay.Key.Date, timePerDay.TotalHours);
            }
        }

        public static TimeSpan ContiguousFreeTimePerDay(IEnumerable<Meeting> meetingsInASingleDay)
        {
            var inASingleDay = meetingsInASingleDay as Meeting[] ?? meetingsInASingleDay.ToArray();
            
            var day = inASingleDay.FirstOrDefault();
            if (day == null)
            {
                return TimeSpan.FromHours(8);
            }

            var meetings = inASingleDay.Concat(new[]
            {
                new Meeting(day.StartTime.Date + TimeSpan.FromHours(9), TimeSpan.Zero, "Dummy meeting to mark start of day"),
                new Meeting(day.StartTime.Date + TimeSpan.FromHours(17), TimeSpan.Zero, "Dummy meeting to mark end of day"),
            }).OrderBy(x => x.StartTime).ToArray();

            var bestTimeBetweenMeetings = TimeSpan.FromHours(0);
            for (var i = 0; i < meetings.Length - 1; ++i)
            {
                var timeBetweenMeetings = meetings[i + 1].StartTime - (meetings[i].StartTime + meetings[i].Duration);
                bestTimeBetweenMeetings = TimeSpan.FromSeconds(Math.Max(timeBetweenMeetings.TotalSeconds, bestTimeBetweenMeetings.TotalSeconds));
            }

            return bestTimeBetweenMeetings;
        }

        public static void TotalMeetingDuration(IEnumerable<Meeting> meetings, DateTime startDate, StreamWriter output)
        {
            output.WriteLine("Week Number,Total Meeting Count,Duration (hours)");
            var week = 0;

            foreach (var meetingsGroupedByWeek in meetings.GroupBy(x => ((int)(Math.Floor((x.StartTime - startDate).TotalDays) / 7.0))))
            {
                var total = meetingsGroupedByWeek.Aggregate(
                    AggregatedMeeting.Zero,
                    (current, meetingInAWeek) => new AggregatedMeeting(current.TotalMeetings + 1, current.Duration + meetingInAWeek.Duration));

                output.WriteLine("{0},{1},{2}", week++, total.TotalMeetings, total.Duration.TotalHours);
            }
        }

        public static void AcceptedMeetings(IEnumerable<Meeting> meetings, StreamWriter output)
        {
            output.WriteLine("Meeting name,Meeting Date,Duration (hours)");
            foreach (var meeting in meetings)
            {
                output.WriteLine("{0},{1},{2}", meeting.Name.Replace(',', ' '), meeting.StartTime,meeting.Duration.TotalHours);
            }
        }
    }
}