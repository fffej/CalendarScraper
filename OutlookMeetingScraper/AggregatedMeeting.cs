using System;

namespace OutlookMeetingScraper
{
    public class AggregatedMeeting
    {
        private readonly int m_TotalMeetings;
        private readonly TimeSpan m_Duration;

        public static readonly AggregatedMeeting Zero = new AggregatedMeeting(0,TimeSpan.Zero);

        public AggregatedMeeting(int totalMeetings, TimeSpan duration)
        {
            this.m_TotalMeetings = totalMeetings;
            this.m_Duration = duration;
        }

        public int TotalMeetings
        {
            get { return m_TotalMeetings; }
        }

        public TimeSpan Duration
        {
            get { return m_Duration; }
        }
    }
}