using System;

namespace OutlookMeetingScraper
{
    public class Meeting
    {
        private readonly DateTime m_StartTime;
        private readonly TimeSpan m_Duration;
        private readonly string m_Name;

        public Meeting(DateTime startTime, TimeSpan duration, string name)
        {
            m_StartTime = startTime;
            m_Duration = duration;
            m_Name = name;
        }

        public DateTime StartTime
        {
            get { return m_StartTime; }
        }

        public TimeSpan Duration
        {
            get { return m_Duration; }
        }

        public string Name
        {
            get { return m_Name; }
        }
    }
}