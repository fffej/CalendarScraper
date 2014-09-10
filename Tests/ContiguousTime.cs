using System;
using NUnit.Framework;
using OutlookMeetingScraper;

namespace Tests
{
    [TestFixture]
    public class ContiguousTime
    {
        [Test]
        public void NoMeetingsIs8HoursFreeTimePerDay()
        {
            Assert.That(ContinuousFreeTime(), Is.EqualTo(TimeSpan.FromHours(8)));
        }

        [Test]
        public void ASingleHourLongMeetingPerDay()
        {
            Assert.That(ContinuousFreeTime(MeetingAt(16, TimeSpan.FromHours(1))), Is.EqualTo(TimeSpan.FromHours(7)));
        }

        [Test]
        public void AMeetingInTheMiddleOfTheDay()
        {
            Assert.That(ContinuousFreeTime(MeetingAt(12, TimeSpan.FromHours(1))), Is.EqualTo(TimeSpan.FromHours(4)));
            Assert.That(ContinuousFreeTime(MeetingAt(11, TimeSpan.FromHours(2))), Is.EqualTo(TimeSpan.FromHours(4)));
            Assert.That(ContinuousFreeTime(MeetingAt(11, TimeSpan.FromHours(1))), Is.EqualTo(TimeSpan.FromHours(5)));
        }

        private static TimeSpan ContinuousFreeTime(params Meeting[] meetings)
        {
            return Reports.ContiguousFreeTimePerDay(meetings);
        }

        private static Meeting MeetingAt(int startHour, TimeSpan duration)
        {
            return new Meeting(DateTime.Today.AddHours(startHour), duration, "Irrelevant name");
        }
    }
}
