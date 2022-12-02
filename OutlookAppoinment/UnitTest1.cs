using Microsoft.Office.Interop.Outlook;
using Microsoft.VisualBasic;
using System.Security.Cryptography.X509Certificates;
using static System.Formats.Asn1.AsnWriter;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAppoinment
{
    public class Tests
    {
        [SetUp]
        public void Setup()
        {
            //HACK https://stackoverflow.com/questions/58130446/net-core-3-0-and-ms-office-interop
        }

        [Test]
        public void Test1()
        {
            Outlook.Application application = new Outlook.Application();



            var Appointment = application.CreateItem(Outlook.OlItemType.olAppointmentItem);//as Outlook.OlItemType;
            Outlook.AppointmentItem newAppointment = (Outlook.AppointmentItem)application.CreateItem(Outlook.OlItemType.olAppointmentItem);
            newAppointment.Start =  DateTime.Now.AddHours(2);
            newAppointment.End = DateTime.Now.AddHours(3);
            newAppointment.Subject = "Test "+Guid.NewGuid().ToString();
            newAppointment.Location = "Test " + Guid.NewGuid().ToString();
            newAppointment.Body = "Test " + Guid.NewGuid().ToString();
            newAppointment.AllDayEvent = false;
            //newAppointment.Recipients.Add(application.Session.CurrentU ser.Name);
             newAppointment.Recipients.Add("joche.ojeda@bitframeworks.com");
            //Outlook.Recipients sentTo = newAppointment.Recipients;
            //sentTo.ResolveAll();
            newAppointment.Save();
            //newAppointment.Display(true);
            newAppointment.Close(OlInspectorClose.olSave);
            //Microsoft.Office.Interop.Outlook.Application outlookApplication = GetApplicationObject();
            //var my_Account = (Account)GetAccountForEmailAddress(outlookApplication, DefaultEmailAddress);
            //newAppointment = (AppointmentItem)my_Account.Application.CreateItem(OlItemType.olAppointmentItem);

            //var recurrencePattern = newAppointment.GetRecurrencePattern();
            //recurrencePattern.RecurrenceType = OlRecurrenceType.olRecursYearly;
            //recurrencePattern.PatternStartDate = dtStartDate; // new DateTime(2015, 01, 13);
            //recurrencePattern.PatternEndDate = dtEndDate; // new DateTime(2015, 03, 20);
            //recurrencePattern.StartTime = dtStartTime;
            //recurrencePattern.EndTime = dtEndTime;
            //recurrencePattern.MonthOfYear = dtStartDate.Month;
            //recurrencePattern.DayOfMonth = dtStartDate.Day;
            //recurrencePattern.NoEndDate = true;
            //recurrencePattern.Duration = 1440;

            //newAppointment.Categories = strCategory;
            //newAppointment.Start = dtStartDate;
            //newAppointment.End = dtEndDate;
            //newAppointment.Location = strLocation;
            //newAppointment.Body = strBody;
            //newAppointment.AllDayEvent = blAllDayEvent;
            //newAppointment.ReminderSet = blReminder;
            //newAppointment.ReminderMinutesBeforeStart = intMinutes;
            //newAppointment.Subject = strSubject;
            //newAppointment.Recipients.Add(strTo);
            //newAppointment.MeetingStatus = OlMeetingStatus.olMeeting;
            //newAppointment.Save();

        }
    }
}