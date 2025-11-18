import win32com.client
import datetime

outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")
calendar = namespace.GetDefaultFolder(9)


def GetUpcomingEvents():
    today = datetime.datetime.now()
    EndDate = today + datetime.timedelta(days=7)

    items = calendar.Items
    items.Sort("[Start]")
    items = items.Restrict(
        "[Start] >= '{}' AND [Start] <= '{}'".format(
            today.strftime("%m/%d/%Y %H:%M %p"), EndDate.strftime("%m/%d/%Y %H:%M %p")
        )
    )
    return items


def ParseEvent(event):
    subject = event.Subject
    start = event.Start
    location = event.Location
    return subject, start, location


def ScheduleReminders(events):
    reminders = []
    for event in events:
        subject, start, location = ParseEvent(event)
        ReminderTime = start - datetime.timedelta(minutes=30)
        ReminderTime = ReminderTime.replace(tzinfo=None)
        now = datetime.datetime.now().replace(tzinfo=None)
        if ReminderTime > now:
            reminders.append((subject, start, location))
    SendEmailReminder(reminders)


def SendEmailReminder(reminders):
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.Subject = "Upcoming Meeting Reminders"
    body = "Good morning Mr Richardson. Here is your schedule for the week sir. Have a blessed one.\n\n"
    for subject, start, location in reminders:
        body += f"Subject: {subject}\nStart Time: {start}\nLocation: {location}\n\n"
    mail.Body = body
    mail.To = "EmailAddressHere"
    mail.Send()


if __name__ == "__main__":
    events = GetUpcomingEvents()
    ScheduleReminders(events)
