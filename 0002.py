import win32com.client

#Setting up a Meeting Request

oOutlook = win32com.client.Dispatch("Outlook.Application")

appt = oOutlook.CreateItem(1) # 1 - olAppointmentItem
appt.Start = '2015-01-01 09:30'
appt.Subject = 'Wake up'
appt.Duration = 15
appt.Location = 'Home'
appt.MeetingStatus = 1 # 1 - olMeeting; Changing the appointment to meeting
#only after changing the meeting status recipients can be added
appt.Recipients.Add("jan.de.vos@unit4.com")
appt.Body = 'This is the text'
appt.ReminderMinutesBeforeStart = 15
appt.ReminderSet = 1
appt.Save()
appt.Send()
print("Done")

