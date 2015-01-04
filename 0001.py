import win32com.client

#Setting up an Appointment
oOutlook = win32com.client.Dispatch("Outlook.Application")

appt = oOutlook.CreateItem(1) # 1 - olAppointmentItem
appt.Start = '2012-01-28 17:00'
appt.Subject = 'Follow Up Meeting'
appt.Duration = 15
appt.Location = 'Office - Room 132A'
appt.Save()
print("Done")

