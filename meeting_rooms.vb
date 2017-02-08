Sub meeting_rooms()
	Dim newMeeting As Object 
	Dim meetingResource As Outlook.Recipient 
	Dim room1Email, room2Email, room3Email, room4Email As String
	Dim room As Variant

	'### List of rooms ###
	room1Email = "room1@yourcompany.com"
	room2Email = "room2@yourcompany.com"
	room3Email = "room3@yourcompany.com"
	room4Email = "room4@yourcompany.com"
 
 	'### Create a new meeting ###
	Set newMeeting = Application.CreateItem(olAppointmentItem) 
	newMeeting.MeetingStatus = olMeeting 
	
	'### Iterate through rooms, adding them as resources ###
	For Each room in Array(room1Email, room2Email, room3Email, room4Email)
		Set meetingResource = newMeeting.Recipients.Add(room) 
		meetingResource.Type = olResource 
	Next room

	'### View message ###
	newMeeting.Display 
End Sub
