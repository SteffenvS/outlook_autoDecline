VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} declineAllWindow 
   Caption         =   "Cancel All Appointments"
   ClientHeight    =   8775.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4605
   OleObjectBlob   =   "declineAllWindow.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "declineAllWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim activeTextbox As String



Enum ask_user
    ask_user_no = 0
    ask_user_day = 1
    ask_user_all = 2
End Enum

Private Sub CheckBox_appts_nonrecurring_Change()
    If CheckBox_appts_nonrecurring.Value = False Then
        CheckBox_appts_recurring.Value = True
    End If
End Sub

Private Sub CheckBox_appts_recurring_Change()
    If CheckBox_appts_recurring.Value = False Then
        CheckBox_appts_nonrecurring.Value = True
    End If
End Sub

Private Sub CheckBox_appts_invited_Change()
    If CheckBox_appts_invited.Value = False Then
        CheckBox_appts_organizer.Value = True
        checkbox_response.Enabled = False
    Else
        checkbox_response.Enabled = True
    End If
End Sub

Private Sub CheckBox_appts_organizer_Change()
    If CheckBox_appts_organizer.Value = False Then
        CheckBox_appts_invited.Value = True
        checkbox_notification.Enabled = False
    Else
        checkbox_notification.Enabled = True
    End If
End Sub

Private Sub CommandButton_p1_Click()
    newDate = modifyDateBy(1)
    writeDateToSelected (newDate)
    refocus
End Sub
Private Sub CommandButton_p7_Click()
    newDate = modifyDateBy(7)
    writeDateToSelected (newDate)
    refocus
End Sub
Private Sub CommandButton_m1_Click()
    newDate = modifyDateBy(-1)
    writeDateToSelected (newDate)
    refocus
End Sub
Private Sub CommandButton_m7_Click()
    newDate = modifyDateBy(-7)
    writeDateToSelected (newDate)
    refocus
End Sub

Private Function writeDateToFrom(modifiedDate As Date)
        textbox_from_TT.Value = Day(modifiedDate)
        textbox_from_MM.Value = Month(modifiedDate)
        textbox_from_JJJJ.Value = Year(modifiedDate)
        Label_from_weekday = Format(modifiedDate, "dddd")
        
        If modifiedDate > date_to Then
            writeDateToTo (modifiedDate)
        End If
            
End Function

Private Function writeDateToTo(modifiedDate As Date)
        textbox_to_TT.Value = Day(modifiedDate)
        textbox_to_MM.Value = Month(modifiedDate)
        textbox_to_JJJJ.Value = Year(modifiedDate)
        Label_to_weekday = Format(modifiedDate, "dddd")
        
        If modifiedDate < date_from Then
            writeDateToFrom (modifiedDate)
        End If
End Function

Private Function writeDateToSelected(modifiedDate As Date)
    
    If (0 = StrComp(Left(activeTextbox, Len("textbox_fr")), "textbox_fr")) Then
        writeDateToFrom (modifiedDate)
    Else
        writeDateToTo (modifiedDate)
    End If
    
End Function

Private Function modifyDateBy(modifyValue As Integer) As Date
    Select Case activeTextbox
    Case textbox_from_TT.Name
        modifyDateBy = DateAdd("d", modifyValue, date_from)
    Case textbox_from_MM.Name
        modifyDateBy = DateAdd("m", modifyValue, date_from)
    Case textbox_from_JJJJ.Name
        modifyDateBy = DateAdd("yyyy", modifyValue, date_from)
    Case textbox_to_TT.Name
        modifyDateBy = DateAdd("d", modifyValue, date_to)
    Case textbox_to_MM.Name
        modifyDateBy = DateAdd("m", modifyValue, date_to)
    Case textbox_to_JJJJ.Name
        modifyDateBy = DateAdd("yyyy", modifyValue, date_to)
    End Select
    
End Function

Private Function refocus()
    selectContent (activeTextbox)
    Dim currentTextbox As TextBox
    Set currentTextbox = declineAllWindow.Controls(activeTextbox)
    currentTextbox.SetFocus
End Function

Private Function selectContent(ByRef thisTextbox As String)
    Set currentTextbox = declineAllWindow.Controls(thisTextbox)
    With currentTextbox
        .SelStart = 0
        .SelLength = Len(.text)
    End With
End Function

Private Function setActiveTextbox(ByRef thisTextbox As String)
    activeTextbox = declineAllWindow.Controls(thisTextbox).Name
End Function

Private Sub CommandButton_reset_Click()
    UserForm_Initialize
End Sub

Private Sub textbox_from_MM_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If textbox_from_MM.Value > 12 Then
        textbox_from_MM.Value = 12
    End If
    If textbox_from_MM.Value < 1 Then
        textbox_from_MM.Value = 1
    End If
End Sub

Private Sub textbox_from_TT_Exit(ByVal Cancel As MSForms.ReturnBoolean)
     If textbox_from_TT.Value > 31 Then
        textbox_from_TT.Value = 31
    End If
    If textbox_from_TT.Value < 1 Then
        textbox_from_TT.Value = 1
    End If
End Sub

Private Sub textbox_to_MM_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If textbox_to_MM.Value > 12 Then
        textbox_to_MM.Value = 12
    End If
    If textbox_to_MM.Value < 1 Then
        textbox_to_MM.Value = 1
    End If
End Sub

Private Sub textbox_to_TT_Exit(ByVal Cancel As MSForms.ReturnBoolean)
     If textbox_to_TT.Value > 31 Then
        textbox_to_TT.Value = 31
    End If
    If textbox_to_TT.Value < 1 Then
        textbox_to_TT.Value = 1
    End If
End Sub

Private Sub textbox_from_JJJJ_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, _
ByVal X As Single, ByVal Y As Single)
    setActiveTextbox (textbox_from_JJJJ.Name)
    refocus
End Sub

Private Sub textbox_from_MM_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, _
ByVal X As Single, ByVal Y As Single)
    setActiveTextbox (textbox_from_MM.Name)
    refocus
End Sub

Private Sub textbox_from_TT_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, _
ByVal X As Single, ByVal Y As Single)
    setActiveTextbox (textbox_from_TT.Name)
    refocus
End Sub

Private Sub textbox_to_JJJJ_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, _
ByVal X As Single, ByVal Y As Single)
    setActiveTextbox (textbox_to_JJJJ.Name)
    refocus
End Sub

Private Sub textbox_to_MM_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, _
ByVal X As Single, ByVal Y As Single)
    setActiveTextbox (textbox_to_MM.Name)
    refocus
End Sub

Private Sub textbox_to_TT_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, _
ByVal X As Single, ByVal Y As Single)
    setActiveTextbox (textbox_to_TT.Name)
    refocus
End Sub

Private Sub UserForm_Initialize()
    'Call FormatUserForm(Me.Caption)
    setActiveTextbox (textbox_from_TT.Name)
    
    Dim datStart As Date
    Dim datEnd As Date
    Dim oView As Outlook.View
    Dim oCalView As Outlook.CalendarView
    Dim oExpl As Outlook.Explorer

    Set oExpl = Application.ActiveExplorer
    Set oView = oExpl.CurrentView

    If oView.ViewType = olCalendarView Then
        Set oCalView = oExpl.CurrentView

        datStart = oCalView.SelectedStartTime
        datEnd = oCalView.SelectedEndTime
    Else
        datStart = Date
        datEnd = DateAdd("d", 1, Date)
    End If
    
    writeDateToFrom (datStart)
    writeDateToTo (datEnd)

End Sub

Private Sub checkbox_response_Click()
    If checkbox_response.Value = False Then
        textbox_message.Enabled = False
    Else
        textbox_message.Enabled = True
    End If
    
End Sub

Private Sub button_cancel_Click()
    Me.Hide
    declineAllMeetings
    Me.Show
End Sub

Private Function date_from() As Date
    Dim date_from_TT, date_from_MM, date_from_JJJJ
    
    date_from_TT = CInt(textbox_from_TT.Value)
    date_from_MM = CInt(textbox_from_MM.Value)
    date_from_JJJJ = CInt(textbox_from_JJJJ.Value)
    
    date_from = DateValue(date_from_JJJJ & "." & date_from_MM & "." & date_from_TT)
End Function

Private Function set_date_from(theDate As Date)
    textbox_from_TT.Value = Day(theDate)
    textbox_from_MM.Value = Month(theDate)
    textbox_from_JJJJ.Value = Year(theDate)
End Function

Private Function set_date_to(theDate As Date)
    textbox_to_TT.Value = Day(theDate)
    textbox_to_MM.Value = Month(theDate)
    textbox_to_JJJJ.Value = Year(theDate)
End Function

Private Function date_to() As Date
    Dim date_to_TT, date_to_MM, date_to_JJJJ
    
    date_to_TT = CInt(textbox_to_TT.Value)
    date_to_MM = CInt(textbox_to_MM.Value)
    date_to_JJJJ = CInt(textbox_to_JJJJ.Value)
    
    date_to = DateValue(date_to_JJJJ & "." & date_to_MM & "." & date_to_TT)
End Function

Private Function getOcurrence(ByRef oAppt As Outlook.AppointmentItem) As Outlook.AppointmentItem
   
    If oAppt.IsRecurring Then
        Set getOcurrence = oAppt.GetRecurrencePattern.GetOccurrence(oAppt.Start)
    Else
        Set getOcurrence = oAppt
    End If
    
End Function

Private Function askUser() As Integer
    If OptionButton_no.Value = True Then
        askUser = ask_user_no
    ElseIf OptionButton_day = True Then
        askUser = ask_user_day
    ElseIf OptionButton_all = True Then
        askUser = ask_user_all
    End If
End Function
    
Public Sub declineAllMeetings()
    
    Dim myStart As Date
    Dim myEnd As Date
    Dim oCalendar As Outlook.Folder
    Dim oItems As Outlook.Items
    Dim oItemsInDateRange As Outlook.Items
    Dim oFinalItems As Outlook.Items
    Dim oAppt As Outlook.AppointmentItem
    Dim strRestriction As String

    myStart = date_from
    myEnd = DateAdd("h", 24, date_to) 'end of the day given is +24h

    Debug.Print "Start:", myStart
    Debug.Print "End:", myEnd
          
    'Construct filter for the next 30-day date range
    strRestriction = "[Start] >= '" & Format$(myStart, "dd/mm/yyyy") _
    & "' AND [End] <= '" & Format$(myEnd, "dd/mm/yyyy") & "'"
    
    Debug.Print strRestriction
    
    Set oCalendar = Application.Session.GetDefaultFolder(olFolderCalendar)
    Set oItems = oCalendar.Items
    oItems.IncludeRecurrences = True
    oItems.Sort "[Start]"
    
    'Restrict the Items collection for the date range
    Set oFinalItems = oItems.Restrict(strRestriction)
    
    'Sort and Debug.Print final results
    oFinalItems.Sort "[Start]"
    
    Dim currentDay As Date
    currentDay = Int(DateAdd("h", -24, Date))  'initialize with non-valid value
    
    answer = vbNo
    
    For Each oAppt In oFinalItems
        Debug.Print oAppt.Start, oAppt.Subject
        
        If oAppt.IsRecurring = True And CheckBox_appts_recurring.Value = False Then
            GoTo continue
        End If
        
        If oAppt.IsRecurring = False And CheckBox_appts_nonrecurring.Value = False Then
            GoTo continue
        End If
        
        If oAppt.AllDayEvent = True And CheckBox_appts_allDay.Value = False Then
            GoTo continue
        End If
        
        If oAppt.Organizer = Outlook.Session.CurrentUser And CheckBox_appts_organizer.Value = False Then
            GoTo continue
        End If
        
        If oAppt.Organizer <> Outlook.Session.CurrentUser And CheckBox_appts_invited.Value = False Then
            GoTo continue
        End If
        
        Dim oApptDay As Date
        oApptDay = Int(oAppt.Start)
        
        Dim text As String
        
        If askUser = ask_user.ask_user_all Then
            text = "Cancel " & oAppt.Start & " - " & oAppt.Subject
        ElseIf askUser = ask_user.ask_user_day And oApptDay <> currentDay Then
            Debug.Print "first of day"
            currentDay = oApptDay
            text = "Cancel all Meetings of Day " & Format(oApptDay, "dd.mm.yyyy")
        ElseIf askUser = ask_user.ask_user_day And oApptDay = currentDay Then
            'keep old answer
            GoTo skip_ask
        ElseIf askUser = ask_user.ask_user_no Then
            answer = vbYes
            GoTo skip_ask
        End If
        
        answer = MsgBox(text, vbQuestion + vbYesNoCancel + vbDefaultButton, "Decline Meeting")
skip_ask:
        
        If answer = vbYes Then
            Dim ocurrence As Outlook.AppointmentItem
            Set ocurrence = getOcurrence(oAppt)
            
            If ocurrence.Organizer = Outlook.Session.CurrentUser Then
                ocurrence.MeetingStatus = olMeetingCanceled
                ocurrence.Save
                If (checkbox_notification.Value = True) Then
                    ocurrence.Body = textbox_message.Value
                    ocurrence.Send
                End If
                ocurrence.Delete
            Else
                Set xMtResponse = oAppt.Respond(olMeetingDeclined, True)
                If checkbox_response.Value = True Then
                    xMtResponse.Body = textbox_message.Value
                    xMtResponse.Send
                End If
            End If
            

        ElseIf answer = vbCancel Then
            Exit For
        End If
        

continue:
    Next 'For Each oAppt In oFinalItems
    
    
    
End Sub



