Option Compare Database

Private Sub btn_ClientsByCity_Click()
    DoCmd.OpenQuery "qry_GetClientsByCity"
End Sub

Private Sub btn_SelectReport_Click()
    If IsNull(Me.combo_Reports.Value) Then
        
    ElseIf Me.combo_Reports.Value = "All active clients" Then
        DoCmd.OpenReport "rpt_AllActiveClients", acViewReport
    ElseIf Me.combo_Reports.Value = "All inactive clients" Then
        DoCmd.OpenReport "rpt_AllInactiveClients", acViewReport
    Else
        DoCmd.OpenReport "rpt_GetAllServices", acViewReport
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    'Resize the Window
    DoCmd.MoveSize , , 10.5 * 1440, 7.85 * 1440
    
    Dim vPermID As Long
    Dim vPermLevel As Integer
    Dim vPermName As String
    Dim PersonName As String
    Dim vDate As Date
    
    'Get today's day and month number
    Dim DayNumber As Integer
    DayNumber = Weekday(Date, vbMonday)
    Dim DayMonth As Integer
    DayMonth = Day(Date)
    Dim MonthNumber As Integer
    MonthNumber = Month(Date)
    
    Dim MonthStr As String
    MonthStr = MonthName(MonthNumber, False)
    
    'Set the string of the date
    Dim DisplayDate As String
    DisplayDate = WeekdayName(DayNumber, False, vbMonday)
    DisplayDate = DisplayDate & ", " & MonthStr & " " & CStr(DayMonth)
    
    vDate = Date
    
    Me.lbl_Time.Caption = CStr(Format(Time(), "Medium Time"))
    Me.lbl_DateNow.Caption = DisplayDate
    
    vPermID = DLookup("PermissionID", "Login", "PersonID = " & PersonIDi)
    vPermLevel = DLookup("PermissionLevel", "Permissions", "PermissionID = " & vPermID)
    vPermName = DLookup("Description", "Permissions", "PermissionID = " & vPermID)
       
       
    Me.lbl_PersonID.Caption = CStr(PersonIDi)
    Me.lbl_PermLevel.Caption = CStr(vPermLevel)
    
    
    PersonName = DLookup("FirstName", "Persons", "PersonID = " & LoggedInUserID)
    Dim surname As String
    surname = DLookup("Surname", "Persons", "PersonID = " & LoggedInUserID)
    
    
    'GET THE COLORS
    '
    '
    'Find the ID of the theme that will be used (IDTheme is a global variable found in Globals)
    IDTheme = DLookup("[ThemeID]", "Themes", "Persons![PersonID] = " & LoggedInUserID)
    
    'Find the background color value
    BGColor = DLookup("[BackgroundColor]", "Themes", "Themes![ThemeID] = " & IDTheme)
    
    'Find the menu background color
    MenuColor = DLookup("[MenuBackground]", "Themes", "Themes![PersonID] = " & LoggedInUserID)
    
    'Find the ContentSubFormBackground color
    FormContentColor = DLookup("[ContentSubFormBackground]", "Themes", "Themes![PersonID] = " & LoggedInUserID)
    
    'Find the text color
    TextColor = DLookup("[TextColor]", "Themes", "Themes![PersonID] = " & LoggedInUserID)
    
    'Get the Colors
    'Me.Detail.BackColor = GetSetting(cAppName, "Options", "Major BgColor", vRed)
    
    
    'SET THE COLORS
    '
    '
    'Set the bg color
    Detail.BackColor = BGColor
    'Me.lbl_PermLevel.ForeColor = TextColor

    
    
    
    'Display the user's real name in a welcome message
    'Me.lbl_Name.Caption = "Hello, " & PersonName & "!"
    Me.lbl_Name.Caption = PersonName & " " & surname
    Me.lbl_Name.ForeColor = TextColor


    
    'Icon Paths
    If Me.MenuButton.Value Then
        Me.MenuButton.Picture = GetImagePath & "Close.png"
        Me.MenuButton.QuickStyle = 0
    Else
        Me.MenuButton.Picture = GetImagePath & "Menu.png"
        Me.MenuButton.QuickStyle = 0
    End If
    
    Me.ImgDock.Picture = GetImagePath & "Dock75.png"
    Me.ImgDock.Width = 2 * 1440
    Me.ImgDock.Height = 2 * 1440
    
    Me.imgLogo.Picture = GetImagePath & "Logo.png"
    Me.imgLogo.Width = 0.75 * 1440
    Me.imgLogo.Height = 0.75 * 1440
    
End Sub

Private Sub MenuButton_Click()
    
    If Me.MenuButton.Value Then
        Me.MenuButton.Picture = GetImagePath & "Close.png"
    Else
        Me.MenuButton.Picture = GetImagePath & "Menu.png"
    End If
    
    Dim x As Integer
    Me.Menu.Visible = True
    Do
        DoEvents
        Me.Menu.Width = Me.Menu.Width + IIf(Me.MenuButton, 200, -200)
        Me.Menu.Left = Me.Menu.Left - IIf(Me.MenuButton, 200, -200)
        Me.MenuButton.Left = Me.MenuButton.Left - IIf(Me.MenuButton, 200, -200)
        timeout (0.01)
        x = x + 1
    Loop Until x = 10
    Me.Menu.Visible = Me.MenuButton
End Sub

Sub timeout(duration_ms As Double)
    Start_Time = Timer
    Do
    DoEvents
    Loop Until (Timer - Start_Time) >= duration_ms
End Sub

'Public Function GetImagePath() As String
'    GetImagePath = Application.CurrentProject.Path & "/Images/"
'End Function


'Private Sub Page53_Click()
'
'End Sub
'
'Private Sub TabCtl40_Change()
'    If Me.TabCtl40.Value = 4 Then
'        Me.TabCtl40.Value = 0
'        DoCmd.OpenForm ("frm_AddGame")
'    ElseIf Me.TabCtl40.Value = 3 Then
'        Me.TabCtl40.Value = 0
'        DoCmd.OpenForm ("frm_GamesList")
'    Else
'
'    End If
'End Sub
