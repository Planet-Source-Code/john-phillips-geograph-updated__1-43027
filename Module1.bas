Attribute VB_Name = "modMain"
Public DB As Database
Public RS As Recordset

Public Function PlotMap(lLat As Integer, lLong As Integer) As String
Dim x As Long ' long value
Dim y As Long ' long value
Dim lXHalf As Long ' half way point of the width of the map
Dim lYHalf As Long ' half way point of the height of the map
Dim l1XPoint As Long ' the value of what 1 deg on the x axis is equal to
Dim l1YPoint As Long ' the value of what 1 deg on the y axis is equal to
Dim lTemp As Long ' temporary value

' get the scale width and scale height values
x = frmGeoGraph.Picture1.ScaleWidth
y = frmGeoGraph.Picture1.ScaleHeight

' get the half way mark of the map in pixels
lXHalf = x / 2 ' get the 0 deg mark
lYHalf = y / 2 ' get the 0 deg mark

' determine what 1 deg on the map is equal to
l1XPoint = frmGeoGraph.Picture1.ScaleWidth / 360
l1YPoint = frmGeoGraph.Picture1.ScaleHeight / 180

' Latitude

If lLat < 0 Then ' negative value -180 to 0

' make the value positive - just easyier to work with
lLat = lLat * -1

    If lLat = 180 Then ' first point on map
        lLat = 1 ' just use 1 so the hair shows on the map
    Else
        lTemp = 180 - lLat
        
        lLat = lTemp * l1XPoint
    End If

ElseIf lLat = 0 Then
    ' half way point on map
    lLat = lXHalf
Else ' positive value 0 to 180
    
    lLat = (lLat * l1XPoint) + lXHalf
End If

' Longitude

If lLong < 0 Then ' negative value -90 to 0

' make the value positive
lLong = lLong * -1

    If lLong = 90 Then ' first point on map
        lLong = frmGeoGraph.Picture1.ScaleHeight - 1 ' just use -1 so the hair shows on the map
    Else
        lTemp = 90 + lLong
        
        lLong = lTemp * l1YPoint
    End If

ElseIf lLong = 0 Then
    lLong = lYHalf
Else
    lLong = (90 - lLong) * l1YPoint
    ' lLong = (lLong - lYHalf) * l1YPoint ') - lYHalf
End If

' draw the lines
frmGeoGraph.Line1.X1 = lLat
frmGeoGraph.Line1.X2 = lLat

frmGeoGraph.Line2.Y1 = lLong
frmGeoGraph.Line2.Y2 = lLong

frmGeoGraph.Shape1.Left = lLat - 2 ' crosshair center
frmGeoGraph.Shape1.Top = lLong - 4

PlotMap = lLat & "," & lLong

End Function

Public Function RoundNum(Number As Double) As Integer
' round off the number
' this was taken from PSC, I forgot who uploaded
' this function, but thank you whoever you are

    If Int(Number + 0.5) > Int(Number) Then
        RoundNum = Int(Number) + 1
    Else
        RoundNum = Int(Number)
    End If
    
End Function

Public Function FindLatLong(x As Integer, y As Integer) As String
On Error GoTo errPlot
Dim lXHalf As Long
Dim lYHalf As Long
Dim xPlot As String
Dim yPlot As String
Dim l1XPoint As Long
Dim l1YPoint As Long

DoEvents

lXHalf = frmGeoGraph.Picture1.ScaleWidth / 2
lYHalf = frmGeoGraph.Picture1.ScaleHeight / 2

l1XPoint = frmGeoGraph.Picture1.ScaleWidth / 360
l1YPoint = frmGeoGraph.Picture1.ScaleHeight / 180

If x = lXHalf Then ' half way point
    xPlot = "0"
ElseIf x < lXHalf Then ' negative value
    xPlot = x - lXHalf
    
    xPlot = xPlot / l1XPoint
Else
    xPlot = (x - lXHalf) / l1XPoint
End If

If y = lYHalf Then ' half way point
    yPlot = "0"
ElseIf y < lYHalf Then ' positive value
    yPlot = y - lYHalf
    
    yPlot = yPlot / l1YPoint
Else
    yPlot = (y - lYHalf) / l1YPoint
End If

frmGeoGraph.Label7 = xPlot
frmGeoGraph.Label8 = yPlot

FindLatLong = "Lat:" & xPlot & " - Long:" & yPlot

Exit Function
errPlot:
    FindLatLong = "Error..."
    Exit Function
End Function

Public Function PlotIP(sString As String) As String
Dim s As Long, e As Long, x As Long
Dim sHost As String
Dim sTemp As String
    
    ' fint the servers in the whois results
    s = InStr(1, sString, "Domain server", vbTextCompare)
    
    If s = 0 Then
        Do2ndLU ' not found error
        Exit Function
    End If
    ' fint the servers in the whois results
    e = InStr(s, sString, "</", vbTextCompare)
    
    If e = 0 Then GoTo errNotFound ' not found error
    
    ' fint the servers in the whois results
    s = InStr(e, sString, "href=", vbTextCompare)
    
    If s = 0 Then GoTo errNotFound
    
    ' s = InStr(s + 6, sString, "href=", vbTextCompare)
    
    ' If s = 0 Then GoTo errNotFound
    
    ' Get the IP address from the the whois results
    s = InStr(s + 6, sString, "query=", vbTextCompare)
    
    If s = 0 Then GoTo errNotFound ' not found error
    
    s = s + 6
    
    e = InStr(s, sString, """", vbTextCompare)
    
    If e = 0 Then GoTo errNotFound
    
    ' get the host whois information
    sHost = Mid(sString, s, e - s)
    
    If sHost = "" Then GoTo errNotFound
    
    frmGeoGraph.Label4.Caption = "Resolving Host IP"
    ' query the whois server for the host information

    sTemp = frmGeoGraph.Inet1.OpenURL("http://www.nic.com/cgi-bin/whois.cgi?query=" & sHost)

DoWebPause:

    Pause 3

' wait for the inet control to get all the info
If frmGeoGraph.Inet1.StillExecuting = True Then GoTo DoWebPause
    
' look for PostalCode:
    
    s = InStr(1, sTemp, "PostalCode:", vbTextCompare)
    
    If s = 0 Then GoTo errNotFound ' not found error
    
    s = s + 12
    
    ' get the first 5 digits of the postal code
    e = InStr(s, sTemp, "-", vbTextCompare)
    
    If e = 0 Then GoTo errNotFound ' not found error
    
Dim sZip As String
    
    ' get the zip code
    sZip = Mid(sTemp, s, e - s)
    
    If sZip = "" Then GoTo errNotFound ' not found error
    
    ' do the database lookup of the zip code
    
    frmGeoGraph.Label4.Caption = "Finding ZipCode Info"
    Set DB = OpenDatabase(App.Path & "\gps.mdb")
    
    Set RS = DB.OpenRecordset("SELECT * FROM zips WHERE zip='" & sZip & "'")
    
    If RS.EOF = True And RS.BOF = True Then
        ' no zipcode was found in the database
        ' lets try and match the closest one
        Dim iPlus As String, iMinu As String
        
        Set RS = DB.OpenRecordset("SELECT * FROM zips WHERE zip>'" & sZip & "' ORDER BY zip ASC")
        
        RS.MoveFirst
        
        iPlus = RS!zip
        
        Set RS = DB.OpenRecordset("SELECT * FROM zips WHERE zip<'" & sZip & "' ORDER BY zip ASC")
        
        RS.MoveLast
        
        iMinu = RS!zip
        If Len(sZip) > 5 Then
            sZip = Left(sZip, 5)
        End If
        
            If (sZip - iMinu) > (iPlus - sZip) Then
                ' use the higher zipcode
                Set RS = DB.OpenRecordset("SELECT * FROM zips WHERE zip='" & iPlus & "'")
            Else
                ' use the lower zip code
                Set RS = DB.OpenRecordset("SELECT * FROM zips WHERE zip='" & iMinu & "'")
            End If
    End If
    
    RS.MoveFirst
    
    ' plot and draw the coords on the map
    ' we use a negative lat value since the
    ' only zips we have in the database are
    ' of the united states
    
    PlotMap RoundNum("-" & RS!lat), RoundNum(RS!lon)
    frmGeoGraph.Label7.Caption = "-" & RS!lat
    frmGeoGraph.Label9.Caption = RS!town & ", " & RS!state
    frmGeoGraph.Label8.Caption = RS!lon
    frmGeoGraph.Label4.Caption = "Done!"
Exit Function

errNotFound:
'    MsgBox Err.Number & vbCrLf & Err.Description
    frmGeoGraph.Label4.Caption = "Error..."
    MsgBox "There was an error find the Longitude and Latitude Information for this server!", vbOKOnly + vbInformation, "Error"
    PlotMap 0, 0
    PlotIP = "0, 0"
    Exit Function

End Function

Public Sub Pause(PauseTime As Integer)
    StartTime = Timer

    Do While Timer - StartTime < PauseTime
        DoEvents
    Loop
End Sub

Private Function Do2ndLU() As String
Dim sTemp2 As String

frmGeoGraph.Label4.Caption = "Searching..."

    sTemp2 = frmGeoGraph.Inet1.OpenURL("http://ws.arin.net/cgi-bin/whois.pl?queryinput=" & Trim(frmGeoGraph.Text3.Text))

DoWebPause:

    Pause 3

' wait for the inet control to get all the info
If frmGeoGraph.Inet1.StillExecuting = True Then GoTo DoWebPause
    
' look for PostalCode:
    
    s = InStr(1, sTemp2, "PostalCode:", vbTextCompare)
    
    If s = 0 Then GoTo errNotFound ' not found error
    
    s = s + 12
    
        
Dim sZip As String
    
    ' get the zip code
    sZip = Mid(sTemp2, s, 5)
    
    If sZip = "" Then GoTo errNotFound ' not found error
    
    ' do the database lookup of the zip code
    
    frmGeoGraph.Label4.Caption = "Finding ZipCode Info"
    Set DB = OpenDatabase(App.Path & "\gps.mdb")
    
    Set RS = DB.OpenRecordset("SELECT * FROM zips WHERE zip='" & sZip & "'")
    
    If RS.EOF = True And RS.BOF = True Then
        ' no zipcode was found in the database
        ' lets try and match the closest one
        Dim iPlus As String, iMinu As String
        
        Set RS = DB.OpenRecordset("SELECT * FROM zips WHERE zip>'" & sZip & "' ORDER BY zip ASC")
        
        RS.MoveFirst
        
        iPlus = RS!zip
        
        Set RS = DB.OpenRecordset("SELECT * FROM zips WHERE zip<'" & sZip & "' ORDER BY zip ASC")
        
        RS.MoveLast
        
        iMinu = RS!zip
        If Len(sZip) > 5 Then
            sZip = Left(sZip, 5)
        End If
        
            If (sZip - iMinu) > (iPlus - sZip) Then
                ' use the higher zipcode
                Set RS = DB.OpenRecordset("SELECT * FROM zips WHERE zip='" & iPlus & "'")
            Else
                ' use the lower zip code
                Set RS = DB.OpenRecordset("SELECT * FROM zips WHERE zip='" & iMinu & "'")
            End If
    End If
    
    RS.MoveFirst
    
    ' plot and draw the coords on the map
    ' we use a negative lat value since the
    ' only zips we have in the database are
    ' of the united states
    
    PlotMap RoundNum("-" & RS!lat), RoundNum(RS!lon)
    frmGeoGraph.Label7.Caption = "-" & RS!lat
    frmGeoGraph.Label9.Caption = RS!town & ", " & RS!state
    frmGeoGraph.Label8.Caption = RS!lon
    frmGeoGraph.Label4.Caption = "Done!"
Exit Function

errNotFound:
'    MsgBox Err.Number & vbCrLf & Err.Description
    frmGeoGraph.Label4.Caption = "Error..."
    MsgBox "There was an error find the Longitude and Latitude Information for this server!", vbOKOnly + vbInformation, "Error"
    PlotMap 0, 0
    Do2ndLU = "0, 0"
    Exit Function



End Function
