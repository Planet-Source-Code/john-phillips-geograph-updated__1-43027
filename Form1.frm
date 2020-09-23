VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmGeoGraph 
   Caption         =   "GeoGraph Mapper"
   ClientHeight    =   6870
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10875
   LinkTopic       =   "Form1"
   MousePointer    =   1  'Arrow
   ScaleHeight     =   6870
   ScaleWidth      =   10875
   StartUpPosition =   2  'CenterScreen
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   0
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Mouse Tracking"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9120
      TabIndex        =   13
      Top             =   5640
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Caption         =   "Server Location"
      Height          =   1335
      Left            =   3000
      TabIndex        =   8
      Top             =   5520
      Width           =   6015
      Begin VB.CommandButton Command3 
         Caption         =   "View Info"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         TabIndex        =   15
         Top             =   840
         Width           =   1455
      End
      Begin VB.CheckBox chkIP 
         Caption         =   "Check2"
         Height          =   195
         Left            =   5640
         TabIndex        =   12
         Top             =   360
         Width           =   255
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   2760
         TabIndex        =   10
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Idle"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label Label3 
         Caption         =   "IP Address / Host Name:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Plot Map"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9120
      TabIndex        =   7
      Top             =   6000
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Lat / Long"
      Height          =   1335
      Left            =   0
      TabIndex        =   2
      Top             =   5520
      Width           =   2895
      Begin VB.CheckBox chkLonLat 
         Caption         =   "Check1"
         Height          =   195
         Left            =   2520
         TabIndex        =   11
         Top             =   240
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Longitude:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Latitude:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9120
      TabIndex        =   1
      Top             =   6480
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   5460
      Left            =   0
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   360
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   721
      TabIndex        =   0
      Top             =   0
      Width           =   10875
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Location:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   21
         Top             =   4920
         Width           =   975
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1320
         TabIndex        =   20
         Top             =   4920
         Width           =   5655
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   10080
         TabIndex        =   19
         Top             =   4920
         Width           =   735
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   8160
         TabIndex        =   18
         Top             =   4920
         Width           =   735
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Longitude:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9000
         TabIndex        =   17
         Top             =   4920
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Latitude:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7200
         TabIndex        =   16
         Top             =   4920
         Width           =   975
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         Height          =   135
         Left            =   1560
         Shape           =   3  'Circle
         Top             =   3600
         Width           =   75
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000000FF&
         X1              =   0
         X2              =   720
         Y1              =   248
         Y2              =   248
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         X1              =   104
         X2              =   104
         Y1              =   0
         Y2              =   360
      End
   End
End
Attribute VB_Name = "frmGeoGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bMouseTracking As Boolean
Dim xx As Integer
Dim yy As Integer

Private Sub Check4_Click()
If Check4.Value = 1 Then
    bMouseTracking = True
    Label5.Visible = True
    Label6.Visible = True
    Label7.Visible = True
    Label8.Visible = True
Else
    bMouseTracking = False
    Label5.Visible = False
    Label6.Visible = False
    Label7.Visible = False
    Label8.Visible = False
End If
End Sub

Private Sub chkIP_Click()
If chkIP = 1 Then
    chkLonLat = 0
End If
End Sub

Private Sub chkLonLat_Click()
If chkLonLat = 1 Then
    chkIP = 0
End If
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
' On Error GoTo errPlot
' plot in new york
' lat Pos 43.02663 : Long Neg -79.00138
Dim nLat As Integer
Dim nLong As Integer
Dim sTempString As String

Screen.MousePointer = vbHourglass

' do a check to see how the latitude and lonitude was
' entered, if it was entered as
' 40n and 79w then we need to convert to the appropriate
' value then submiot it to the plot map function
' else if they entere a positive or negative value
' then we can just submit it
' also in this part we round the number off
' since this map isn't so detailed we don't need to do
' any decimal point calculation to get the plot
' just the general area

If chkLonLat = 1 Then ' do plot by longitude latitude

' check to make sure the user entered values in the textbox's
If Text1.Text = "" Or Text2.Text = "" Then
    MsgBox "You must enter both Longitude and Latitude values!", vbOKOnly + vbInformation, "Error"
    Text1.SetFocus
    Screen.MousePointer = vbNormal
    Exit Sub
End If

    If Right(Text1.Text, 1) = "n" Or Right(Text1.Text, 1) = "N" Then
        nLat = RoundNum(Left(Text1.Text, Len(Text1.Text) - 1) * -1)
    ElseIf Right(Text1.Text, 1) = "s" Or Right(Text1.Text, 1) = "S" Then
        nLat = RoundNum(Left(Text1.Text, Len(Text1.Text) - 1))
    Else
        nLat = RoundNum(Text1.Text)
    End If
    
    If Right(Text2.Text, 1) = "w" Or Right(Text2.Text, 1) = "W" Then
        nLong = RoundNum(Left(Text2.Text, Len(Text2.Text) - 1) * -1)
    ElseIf Right(Text2.Text, 1) = "e" Or Right(Text2.Text, 1) = "E" Then
        nLong = RoundNum(Left(Text2.Text, Len(Text2.Text) - 1))
    Else
        nLong = RoundNum(Text2.Text)
    End If
    
    ' plot the lines on the map
    PlotMap nLong, nLat

ElseIf chkIP = 1 Then ' do serach and plot by IP address
' Domain servers in
' then
' </
' then
' href="

    Label4.Caption = "Searching..."
    sTempString = Inet1.OpenURL("http://www.nic.com/cgi-bin/whois.cgi?query=" & Trim(Text3.Text))

DoWebPause:
    Label4.Caption = "Getting IP Info"
    DoEvents
    
    Pause 3

If Inet1.StillExecuting = True Then GoTo DoWebPause

    PlotIP sTempString
    
Else

    MsgBox "There was an error processing your request!", vbOKOnly + vbInformation, "Error"
    
End If


Screen.MousePointer = vbNormal
Exit Sub
errPlot:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbOKOnly + vbCritical, "Error"
    Screen.MousePointer = vbNormal
    Exit Sub
    
End Sub

Private Sub Form_Load()
' turn the mouse tracking on
bMouseTracking = True

' set the crosshairs in the middle of the map
PlotMap 0, 0

End Sub

Private Sub Form_Resize()
If Me.WindowState = 1 Then Exit Sub

'Label5.Top = Picture1.Height - Label5.Height
'Label5.Left = Picture1.Width - Label5.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If bMouseTracking = True Then
Line1.X1 = x ' vertical line
Line1.X2 = x ' or latitude

Line2.Y1 = y ' horizontal line
Line2.Y2 = y ' or longitude

Shape1.Left = x - 2 ' crosshair center
Shape1.Top = y - 4  ' or lat and long

 xx = x
 yy = y

' have to work on this to get it to display
' the coords on the screen as the mouse moves

FindLatLong xx, yy

'label5.Top = y
'label5.Left = x

End If

End Sub
