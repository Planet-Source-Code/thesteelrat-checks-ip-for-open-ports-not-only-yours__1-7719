VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3960
   LinkTopic       =   "Form1"
   ScaleHeight     =   2445
   ScaleWidth      =   3960
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Text            =   "IP to check"
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   1560
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   3615
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1800
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Check"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Status 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "Status"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   2040
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
p = Text1.Text
If p = "" Then
    MsgBox "Please enter the valid ip"
Else
Text1.Enabled = False
    Status.Caption = "-"
        For i = "1" To "65530"
            Status.Caption = "-"
            Winsock1.Connect p, i
            Status.Caption = "Checking port: " & p & ":" & i
            Wait (2)
            If Status.Caption = "Found" Then
                List1.AddItem "Found: " & p & ":" & i
                Beep
                Status.Caption = ""
            End If
            Winsock1.Close
        Next i
End If
End Sub

Private Sub Command2_Click()
    Text1.Enabled = True
    Winsock1.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Winsock1.Close
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim TheData As String
    Winsock1.GetData TheData, vbString
End Sub

Private Sub winsock1_Connect()
    Status.Caption = "Found"
End Sub

Function Wait(numseconds As Long)
    Dim start As Variant, rightnow As Variant
    Dim HourDiff As Variant, MinuteDiff As Variant, SecondDiff As Variant
    Dim TotalMinDiff As Variant, TotalSecDiff As Variant
    start = Now
    While True
        rightnow = Now
        HourDiff = Hour(rightnow) - Hour(start)
        MinuteDiff = Minute(rightnow) - Minute(start)
        SecondDiff = Second(rightnow) - Second(start) + 1
        If SecondDiff = 60 Then
            MinuteDiff = MinuteDiff + 1 ' Add 1 to minute.
            SecondDiff = 0 ' Zero seconds.
        End If
        If MinuteDiff = 60 Then
            HourDiff = HourDiff + 1 ' Add 1 to hour.
            MinuteDiff = 0 ' Zero minutes.
        End If
        TotalMinDiff = (HourDiff * 60) + MinuteDiff ' Get totals.
        TotalSecDiff = (TotalMinDiff * 60) + SecondDiff
        If TotalSecDiff >= numseconds Then
            Exit Function
        End If
        DoEvents
            'Debug.Print rightnow
        Wend
End Function
