VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Scheduler"
   ClientHeight    =   780
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   4125
   LinkTopic       =   "Form1"
   ScaleHeight     =   780
   ScaleWidth      =   4125
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrMain 
      Interval        =   30000
      Left            =   4080
      Top             =   2640
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'By Mike Hartrick, All Rights Reserved 2001

Private Mon As Boolean
Private Tues As Boolean
Private Wed As Boolean
Private Thur As Boolean
Private Fri As Boolean
Private Sat As Boolean
Private Sun As Boolean

Private RunAgain As Boolean

Private BatFile As String
Private RunTime As String

Private Sub GetVars(FileSpec As String)
    Open FileSpec For Input As #1
        Input #1, m, t, w, th, f, s, su
            If m = 1 Then Mon = True Else Mon = False
            If t = 1 Then Tues = True Else Tued = False
            If w = 1 Then Wed = True Else Wed = False
            If th = 1 Then Thur = True Else Thur = False
            If f = 1 Then Fri = True Else Fri = False
            If s = 1 Then Sat = True Else Sat = False
            If su = 1 Then Sun = True Else Sun = False
        Input #1, BatFile
        Input #1, RunTime
    Close #1
    
End Sub


Private Sub LetsGo()
Dim thisday As String
    thisday = Format(Date, "dddd")

    If RunToday(GiveDay(thisday)) = True Then
        If RunAgain <> False Then
            i = Shell(BatFile)
            RunAgain = False
        Else
            RunAgain = True
        End If
    End If

End Sub

Private Sub DisplayData()
Dim thisday As String
    Cls
    Print "The Current Time Is: "; Time
    Print "The Current Date Is: "; Date
    Print "Execution is scheduled for: "; RunTime
    thisday = Format(Date, "dddd")
    If RunToday(GiveDay(thisday)) = True Then Print "Scheduled for today: "; BatFile
End Sub

Private Function GiveDay(Today As String) As Integer
    Select Case Today
        Case Is = "Monday"
            GiveDay = 1
        Case Is = "Tuesday"
            GiveDay = 2
        Case Is = "Wednesday"
            GiveDay = 3
        Case Is = "Thursday"
            GiveDay = 4
        Case Is = "Friday"
            GiveDay = 5
        Case Is = "Saturday"
            GiveDay = 6
        Case Is = "Sunday"
            GiveDay = 7
End Select
End Function

Private Function RunToday(Today As Integer) As Boolean
    Select Case Today
        Case Is = 1
            If Mon = True Then RunToday = True
        Case Is = 2
            If Tues = True Then RunToday = True
        Case Is = 3
            If Wed = True Then RunToday = True
        Case Is = 4
            If Thur = True Then RunToday = True
        Case Is = 5
            If Fri = True Then RunToday = True
        Case Is = 6
            If Sat = True Then RunToday = True
        Case Is = 7
            If Sun = True Then RunToday = True

    End Select
End Function

Private Sub Form_Load()
    Call GetVars(App.Path + "\days.ini")
    Call tmrMain_Timer
End Sub

Private Sub tmrMain_Timer()
Dim RTime As String
Dim CTime As String
Call DisplayData
    RTime = Format(RunTime, "hh:mm")
    CTime = Format(Time, "hh:mm")
    
    If RTime = CTime Then Call LetsGo

End Sub
