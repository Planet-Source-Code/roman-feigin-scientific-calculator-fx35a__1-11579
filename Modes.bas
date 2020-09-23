Attribute VB_Name = "Modes_code"
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''
' If anyone has a better solution to all of the  '
' modules below, please send it to me at         '
' foxdetective007@mailcity.com                   '
''''''''''''''''''''''''''''''''''''''''''''''''''

Public S_R(1 To 20) As Double
Public Index As Integer
Public Mode_number As Integer

Public Sub Mode_Validation()

With frmMain

Select Case Mode_number
  
    Case "1"
        .Mode_type = "DEG"
        .Statistics_mode.Text = ""
        .Mode_indicator.Text = ""
        .Statistic_score.Text = ""
        Misc.Mode_value = False
        .Statistic_score.Visible = False
    Case "2"
        .Mode_type.Text = "DEG"
        .Statistics_mode.Text = ""
        .Mode_indicator.Text = ""
        Misc.Mode_value = False
        .Statistic_score.Visible = False
    Case "3"
        .Mode_type.Text = "RAD"
        .Statistics_mode.Text = ""
        .Mode_indicator.Text = ""
        Misc.Mode_value = False
        .Statistic_score.Visible = False
    Case "4"
        .Statistics_mode.Text = "SD"
        .Mode_indicator.Text = ""
        Misc.Mode_value = False
        Dot = False
        .Statistic_score.Visible = True
        .Statistic_score.Text = "n"
        Index = 0

End Select

End With

End Sub

