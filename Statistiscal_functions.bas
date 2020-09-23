Attribute VB_Name = "Statistical_functions"
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''
' If anyone has a better solution to all of the  '
' modules below, please send it to me at         '
' foxdetective007@mailcity.com                   '
''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sigma As Double
Public N_scores As Integer
Public Arithmetic_Mean As Double
Public P_S_D As Double
Public S_S_D As Double
Public SigmaSq As Double

Public Sub Sum()
On Error GoTo solution
Sigma = S_R(1) + S_R(2) + S_R(3) + S_R(4) + S_R(5) + S_R(6) + _
    S_R(7) + S_R(8) + S_R(9) + S_R(10) + S_R(11) + S_R(12) + S_R(13) + _
    S_R(14) + S_R(15) + S_R(16) + S_R(17) + S_R(18) + S_R(19) + S_R(20)
    
frmMain.Number_space.Text = Sigma
Exit Sub
solution:
Call Misc.solution
End Sub
    
Public Sub Mean()
On Error GoTo solution
Sigma = S_R(1) + S_R(2) + S_R(3) + S_R(4) + S_R(5) + S_R(6) + _
    S_R(7) + S_R(8) + S_R(9) + S_R(10) + S_R(11) + S_R(12) + S_R(13) + _
    S_R(14) + S_R(15) + S_R(16) + S_R(17) + S_R(18) + S_R(19) + S_R(20)

Arithmetic_Mean = Sigma / Index

frmMain.Number_space.Text = Arithmetic_Mean
Exit Sub
solution:
Call Misc.solution
End Sub

Public Sub Population_standard_deviation()
On Error GoTo solution
SigmaSq = (S_R(1) * S_R(1)) + (S_R(2) * S_R(2)) + (S_R(3) * S_R(3)) + _
    (S_R(4) * S_R(4)) + (S_R(5) * S_R(5)) + (S_R(6) * S_R(6)) + (S_R(7) * _
    S_R(7)) + (S_R(8) * S_R(8)) + (S_R(9) * S_R(9)) + (S_R(10) * S_R(10)) + _
    (S_R(11) * S_R(11)) + (S_R(12) * S_R(12)) + (S_R(13) * S_R(13)) + _
    (S_R(14) * S_R(14)) + (S_R(15) * S_R(15)) + (S_R(16) * S_R(16)) + _
    (S_R(17) * S_R(17)) + (S_R(18) * S_R(18)) + (S_R(19) * S_R(19)) + _
    (S_R(20) * S_R(20))

Sigma = S_R(1) + S_R(2) + S_R(3) + S_R(4) + S_R(5) + S_R(6) + _
    S_R(7) + S_R(8) + S_R(9) + S_R(10) + S_R(11) + S_R(12) + S_R(13) + _
    S_R(14) + S_R(15) + S_R(16) + S_R(17) + S_R(18) + S_R(19) + S_R(20)
    
P_S_D = Sqr((SigmaSq - ((Sigma * Sigma) / Index)) / Index)

frmMain.Number_space.MaxLength = 15
frmMain.Number_space.Text = P_S_D
Exit Sub
solution: Call Misc.solution
End Sub

Public Sub Sample_standard_deviation()

On Error GoTo solution
SigmaSq = (S_R(1) * S_R(1)) + (S_R(2) * S_R(2)) + (S_R(3) * S_R(3)) + _
    (S_R(4) * S_R(4)) + (S_R(5) * S_R(5)) + (S_R(6) * S_R(6)) + (S_R(7) * _
    S_R(7)) + (S_R(8) * S_R(8)) + (S_R(9) * S_R(9)) + (S_R(10) * S_R(10)) + _
    (S_R(11) * S_R(11)) + (S_R(12) * S_R(12)) + (S_R(13) * S_R(13)) + _
    (S_R(14) * S_R(14)) + (S_R(15) * S_R(15)) + (S_R(16) * S_R(16)) + _
    (S_R(17) * S_R(17)) + (S_R(18) * S_R(18)) + (S_R(19) * S_R(19)) + _
    (S_R(20) * S_R(20))

Sigma = S_R(1) + S_R(2) + S_R(3) + S_R(4) + S_R(5) + S_R(6) + _
    S_R(7) + S_R(8) + S_R(9) + S_R(10) + S_R(11) + S_R(12) + S_R(13) + _
    S_R(14) + S_R(15) + S_R(16) + S_R(17) + S_R(18) + S_R(19) + S_R(20)
    
S_S_D = Sqr((SigmaSq - ((Sigma * Sigma) / Index)) / (Index - 1))

frmMain.Number_space.MaxLength = 15
frmMain.Number_space.Text = S_S_D
Exit Sub
solution: Call Misc.solution
End Sub

Public Sub Clear_s_r()
Dim Response As String
   On Error GoTo solution
   If frmMain.Statistic_score.Visible = True Then
        Response = MsgBox("Are you sure you want to clear the contents of the statistical register" _
        , vbExclamation + vbYesNo, "Clear the statistical register?")
        If Response = vbYes Then
            frmMain.Statistic_score.Text = "n"
            frmMain.Number_space.Text = "0"
            For Index = 1 To 20
                S_R(Index) = 0
            Next Index
            Index = 0
            Sigma = 0
            Dot = False
            N_scores = 0
            Arithmetic_Mean = 0
        ElseIf Response = vbNo Then
            frmMain.Refresh
        End If
    Else
    
    End If
    Exit Sub
solution:
    Call Misc.solution
End Sub


