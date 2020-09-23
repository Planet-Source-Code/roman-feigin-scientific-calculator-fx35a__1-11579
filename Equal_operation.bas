Attribute VB_Name = "Equal_operation"
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''
' If anyone has a better solution to all of the  '
' modules below, please send it to me at         '
' foxdetective007@mailcity.com                   '
''''''''''''''''''''''''''''''''''''''''''''''''''

Public Exponential_value1 As String
Public Exponential_value2 As String
Public Operation As String
Public Power As Double
Public Loop_counter As Integer
Public Base As Double

Public Sub Equals()
Dim Exp_value As String

With frmMain
    If Exponent = False Then
        Misc.Value2 = .Number_space.Text
    ElseIf Exponent = True Then
        Exponential_value2 = .EXP.Text
        Misc.Value2 = .Number_space.Text * 10 ^ (Exponential_value2)
    End If

    Select Case Operation
        Case "+"
            Call Addition
        Case "-"
            Call Subtraction
        Case "×"
            Call Multiplication
        Case "÷"
            Call Division
        Case "N-RT"
            Call Misc.N_root
        Case "x^y"
            .Number_space.MaxLength = 18
            Power = CDbl(.Number_space.Text)
            Result = Base ^ Power
            .Number_space.Text = Result
    End Select
    
    If Exponent = False Then
        .Number_space.Text = CStr(Result)
    ElseIf Exponent = True Then
        Exp_value = Left(Right(CStr(Result), 4), 2)
        If Exp_value = "E+" Then
            .EXP.Text = Right(CStr(Result), 2)
            .Number_space.Text = Left(CStr(Result), Len(CStr(Result)) - 4)
        ElseIf Exp_value = "E-" Then
            .EXP.Text = -(Right(CStr(Result), 2))
            .Number_space.Text = Left(CStr(Result), Len(CStr(Result)) - 4)
        Else:
            .Number_space.Text = CStr(Result)
            .EXP.Visible = False
        End If
    End If
    
    Value1 = .Number_space.Text
    Sign = False
    First_digit = False
    Value2 = " "
    Result = 0

End With

End Sub

Public Sub Addition()
    
With frmMain
    If Exponent = False Then
        Result = CDbl(Value1) + CDbl(Value2)
    ElseIf Exponent = True Then
        If CDbl(Exponential_value1) + CDbl(Exponential_value2) <= 99 And _
        CDbl(Exponential_value1) + CDbl(Exponential_value2) >= -99 Then
            Result = CDbl(Value1) + CDbl(Value2)
        Else:
            .EXP.Text = ""
            .Number_space.Text = "-OVERFLOW ERROR-"
            MsgBox "The exponent must lie within the range -99 < e < 99.", _
            vbOKOnly + vbCritical, "Error"
            .Number_space.Text = "0"
        End If
    End If
End With

End Sub

Public Sub Subtraction()

With frmMain
    If Exponent = False Then
        Result = CDbl(Value1) - CDbl(Value2)
    ElseIf Exponent = True Then
        If CDbl(Exponential_value1) - CDbl(Exponential_value2) <= 99 And _
        CDbl(Exponential_value1) - CDbl(Exponential_value2) >= -99 Then
            Result = CDbl(Value1) - CDbl(Value2)
        Else:
            .EXP.Text = ""
            .Number_space.Text = "-OVERFLOW ERROR-"
            MsgBox "The exponent must lie within the range -99 < e < 99.", _
            vbOKOnly + vbCritical, "Error"
            .Number_space.Text = "0"
        End If
    End If
End With

End Sub

Public Sub Multiplication()
    
With frmMain
    If Exponent = False Then
        If .Percentage_indicator.Text = "%" Then
            Result = (CDbl(Value1) * CDbl(Value2)) / 100
        Else
            Result = CDbl(Value1) * CDbl(Value2)
        End If
    ElseIf Exponent = True Then
        If CDbl(Exponential_value1) + CDbl(Exponential_value2) <= 99 And _
        CDbl(Exponential_value1) + CDbl(Exponential_value2) >= -99 Then
            If .Percentage_indicator.Text = "%" Then
                Result = CDbl(Value1) * CDbl(Value2) / 100
            Else
                Result = CDbl(Value1) * CDbl(Value2)
            End If
        Else:
            .EXP.Text = ""
            .Number_space.Text = "-OVERFLOW ERROR-"
            MsgBox "The exponent must lie within the range -99 < e < 99.", _
            vbOKOnly + vbCritical, "Error"
            .Number_space.Text = "0"
        End If
    End If
End With

End Sub

Public Sub Division()

With frmMain
    If Exponent = False Then
        If .Percentage_indicator.Text = "%" Then
            If Value2 <> 0 Then
                Result = (CDbl(Value1) / CDbl(Value2)) * 100
            ElseIf Value2 = 0 Then
                .EXP.Text = ""
                .Percentage_indicator.Text = ""
                .Number_space.Text = "-ERROR-"
                MsgBox "You can't divide by zero", vbOKOnly + vbCritical, "Error"
                .Number_space.Text = "0"
                .Percentage_indicator.Text = " "
            End If
        Else
            If Value2 <> 0 Then
                Result = CDbl(Value1) / CDbl(Value2)
            ElseIf Value2 = 0 Then
                .EXP.Text = ""
                .Number_space.Text = "-ERROR-"
                MsgBox "You can't divide by zero", vbOKOnly + vbCritical, "Error"
                .Number_space.Text = "0"
            End If
        End If
    ElseIf Exponent = True Then
        If CDbl(Exponential_value1) - CDbl(Exponential_value2) <= 99 And _
        CDbl(Exponential_value1) - CDbl(Exponential_value2) >= -99 Then
            If .Percentage_indicator.Text = "%" Then
                If Value2 <> 0 Then
                    Result = (CDbl(Value1) / CDbl(Value2)) * 100
                ElseIf Value2 = 0 Then
                    .Percentage_indicator.Text = ""
                    .EXP.Text = ""
                    .Number_space.Text = "-ERROR-"
                    MsgBox "You can't divide by zero", vbOKOnly + vbCritical, "Error"
                    .Number_space.Text = "0"
                    .Percentage_indicator.Text = " "
                End If
            Else
                Result = CDbl(Value1) / CDbl(Value2)
            End If
        Else:
            .EXP.Text = ""
            .Number_space.Text = "-OVERFLOW ERROR-"
            MsgBox "The exponent must lie within the range -99 < e < 99.", _
            vbOKOnly + vbCritical, "Error"
            .Number_space.Text = "0"
        End If
    End If
End With
    Dot = False
End Sub
