Attribute VB_Name = "Misc"
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sorry, I couldn't come up with a good solution for '
' scientific notation (see "Exponential_validation").'
' That's why this module has so much code. If anyone '
' has a better solution to this module, please send  '
' it to me at foxdetective007@mailcity.com           '
''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Exponent As Boolean
Public New_dot_position As Integer
Public Initial_length As Integer
Public Exponential_minus As Boolean
Public Minus As Boolean
Public Sign As Boolean
Public First_digit As Boolean
Public Value1, Value2 As String
Public Result As Double
Public Dot As Boolean
Public x As String
Public Loop_counter As Integer
Public Mode_value As Boolean
Global Memory_register As Double
Public Const E = 2.71828182846
Public Whole_portion As Integer

Public Sub Change_sign()
Dim Length As Integer

With frmMain

   Length = Len(.Number_space.Text) - 1 'program identifies "-" sign

    'If "0" is displayed do not print "-" sign

    If .Number_space.Text = "0" Then

    ElseIf .Number_space.Text <> "0" Then
            
        'If minus is allowed then print it, otherwise hide it

        If Minus = False Then
            .Number_space.Text = "-" & .Number_space.Text
            Minus = True
        ElseIf Minus = True Then
            .Number_space.Text = Right(.Number_space.Text, Length) 'program hides "-" sign
            Minus = False 'minus is allowed again
        End If
        
    End If
        
End With

End Sub

Public Sub Reset()
    
    With frmMain
        Format .Number_space.Text, "######0.##########"
        Sign = False
        First_digit = False
        .Number_space.Text = "0"
        .Function.Text = " "
        .Percentage_indicator.Text = " "
        Value1 = ""
        Value2 = ""
        .Memory_indicator.Text = " "
        Result = 0
        x = 0
        Loop_counter = 0
        Whole_portion = 0
        Dot = False  'decimal point is allowed
        Minus = False 'minus is allowed
        Exponential_minus = False
        Mode_value = False
        .Mode_indicator.Text = ""
        S_S_D = 0
        P_S_D = 0
        Arithmetic_Mean = 0
    End With
    
End Sub

Public Sub Square_root()
Dim Square_root As Double
Dim Argument As Double

With frmMain
    
    Argument = CDbl(.Number_space.Text)
    If Argument >= 0 Then
        Square_root = Sqr(CDbl(.Number_space.Text))
        .Number_space.Text = Square_root
    Else
        .Number_space.Text = "-ERROR-"
        .Function.Text = " "
        MsgBox "Square root of a negative number is not a real number", _
        vbCritical + vbOKOnly, "Error"
        .Number_space.Text = "0"
        Minus = False
    End If

End With

End Sub

Public Sub Reciprocal()
Dim Original_value As Double

With frmMain
    On Error GoTo solution
    Original_value = CDbl(.Number_space.Text)
    If Original_value <> 0 Then
        .Number_space.Text = 1 / Original_value
    Else
        .Number_space.Text = "-ERROR-"
        MsgBox "There is no reciprocal value of 0.", vbCritical + vbOKOnly, "Error"
        .Number_space.Text = "0"
    End If
    Exit Sub
solution:
    Call Misc.solution
End With

End Sub

Public Sub N_root()

With frmMain
    Power = CDbl(.Number_space.Text)
        If Power = 0 Then
            .Number_space.Text = "-ERROR-"
            MsgBox "Zero root of a number doesn't exist", vbCritical + vbOKOnly, _
            "Error"
            .Number_space.Text = "0"
        Else
            If Base < 0 And Power Mod 2 = 0 Then
                .Number_space.Text = "-ERROR-"
                MsgBox "nth root of a negative number when n is even, is not a real number", _
                vbCritical + vbOKOnly, "Error"
                .Number_space.Text = "0"
            ElseIf Base < 0 And Power Mod 2 <> 0 Then
                 Result = -(Abs(Base) ^ (1 / Power))
                .Number_space.Text = Result
            Else
                Result = Base ^ (1 / Power)
                .Number_space.Text = Result
            End If
        End If
End With

End Sub

Public Function LogN(x As Double, N As Double) As Double
   LogN = Log(x) / Log(N)
End Function

Public Function LgN(x As Double) As Double
   LgN = Log(x) / Log(E)
End Function

Public Sub solution()
    With frmMain
        .Number_space.Text = "-ERROR-"
        .EXP.Text = ""
        MsgBox "An error has occurred in the calculation. Clear all and try again.", _
        vbCritical + vbOKOnly, "Error"
        .Number_space.Text = "0"
    End With
End Sub

Public Sub Change_exponential_sign()
Dim Length As Integer

With frmMain

   Length = Len(.EXP.Text) - 1 'program identifies "-" sign

    'If "00" is displayed do not print "-" sign

    If .EXP.Text = "00" Then

    ElseIf .EXP.Text <> "00" Then
        'If exponential minus is allowed then print it, otherwise hide it
        
        If Exponential_minus = False Then
            .EXP.Text = "-" & .EXP.Text
            Exponential_minus = True
        ElseIf Exponential_minus = True Then
            .EXP.Text = Right(.EXP.Text, Length) 'program hides "-" sign
            Exponential_minus = False 'exponential minus is allowed again
        End If
        
    End If
        
End With

End Sub

Public Sub Exponential_validation()

With frmMain
    If Exponent = False Then
        Value1 = .Number_space.Text
        Value2 = " "
    ElseIf Exponent = True Then
        If Right(Left(Abs(.Number_space.Text), 2), 1) <> "." Then
            If Right(.Number_space.Text, 1) = "" Then
                'do nothing
            Else:
                If Misc.Whole_portion > 1 Then
                    If CDbl(.Number_space.Text) > 0 Then
                        .Number_space.Text = CDbl(.Number_space.Text) / 10 ^ (Whole_portion - 1)
                        New_dot_position = Misc.Whole_portion - 1
                        .EXP.Text = .EXP.Text + New_dot_position
                        Exponential_value1 = .EXP.Text
                        Value1 = .Number_space.Text * 10 ^ (Exponential_value1)
                    ElseIf CDbl(.Number_space.Text) < 0 Then
                        .Number_space.Text = -(Abs(CDbl(.Number_space.Text))) / 10 ^ (Whole_portion - 1)
                        New_dot_position = Misc.Whole_portion - 1
                        .EXP.Text = .EXP.Text + New_dot_position
                        Exponential_value1 = .EXP.Text
                        Value1 = .Number_space.Text * 10 ^ (Exponential_value1)
                    End If
                ElseIf Misc.Whole_portion < 1 Then
                    If CDbl(.Number_space.Text) > 0 Then
                        Initial_length = Len(.Number_space.Text)
                    ElseIf CDbl(.Number_space.Text) < 0 Then
                        Initial_length = Len(.Number_space.Text) - 1
                    End If
                    .Number_space.Text = CDbl(.Number_space.Text) / 10 ^ (Initial_length - 1)
                    .EXP.Text = .EXP.Text + Initial_length - 1
                    Exponential_value1 = .EXP.Text
                    Value1 = .Number_space.Text * 10 ^ (Exponential_value1)
                    If Abs(Exponential_value1) > 99 Then
                        .EXP.Text = ""
                        .Number_space.Text = "-ERROR-"
                        MsgBox "The exponent value must lie within the range from -99 to 99 inclusive.", _
                        vbOKOnly + vbCritical, "Error"
                        .Number_space.Text = "0"
                    End If
                Else
                    Exponential_value1 = .EXP.Text
                    Value1 = .Number_space.Text * 10 ^ (Exponential_value1)
                    Value2 = " "
                End If
            End If
        Else:
            If Right(Left(Abs(.Number_space.Text), 2), 2) = "0." Then
                Do
                    x = Right(Left(Abs(.Number_space.Text), 1), Loop_counter)
                    Loop_counter = Loop_counter + 1
                Loop Until x <> "0" And x <> "."
               
                    .Number_space.Text = CDbl(.Number_space.Text) * 10 ^ (Loop_counter)
                    If CDbl(.Number_space.Text) > 0 Then
                        If CDbl(.EXP.Text) > 0 Then
                            .EXP.Text = .EXP.Text - (Loop_counter)
                        ElseIf CDbl(.EXP.Text) < 0 Then
                            .EXP.Text = -(Abs(.EXP.Text) + Loop_counter)
                        End If
                    ElseIf CDbl(.Number_space.Text) < 0 Then
                        If CDbl(.EXP.Text) > 0 Then
                            .EXP.Text = .EXP.Text - Loop_counter
                        ElseIf CDbl(.EXP.Text) < 0 Then
                            .EXP.Text = -(Abs(.EXP.Text) + Loop_counter)
                        End If
                    End If
                        Exponential_value1 = .EXP.Text
                        Value1 = .Number_space.Text * 10 ^ (Exponential_value1)
                
                If Abs(Exponential_value1) > 99 Then
                    .EXP.Text = ""
                    .Number_space.Text = "-ERROR-"
                    MsgBox "The exponent value must lie within the range from -99 to 99 inclusive.", _
                    , vbOKOnly + vbCritical, "Error"
                    .Number_space.Text = "0"
                End If
            Else
            
            End If
        End If
        Exponent = False
    End If
End With

End Sub

