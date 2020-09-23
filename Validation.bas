Attribute VB_Name = "Trig_functions"
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''
' If anyone has a better solution to all of the  '
' modules below, please send it to me at         '
' foxdetective007@mailcity.com                   '
''''''''''''''''''''''''''''''''''''''''''''''''''

Public Const PI = 3.141592654
Public Trig_operation As Double
Public Angle As Double
Public Minus As Boolean

Public Function Cosec(Angle As Double) As Double
    Cosec = 1 / (Sin(Angle))
End Function

Public Function Sec(Angle As Double) As Double
    Sec = 1 / (Cos(Angle))
End Function

Public Function Cot(Angle As Double) As Double
    Cot = 1 / (Tan(Angle))
End Function

Public Sub Sine_validation()

With frmMain

Angle = CDbl(.Number_space.Text) Mod 360

Select Case Abs(Angle)

    Case "30"
        If .Mode_type = "DEG" Then
            Trig_operation = Sin(Angle * PI / 180)
            .Number_space.Text = Format(Trig_operation, "0.#")
        ElseIf frmMain.Mode_type.Text = "RAD" Then
            Trig_operation = Sin(Angle)
            .Number_space.Text = Trig_operation
        End If
    Case "150"
        If .Mode_type.Text = "DEG" Then
            Trig_operation = Sin(Angle * PI / 180)
            .Number_space.Text = Format(Trig_operation, "0.#")
        ElseIf .Mode_type.Text = "RAD" Then
            Trig_operation = Sin(Angle)
            .Number_space.Text = Trig_operation
        End If
    Case "210"
        If .Mode_type.Text = "DEG" Then
            Trig_operation = Sin(Angle * PI / 180)
            .Number_space.Text = Format(Trig_operation, "0.#")
        ElseIf .Mode_type.Text = "RAD" Then
            Trig_operation = Sin(Angle)
            .Number_space.Text = Trig_operation
        End If
    Case "330"
        If .Mode_type.Text = "DEG" Then
            Trig_operation = Sin(Angle * PI / 180)
            .Number_space.Text = Format(Trig_operation, "0.#")
        ElseIf .Mode_type.Text = "RAD" Then
            Trig_operation = Sin(Angle)
            .Number_space.Text = Trig_operation
        End If
    Case "0"
        If .Mode_type.Text = "DEG" Then
            Trig_operation = Sin(Angle * PI / 180)
            .Number_space.Text = Format(Trig_operation, "0")
        ElseIf .Mode_type.Text = "RAD" Then
            Trig_operation = Sin(Angle)
            .Number_space.Text = Trig_operation
        End If
    Case "180"
        If .Mode_type.Text = "DEG" Then
            Trig_operation = Sin(Angle * PI / 180)
            .Number_space.Text = Format(Trig_operation, "0")
        ElseIf .Mode_type.Text = "RAD" Then
            Trig_operation = Sin(Angle)
            .Number_space.Text = Trig_operation
        End If
    Case "360"
        If .Mode_type.Text = "DEG" Then
            Trig_operation = Sin(Angle * PI / 180)
            .Number_space.Text = Format(Trig_operation, "0")
        ElseIf .Mode_type.Enabled = "RAD" Then
            Trig_operation = Sin(Angle)
            .Number_space.Text = Trig_operation
        End If
    Case "90"
        If .Mode_type.Text = "DEG" Then
            Trig_operation = Sin(Angle * PI / 180)
            .Number_space.Text = Format(Trig_operation, "#")
        ElseIf .Mode_type.Text = "RAD" Then
            Trig_operation = Sin(Angle)
            .Number_space.Text = Trig_operation
        End If
    Case "270"
        If .Mode_type.Text = "DEG" Then
            Trig_operation = Sin(Angle * PI / 180)
            .Number_space.Text = Format(Trig_operation, "#")
        ElseIf .Mode_type.Text = "RAD" Then
            Trig_operation = Sin(Angle)
            .Number_space.Text = Trig_operation
        End If
    Case Else
        If .Mode_type.Text = "DEG" Then
            Trig_operation = Sin(Angle * PI / 180)
            .Number_space.Text = Format(Trig_operation, "0.############")
        ElseIf .Mode_type.Text = "RAD" Then
            Trig_operation = Sin(Angle)
            .Number_space.Text = Trig_operation
        End If
End Select

End With

End Sub

Public Sub Cosine_validation()

With frmMain

Angle = CDbl(.Number_space.Text) Mod 360

Select Case Abs(Angle)

    Case "60"
        If .Mode_type.Text = "DEG" Then
            Trig_operation = Cos(Angle * PI / 180)
            .Number_space.Text = Format(Trig_operation, "0.#")
        ElseIf .Mode_type.Text = "RAD" Then
            Trig_operation = Cos(Angle)
            .Number_space.Text = Trig_operation
        End If
    Case "120"
        If .Mode_type.Text = "DEG" Then
            Trig_operation = Cos(Angle * PI / 180)
            .Number_space.Text = Format(Trig_operation, "0.#")
        ElseIf .Mode_type.Text = "RAD" Then
            Trig_operation = Cos(Angle)
            .Number_space.Text = Trig_operation
        End If
    Case "240"
        If .Mode_type.Text = "DEG" Then
            Trig_operation = Cos(Angle * PI / 180)
            .Number_space.Text = Format(Trig_operation, "0.#")
        ElseIf .Mode_type.Text = "RAD" Then
            Trig_operation = Cos(Angle)
            .Number_space.Text = Trig_operation
        End If
    Case "300"
        If .Mode_type.Text = "DEG" Then
            Trig_operation = Cos(Angle * PI / 180)
            .Number_space.Text = Format(Trig_operation, "0.#")
        ElseIf .Mode_type.Text = "RAD" Then
            Trig_operation = Cos(Angle)
            .Number_space.Text = Trig_operation
        End If
    Case "0"
        If .Mode_type.Text = "DEG" Then
            Trig_operation = Cos(Angle * PI / 180)
            .Number_space.Text = Format(Trig_operation, "#")
        ElseIf .Mode_type.Text = "RAD" Then
            Trig_operation = Cos(Angle)
            .Number_space.Text = Trig_operation
        End If
    Case "180"
        If .Mode_type.Text = "DEG" Then
            Trig_operation = Cos(Angle * PI / 180)
            .Number_space.Text = Format(Trig_operation, "#")
        ElseIf .Mode_type.Text = "RAD" Then
             Trig_operation = Cos(Angle)
            .Number_space.Text = Trig_operation
       End If
    Case "360"
        If .Mode_type.Text = "DEG" Then
            Trig_operation = Cos(Angle * PI / 180)
            .Number_space.Text = Format(Trig_operation, "#")
        ElseIf .Mode_type.Text = "RAD" Then
            Trig_operation = Cos(Angle)
            .Number_space.Text = Trig_operation
        End If
    Case "90"
        If .Mode_type.Text = "DEG" Then
            Trig_operation = Cos(Angle * PI / 180)
            .Number_space.Text = Format(Trig_operation, "0")
        ElseIf .Mode_type.Text = "RAD" Then
            Trig_operation = Cos(Angle)
            .Number_space.Text = Trig_operation
        End If
    Case "270"
        If .Mode_type.Text = "DEG" Then
            Trig_operation = Cos(Angle * PI / 180)
            .Number_space.Text = Format(Trig_operation, "0")
        ElseIf .Mode_type.Text = "RAD" Then
            Trig_operation = Cos(Angle)
            .Number_space.Text = Trig_operation
        End If
    Case Else
        If .Mode_type.Text = "DEG" Then
            Trig_operation = Cos(Angle * PI / 180)
            .Number_space.Text = Format(Trig_operation, "0.############")
        ElseIf .Mode_type.Text = "RAD" Then
            Trig_operation = Cos(Angle)
            .Number_space.Text = Trig_operation
        End If
End Select

End With

End Sub

Public Sub Tangent_validation()

Dim Actual_angle As String

With frmMain
       
Actual_angle = .Number_space.Text
Angle = CDbl(.Number_space.Text) Mod 360

Select Case Abs(Angle)
        
    Case "45"
        If .Mode_type.Text = "DEG" Then
            Trig_operation = Tan(Angle * PI / 180)
            .Number_space.Text = Format(Trig_operation, "#")
        ElseIf .Mode_type.Text = "RAD" Then
            Trig_operation = Tan(Angle)
            .Number_space.Text = Trig_operation
        End If
    Case "135"
        If .Mode_type.Text = "DEG" Then
            Trig_operation = Tan(Angle * PI / 180)
            .Number_space.Text = Format(Trig_operation, "#")
        ElseIf .Mode_type.Text = "RAD" Then
            Trig_operation = Tan(Angle)
            .Number_space.Text = Trig_operation
        End If
    Case "225"
        If .Mode_type.Text = "DEG" Then
            Trig_operation = Tan(Angle * PI / 180)
            .Number_space.Text = Format(Trig_operation, "#")
        ElseIf .Mode_type.Text = "RAD" Then
            Trig_operation = Tan(Angle)
            .Number_space.Text = Trig_operation
        End If
    Case "315"
        If .Mode_type.Text = "DEG" Then
            Trig_operation = Tan(Angle * PI / 180)
            .Number_space.Text = Format(Trig_operation, "#")
        ElseIf .Mode_type.Text = "RAD" Then
            Trig_operation = Tan(Angle)
            .Number_space.Text = Trig_operation
        End If
    Case "0"
        If .Mode_type.Text = "DEG" Then
            Trig_operation = Tan(Angle * PI / 180)
            .Number_space.Text = Format(Trig_operation, "0")
        ElseIf .Mode_type.Text = "RAD" Then
            Trig_operation = Tan(Angle)
            .Number_space.Text = Trig_operation
        End If
    Case "180"
        If .Mode_type.Text = "DEG" Then
            Trig_operation = Tan(Angle * PI / 180)
            .Number_space.Text = Format(Trig_operation, "0")
        ElseIf .Mode_type.Text = "RAD" Then
            Trig_operation = Tan(Angle)
            .Number_space.Text = Trig_operation
        End If
    Case "360"
        If .Mode_type.Text = "DEG" Then
            Trig_operation = Tan(Angle * PI / 180)
            .Number_space.Text = Format(Trig_operation, "0")
        ElseIf .Mode_type.Text = "RAD" Then
            Trig_operation = Tan(Angle)
            .Number_space.Text = Trig_operation
        End If
    Case "90"
        If .Mode_type.Text = "DEG" Then
            .Number_space.Text = "-ERROR-"
            MsgBox "The value of tan " & Actual_angle & "° is undefined for degrees." _
            , vbCritical + vbOKOnly, "Error"
            .Number_space.Text = "0"
            Minus = False
            .Function.Text = " "
            Format Trig_operation, "######0.##########"
        ElseIf .Mode_type.Text = "RAD" Then
            Trig_operation = Tan(Angle)
            .Number_space.Text = Trig_operation
        End If
    Case "270"
        If .Mode_type.Text = "DEG" Then
            .Number_space.Text = "-ERROR-"
            MsgBox "The value of tan " & Actual_angle & "° is undefined for degrees.", _
            vbCritical + vbOKOnly, "Error"
            .Number_space.Text = "0"
            Minus = False
            .Function.Text = " "
        ElseIf .Mode_type.Text = "RAD" Then
             Trig_operation = Tan(Angle)
            .Number_space.Text = Trig_operation
        End If
    Case Else
        If .Mode_type.Text = "DEG" Then
            Trig_operation = Tan(Angle * PI / 180)
            .Number_space.Text = Format(Trig_operation, "######0.##########")
        ElseIf .Mode_type.Text = "RAD" Then
            Trig_operation = Tan(Angle)
            .Number_space.Text = Trig_operation
        End If
End Select

End With

End Sub

Public Sub Cosecant_validation()
Dim Actual_angle As String

With frmMain

Actual_angle = .Number_space.Text

Angle = CDbl(.Number_space.Text) Mod 360

Select Case Abs(Angle)

    Case "30"
        If .Mode_type.Text = "DEG" Then
            Trig_operation = Cosec(Angle * PI / 180)
            .Number_space.Text = Format(Trig_operation, "#")
        ElseIf .Mode_type.Text = "RAD" Then
            Trig_operation = Cosec(Angle)
            .Number_space.Text = Trig_operation
        End If
    Case "150"
        If .Mode_type.Text = "DEG" Then
            Trig_operation = Cosec(Angle * PI / 180)
            .Number_space.Text = Format(Trig_operation, "#")
        ElseIf .Mode_type.Text = "RAD" Then
            Trig_operation = Cosec(Angle)
            .Number_space.Text = Trig_operation
        End If
    Case "210"
        If .Mode_type.Text = "DEG" Then
            Trig_operation = Cosec(Angle * PI / 180)
            .Number_space.Text = Format(Trig_operation, "#")
        ElseIf .Mode_type.Text = "RAD" Then
            Trig_operation = Cosec(Angle)
            .Number_space.Text = Trig_operation
        End If
    Case "330"
        If .Mode_type.Text = "DEG" Then
            Trig_operation = Cosec(Angle * PI / 180)
            .Number_space.Text = Format(Trig_operation, "#")
        ElseIf .Mode_type.Text = "RAD" Then
            Trig_operation = Cosec(Angle)
            .Number_space.Text = Trig_operation
        End If
    Case "0"
        If .Mode_type.Text = "DEG" Then
            .Number_space.Text = "-ERROR-"
            MsgBox "The value of cosec " & Actual_angle & "° is undefined for degrees." _
            , vbCritical + vbOKOnly, "Error"
            .Number_space.Text = "0"
            .Function.Text = " "
            Minus = False
        ElseIf .Mode_type.Text = "RAD" Then
            If Actual_angle = "0" Then
                .Number_space.Text = "-ERROR-"
                MsgBox "The value of cosec " & Actual_angle & "° is undefined for radians." _
                , vbCritical + vbOKOnly, "Error"
                .Number_space.Text = "0"
                .Function.Text = ""
            ElseIf Actual_angle Mod 360 >= 0 Then
                Trig_operation = Cosec(CDbl(.Number_space.Text))
                .Number_space.Text = Trig_operation
            End If
        End If
    Case "180"
        If .Mode_type.Text = "DEG" Then
            .Number_space.Text = "-ERROR-"
            MsgBox "The value of cosec " & Actual_angle & "° is undefined for degrees", vbCritical + vbOKOnly, "Error"
            .Number_space.Text = "0"
            .Function.Text = " "
            Minus = False
        ElseIf .Mode_type.Text = "RAD" Then
            Trig_operation = Cosec(Angle)
            .Number_space.Text = Trig_operation
        End If
    Case "360"
        If .Mode_type.Text = "DEG" Then
            .Number_space.Text = "-ERROR-"
            MsgBox "The value of cosec " & Actual_angle & "° is undefined for degrees." _
            , vbCritical + vbOKOnly, "Error"
            .Number_space.Text = "0"
            .Function.Text = " "
            Minus = False
        ElseIf .Mode_type.Text = "RAD" Then
            Trig_operation = Cosec(Angle)
            .Number_space.Text = Trig_operation
        End If
    Case "90"
        If .Mode_type.Text = "DEG" Then
            Trig_operation = Cosec(Angle * PI / 180)
            .Number_space.Text = Format(Trig_operation, "#")
        ElseIf .Mode_type.Text = "RAD" Then
            Trig_operation = Cosec(Angle)
            .Number_space.Text = Trig_operation
        End If
    Case "270"
        If .Mode_type.Text = "DEG" Then
            Trig_operation = Cosec(Angle * PI / 180)
            .Number_space.Text = Format(Trig_operation, "#")
        ElseIf .Mode_type.Text = "RAD" Then
            Trig_operation = Cosec(Angle)
            .Number_space.Text = Trig_operation
        End If
    Case Else
        If .Mode_type.Text = "DEG" Then
            Trig_operation = Cosec(Angle * PI / 180)
            .Number_space.Text = Trig_operation
        ElseIf .Mode_type.Text = "RAD" Then
            Trig_operation = Cosec(Angle)
            .Number_space.Text = Trig_operation
        End If
End Select

End With

End Sub

Public Sub Secant_validation()
Dim Actual_angle As String

With frmMain

Actual_angle = .Number_space.Text
Angle = CDbl(.Number_space.Text) Mod 360

Select Case Abs(Angle)

    Case "60"
        If .Mode_type.Text = "DEG" Then
            Trig_operation = Sec(CDbl(.Number_space.Text) * PI / 180)
            .Number_space.Text = Format(Trig_operation, "#")
        ElseIf .Mode_type.Text = "RAD" Then
            Trig_operation = Sec(CDbl(.Number_space.Text))
            .Number_space.Text = Trig_operation
        End If
    Case "120"
        If .Mode_type.Text = "DEG" Then
            Trig_operation = Sec(CDbl(.Number_space.Text) * PI / 180)
            .Number_space.Text = Format(Trig_operation, "#")
        ElseIf .Mode_type.Text = "RAD" Then
            Trig_operation = Sec(CDbl(.Number_space.Text))
            .Number_space.Text = Trig_operation
        End If
    Case "240"
        If .Mode_type.Text = "DEG" Then
            Trig_operation = Sec(CDbl(.Number_space.Text) * PI / 180)
            .Number_space.Text = Format(Trig_operation, "#")
        ElseIf .Mode_type.Text = "RAD" Then
            Trig_operation = Sec(CDbl(.Number_space.Text))
            .Number_space.Text = Trig_operation
        End If
    Case "300"
        If .Mode_type.Text = "DEG" Then
            Trig_operation = Sec(CDbl(.Number_space.Text) * PI / 180)
            .Number_space.Text = Format(Trig_operation, "#")
        ElseIf .Mode_type.Text = "RAD" Then
            Trig_operation = Sec(CDbl(.Number_space.Text))
            .Number_space.Text = Trig_operation
        End If
    Case "0"
        If .Mode_type.Text = "DEG" Then
            Trig_operation = Sec(CDbl(.Number_space.Text) * PI / 180)
            .Number_space.Text = Format(Trig_operation, "#")
        ElseIf .Mode_type.Text = "RAD" Then
            Trig_operation = Sec(CDbl(.Number_space.Text))
            .Number_space.Text = Trig_operation
        End If
    Case "180"
        If .Mode_type.Text = "DEG" Then
            Trig_operation = Sec(CDbl(.Number_space.Text) * PI / 180)
            .Number_space.Text = Format(Trig_operation, "#")
        ElseIf .Mode_type.Text = "RAD" Then
            Trig_operation = Sec(CDbl(.Number_space.Text))
            .Number_space.Text = Trig_operation
        End If
    Case "360"
        If .Mode_type.Text = "DEG" Then
            Trig_operation = Sec(CDbl(.Number_space.Text) * PI / 180)
            .Number_space.Text = Format(Trig_operation, "#")
        ElseIf .Mode_type.Text = "RAD" Then
            Trig_operation = Sec(CDbl(.Number_space.Text))
            .Number_space.Text = Trig_operation
        End If
    Case "90"
        If .Mode_type.Text = "DEG" Then
            .Number_space.Text = "-ERROR-"
            MsgBox "The value of sec " & Actual_angle & " ° is undefined.", vbCritical + vbOKOnly, "Error"
            .Number_space.Text = "0"
            .Function.Text = " "
            Minus = False
        ElseIf .Mode_type.Text = "RAD" Then
            Trig_operation = Sec(CDbl(.Number_space.Text))
            .Number_space.Text = Trig_operation
        End If
    Case "270"
        If .Mode_type.Text = "DEG" Then
            .Number_space.Text = "-ERROR-"
            MsgBox "The value of sec " & Actual_angle & "° is undefined.", vbCritical + vbOKOnly, "Error"
            .Number_space.Text = "0"
            .Function.Text = " "
            Minus = False
        ElseIf .Mode_type.Text = "RAD" Then
            Trig_operation = Sec(CDbl(.Number_space.Text))
            .Number_space.Text = Trig_operation
        End If
    Case Else
        If .Mode_type.Text = "DEG" Then
            Trig_operation = Sec(CDbl(.Number_space.Text) * PI / 180)
            .Number_space.Text = Trig_operation
        ElseIf .Mode_type.Text = "RAD" Then
            Trig_operation = Sec(CDbl(.Number_space.Text))
            .Number_space.Text = Trig_operation
        End If
End Select

End With

End Sub

Public Sub Cotangent_validation()
Dim Actual_angle As String

With frmMain

Actual_angle = .Number_space.Text
Angle = CDbl(.Number_space.Text) Mod 360

Select Case Abs(Angle)

    Case "45"
        If .Mode_type.Text = "DEG" Then
            Trig_operation = Cot(CDbl(.Number_space.Text) * PI / 180)
            .Number_space.Text = Format(Trig_operation, "#")
        ElseIf .Mode_type.Text = "RAD" Then
            Trig_operation = Cot(CDbl(.Number_space.Text))
            .Number_space.Text = Trig_operation
        End If
    Case "135"
        If .Mode_type.Text = "DEG" Then
            Trig_operation = Cot(CDbl(.Number_space.Text) * PI / 180)
            .Number_space.Text = Format(Trig_operation, "#")
        ElseIf .Mode_type.Text = "RAD" Then
            Trig_operation = Cot(CDbl(.Number_space.Text))
            .Number_space.Text = Trig_operation
        End If
    Case "225"
        If .Mode_type.Text = "DEG" Then
            Trig_operation = Cot(CDbl(.Number_space.Text) * PI / 180)
            .Number_space.Text = Format(Trig_operation, "#")
        ElseIf .Mode_type.Text = "RAD" Then
            Trig_operation = Cot(CDbl(.Number_space.Text))
            .Number_space.Text = Trig_operation
        End If
    Case "315"
        If .Mode_type.Text = "DEG" Then
            Trig_operation = Cot(CDbl(.Number_space.Text) * PI / 180)
            .Number_space.Text = Format(Trig_operation, "#")
        ElseIf .Mode_type.Text = "RAD" Then
            Trig_operation = Cot(CDbl(.Number_space.Text))
            .Number_space.Text = Trig_operation
        End If
    Case "0"
        If .Mode_type.Text = "DEG" Then
            .Number_space.Text = "-ERROR-"
            MsgBox "The value of cot " & Actual_angle & "° is undefined for degrees." _
            , vbCritical + vbOKOnly, "Error"
            .Number_space.Text = "0"
            .Function.Text = " "
            Minus = False
        ElseIf .Mode_type.Text = "RAD" Then
            If Actual_angle = "0" Then
                .Number_space.Text = "-ERROR-"
                MsgBox "The value of cot " & Actual_angle & "° is undefined for radians." _
                , vbCritical + vbOKOnly, "Error"
                .Number_space.Text = "0"
                .Function.Text = ""
            ElseIf Actual_angle Mod 360 >= 0 Then
                Trig_operation = Cot(CDbl(.Number_space.Text))
                .Number_space.Text = Trig_operation
            End If
        End If
    Case "180"
        If .Mode_type.Text = "DEG" Then
            .Number_space.Text = "-ERROR-"
            MsgBox "The value of cot " & Actual_angle & "° is undefined for degrees.", vbCritical + vbOKOnly, "Error"
            .Number_space.Text = "0"
            .Function.Text = " "
            Minus = False
        ElseIf .Mode_type.Text = "RAD" Then
            Trig_operation = Cot(CDbl(.Number_space.Text))
            .Number_space.Text = Trig_operation
        End If
    Case "360"
        If .Mode_type.Text = "DEG" Then
            .Number_space.Text = "-ERROR-"
            MsgBox "The value of cot " & Actual_angle & "° is undefined for degrees.", vbCritical + vbOKOnly, "Error"
            .Number_space.Text = "0"
            .Function.Text = " "
            Minus = False
        ElseIf .Mode_type.Text = "RAD" Then
            Trig_operation = Cot(CDbl(.Number_space.Text))
            .Number_space.Text = Trig_operation
        End If
    Case "90"
        If .Mode_type.Text = "DEG" Then
            Trig_operation = Cot(CDbl(.Number_space.Text) * PI / 180)
            .Number_space.Text = Format(Trig_operation, "0")
        ElseIf .Mode_type.Text = "RAD" Then
            Trig_operation = Cot(CDbl(.Number_space.Text))
            .Number_space.Text = Trig_operation
        End If
    Case "270"
        If .Mode_type.Text = "DEG" Then
            Trig_operation = Cot(CDbl(.Number_space.Text) * PI / 180)
            .Number_space.Text = Format(Trig_operation, "0")
        ElseIf .Mode_type.Text = "RAD" Then
            Trig_operation = Cot(CDbl(.Number_space.Text))
            .Number_space.Text = Trig_operation
        End If
    Case Else
        If .Mode_type.Text = "DEG" Then
            Trig_operation = Cot(CDbl(.Number_space.Text) * PI / 180)
            .Number_space.Text = Format(Trig_operation, "############")
        ElseIf .Mode_type.Text = "RAD" Then
            Trig_operation = Cot(CDbl(.Number_space.Text))
            .Number_space.Text = Trig_operation
        End If

End Select

End With
End Sub

Public Sub Arctangent()
Dim Tangent As String

With frmMain

    Tangent = CDbl(frmMain.Number_space.Text)
    
    Select Case Tangent
        
        Case 0
            If .Mode_type.Text = "DEG" Then
                Trig_operation = Atn(Tangent) * 180 / PI
                frmMain.Number_space.Text = Format(Trig_operation, "0")
            ElseIf .Mode_type.Text = "RAD" Then
                Trig_operation = Atn(Tangent)
                frmMain.Number_space.Text = Trig_operation
            End If
        Case 1
            If .Mode_type.Text = "DEG" Then
                Trig_operation = Atn(Tangent) * 180 / PI
                frmMain.Number_space.Text = Format(Trig_operation, "##")
            ElseIf .Mode_type.Text = "RAD" Then
                Trig_operation = Atn(Tangent)
                frmMain.Number_space.Text = Trig_operation
            End If
        Case -1
            If .Mode_type.Text = "DEG" Then
                Trig_operation = 180 - (Atn(Abs(Tangent)) * 180 / PI)
                frmMain.Number_space.Text = Format(Trig_operation, "###")
            ElseIf .Mode_type.Text = "RAD" Then
                Trig_operation = 180 - (Atn(Abs(Tangent)))
                frmMain.Number_space.Text = Trig_operation
            End If
        Case Else
            If Tangent > 0 Then
                If .Mode_type.Text = "DEG" Then
                    Trig_operation = Atn(Tangent) * 180 / PI
                    frmMain.Number_space.Text = Format(Trig_operation, "###0.######")
                ElseIf .Mode_type.Text = "RAD" Then
                    Trig_operation = Atn(Tangent)
                    frmMain.Number_space.Text = Trig_operation
                End If
            ElseIf Tangent < 0 Then
                If .Mode_type.Text = "DEG" Then
                    Trig_operation = 180 - Abs(Atn(Tangent) * 180 / PI)
                    frmMain.Number_space.Text = Format(Trig_operation, "###0.######")
                ElseIf .Mode_type.Text = "RAD" Then
                    Trig_operation = 180 - Abs(Atn(Tangent))
                    frmMain.Number_space.Text = Trig_operation
                End If
            End If
    End Select
End With
End Sub

Public Sub Arcsine()
Dim Sine As Double

With frmMain
    
    Sine = CDbl(frmMain.Number_space.Text)
    
    Select Case Sine
        Case 0.5
            If .Mode_type.Text = "DEG" Then
                Trig_operation = Atn(Sine / Sqr(-Sine * Sine + 1)) * 180 / PI
                frmMain.Number_space.Text = Format(Trig_operation, "##")
            ElseIf .Mode_type.Text = "RAD" Then
                Trig_operation = Atn(Sine / Sqr(-Sine * Sine + 1))
                frmMain.Number_space.Text = Trig_operation
            End If
        Case -0.5
            If .Mode_type.Text = "DEG" Then
                Trig_operation = 180 + Atn(Abs(Sine / Sqr(-Sine * Sine + 1))) * 180 / PI
                frmMain.Number_space.Text = Format(Trig_operation, "###")
            ElseIf .Mode_type.Text = "RAD" Then
                Trig_operation = 180 + Atn(Abs(Sine / Sqr(-Sine * Sine + 1)))
                frmMain.Number_space.Text = Trig_operation
            End If
        Case 1
            If .Mode_type.Text = "DEG" Then
                frmMain.Number_space.Text = "90"
            ElseIf .Mode_type.Text = "RAD" Then
                .Number_space.Text = "1.570796327"
            End If
        Case -1
            If .Mode_type.Text = "DEG" Then
                frmMain.Number_space.Text = "270"
            ElseIf .Mode_type.Text = "RAD" Then
                .Number_space.Text = "-1.570796327"
            End If
        Case Else
            If Sine > 1 Or Sine < -1 Then
                frmMain.Number_space.Text = "-ERROR-"
                MsgBox "The value of sine must be from -1 to 1 inclusive for both degrees and radians.", vbCritical + vbOKOnly, "Error"
                frmMain.Function.Text = " "
                frmMain.Number_space.Text = "0"
                Minus = False
            Else:
                If Sine > 0 Then
                    If .Mode_type.Text = "DEG" Then
                        Trig_operation = Atn(Sine / Sqr(-Sine * Sine + 1)) * 180 / PI
                        frmMain.Number_space.Text = Format(Trig_operation, "###0.######")
                    ElseIf .Mode_type.Text = "RAD" Then
                        Trig_operation = Atn(Sine / Sqr(-Sine * Sine + 1))
                        frmMain.Number_space.Text = Trig_operation
                    End If
                ElseIf Sine < 0 Then
                    If .Mode_type.Text = "DEG" Then
                        Trig_operation = 360 - Atn(Abs(Sine / Sqr(-Sine * Sine + 1))) * 180 / PI
                        frmMain.Number_space.Text = Format(Trig_operation, "###0.######")
                    ElseIf .Mode_type.Text = "RAD" Then
                        Trig_operation = 360 - Atn(Abs(Sine / Sqr(-Sine * Sine + 1)))
                        frmMain.Number_space.Text = Trig_operation
                    End If
                End If
            End If
    End Select
    
End With

End Sub

Public Sub Arccosine()
Dim Cosine As Double

With frmMain
    
    Cosine = CDbl(frmMain.Number_space.Text)
    
    Select Case Cosine
        
        Case 0
            If .Mode_type.Text = "DEG" Then
                Trig_operation = (Atn(-Cosine / Sqr(-Cosine * Cosine + 1)) * 180 / PI) + (2 * Atn(1) * 180 / PI)
                frmMain.Number_space.Text = Format(Trig_operation, "##")
            ElseIf .Mode_type = "RAD" Then
                Trig_operation = (Atn(-Cosine / Sqr(-Cosine * Cosine + 1))) + (2 * Atn(1))
                frmMain.Number_space.Text = Trig_operation
            End If
        Case 0.5
            If .Mode_type.Text = "DEG" Then
                Trig_operation = (Atn(-Cosine / Sqr(-Cosine * Cosine + 1)) * 180 / PI) + (2 * Atn(1) * 180 / PI)
                frmMain.Number_space.Text = Format(Trig_operation, "##")
            ElseIf .Mode_type.Text = "RAD" Then
                Trig_operation = (Atn(-Cosine / Sqr(-Cosine * Cosine + 1))) + (2 * Atn(1))
                frmMain.Number_space.Text = Trig_operation
            End If
        Case -0.5
            If .Mode_type.Text = "DEG" Then
                Trig_operation = (Atn(-Cosine / Sqr(-Cosine * Cosine + 1)) * 180 / PI) + (2 * Atn(1) * 180 / PI)
                frmMain.Number_space.Text = Format(Trig_operation, "###")
            ElseIf .Mode_type.Text = "RAD" Then
                Trig_operation = (Atn(-Cosine / Sqr(-Cosine * Cosine + 1))) + (2 * Atn(1))
                frmMain.Number_space.Text = Trig_operation
            End If
        Case 1
            frmMain.Number_space.Text = "0"
        Case -1
            frmMain.Number_space.Text = "180"
        Case Else
            If Cosine > 1 Or Cosine < -1 Then
                frmMain.Number_space.Text = "-ERROR-"
                MsgBox "The value of cos must be from -1 to 1 inclusive", vbCritical + vbOKOnly, "Error"
                frmMain.Number_space.Text = "0"
                frmMain.Function.Text = " "
                Minus = False
            Else:
                If Cosine > 0 Then
                    If .Mode_type.Text = "DEG" Then
                        Trig_operation = (Atn(-Cosine / Sqr(-Cosine * Cosine + 1)) * 180 / PI) + (2 * Atn(1) * 180 / PI)
                        frmMain.Number_space.Text = Format(Trig_operation, "###0.######")
                    ElseIf .Mode_type.Text = "RAD" Then
                        Trig_operation = (Atn(-Cosine / Sqr(-Cosine * Cosine + 1))) + (2 * Atn(1))
                        frmMain.Number_space.Text = Trig_operation
                    End If
                ElseIf Cosine < 0 Then
                    If .Mode_type.Text = "DEG" Then
                        Trig_operation = (Atn(-Cosine / Sqr(-Cosine * Cosine + 1)) * 180 / PI) + (2 * Atn(1) * 180 / PI)
                        frmMain.Number_space.Text = Format(Trig_operation, "###0.######")
                    ElseIf .Mode_type.Text = "RAD" Then
                        Trig_operation = (Atn(-Cosine / Sqr(-Cosine * Cosine + 1))) + (2 * Atn(1))
                        frmMain.Number_space.Text = Trig_operation
                    End If
                End If
            End If
    End Select
    
End With

End Sub
