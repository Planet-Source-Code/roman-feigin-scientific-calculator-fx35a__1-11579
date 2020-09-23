VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H80000010&
   Caption         =   "Calculator (fx-35A)"
   ClientHeight    =   5715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5415
   Icon            =   "Main.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   5415
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox EXP 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   4920
      MaxLength       =   3
      MultiLine       =   -1  'True
      TabIndex        =   66
      Text            =   "Main.frx":0442
      Top             =   840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Statistic_score 
      Alignment       =   2  'Center
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   3960
      TabIndex        =   65
      Text            =   "n"
      Top             =   840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Memory_indicator 
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   1080
      TabIndex        =   64
      Text            =   "Min"
      Top             =   240
      Width           =   375
   End
   Begin VB.CommandButton E 
      BackColor       =   &H80000018&
      Height          =   375
      Left            =   3360
      Picture         =   "Main.frx":0445
      Style           =   1  'Graphical
      TabIndex        =   63
      ToolTipText     =   "Constant e"
      Top             =   3120
      Width           =   495
   End
   Begin VB.CommandButton Statistical_data 
      BackColor       =   &H00FFFF80&
      Caption         =   "DATA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   56
      ToolTipText     =   "Statistical data entry"
      Top             =   4440
      Width           =   735
   End
   Begin VB.TextBox Statistics_mode 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   480
      TabIndex        =   55
      Text            =   "SD"
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox Mode_type 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   240
      TabIndex        =   54
      Text            =   "DEG"
      Top             =   240
      Width           =   375
   End
   Begin VB.CommandButton Random_number 
      BackColor       =   &H80000018&
      Caption         =   "RAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   53
      ToolTipText     =   "Random number"
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton Memory_clear 
      BackColor       =   &H80000018&
      Caption         =   "MC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   52
      ToolTipText     =   "Clear memory"
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton PI 
      BackColor       =   &H80000018&
      Caption         =   "p"
      BeginProperty Font 
         Name            =   "Symbol"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   51
      ToolTipText     =   "Constant pi"
      Top             =   3120
      Width           =   495
   End
   Begin VB.CommandButton Memory_recall 
      BackColor       =   &H80000018&
      Caption         =   "MR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   50
      ToolTipText     =   "Memory recall"
      Top             =   3120
      Width           =   495
   End
   Begin VB.CommandButton Memory_input 
      BackColor       =   &H80000018&
      Caption         =   "Min"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   49
      ToolTipText     =   "Memory input"
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton Population_standard_deviation 
      BackColor       =   &H80000018&
      Height          =   375
      Left            =   120
      Picture         =   "Main.frx":07C7
      Style           =   1  'Graphical
      TabIndex        =   48
      ToolTipText     =   "Population standard deviation"
      Top             =   3120
      Width           =   495
   End
   Begin VB.CommandButton Sample_standard_deviation 
      BackColor       =   &H80000018&
      Height          =   375
      Left            =   4560
      Picture         =   "Main.frx":0F41
      Style           =   1  'Graphical
      TabIndex        =   47
      ToolTipText     =   "Sample standard deviation"
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton Mean 
      BackColor       =   &H80000018&
      Height          =   375
      Left            =   3840
      Picture         =   "Main.frx":19D3
      Style           =   1  'Graphical
      TabIndex        =   46
      ToolTipText     =   "Arithmetic mean"
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton Clear_s_r 
      BackColor       =   &H0080C0FF&
      Caption         =   "KAC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   45
      ToolTipText     =   "Clear statistical register"
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton Sum 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Symbol"
         Size            =   11.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      Picture         =   "Main.frx":1F3D
      Style           =   1  'Graphical
      TabIndex        =   44
      ToolTipText     =   "Sum of values"
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton Reciprocal 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      Picture         =   "Main.frx":2567
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "Reciprocal"
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton x_Power 
      BackColor       =   &H80000018&
      Height          =   375
      Left            =   1560
      Picture         =   "Main.frx":2DE9
      Style           =   1  'Graphical
      TabIndex        =   42
      Tag             =   "x^y"
      ToolTipText     =   "Power of any number"
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton Ln 
      BackColor       =   &H80000018&
      Caption         =   "ln"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "Natural logarithm"
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton Log 
      BackColor       =   &H80000018&
      Caption         =   "log"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   40
      Tag             =   "lg"
      ToolTipText     =   "Common logarithm"
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox Percentage_indicator 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   3000
      TabIndex        =   39
      Top             =   795
      Width           =   495
   End
   Begin VB.CommandButton Change_sign 
      BackColor       =   &H80000018&
      Caption         =   "+/-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Change sign"
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton Percentage 
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   37
      ToolTipText     =   "Percentage"
      Top             =   5040
      Width           =   495
   End
   Begin VB.CommandButton Scientific_notation 
      Caption         =   "EXP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   36
      ToolTipText     =   "Scientific notation"
      Top             =   5040
      Width           =   735
   End
   Begin VB.CommandButton Trig 
      BackColor       =   &H80000018&
      Caption         =   "tan-1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Inverse tangent"
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton Trig 
      BackColor       =   &H80000018&
      Caption         =   "cos-1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Inverse cosine"
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton Trig 
      BackColor       =   &H80000018&
      Caption         =   "sin-1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Inverse sine"
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton Trig 
      BackColor       =   &H80000018&
      Caption         =   "cosec"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Cosecant"
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton Trig 
      BackColor       =   &H80000018&
      Caption         =   "cot"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Cotangent"
      Top             =   1680
      Width           =   495
   End
   Begin VB.CommandButton Trig 
      BackColor       =   &H80000018&
      Caption         =   "sec"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Secant"
      Top             =   1680
      Width           =   495
   End
   Begin VB.CommandButton Equals 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   29
      Top             =   5040
      Width           =   495
   End
   Begin VB.TextBox Function 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   2160
      TabIndex        =   28
      Text            =   "f(x)"
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox Mode_indicator 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   1680
      TabIndex        =   27
      Text            =   "M"
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox Number_space 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   465
      Left            =   360
      MaxLength       =   18
      MultiLine       =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "Main.frx":383B
      ToolTipText     =   "LCD Display"
      Top             =   360
      Width           =   4815
   End
   Begin VB.CommandButton Quit 
      BackColor       =   &H00FFC0C0&
      Caption         =   "OFF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Exit"
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton Number_squared 
      BackColor       =   &H80000018&
      Height          =   375
      Left            =   3120
      Picture         =   "Main.frx":383D
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Number squared"
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton Fraction 
      BackColor       =   &H80000018&
      Height          =   375
      Left            =   2400
      Picture         =   "Main.frx":411F
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Fraction"
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton Mode 
      BackColor       =   &H0000FFFF&
      Caption         =   "MODE"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Mode"
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton Trig 
      BackColor       =   &H80000018&
      Caption         =   "tan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Tangent"
      Top             =   1680
      Width           =   495
   End
   Begin VB.CommandButton Trig 
      BackColor       =   &H80000018&
      Caption         =   "cos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Cosine"
      Top             =   1680
      Width           =   495
   End
   Begin VB.CommandButton Trig 
      BackColor       =   &H80000018&
      Caption         =   "sin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Sine"
      Top             =   1680
      Width           =   495
   End
   Begin VB.CommandButton N_root 
      BackColor       =   &H80000018&
      Height          =   375
      Left            =   120
      Picture         =   "Main.frx":4A01
      Style           =   1  'Graphical
      TabIndex        =   18
      Tag             =   "N-RT"
      ToolTipText     =   "N-root"
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton Arithmetic_operations 
      Caption         =   "÷"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   3840
      TabIndex        =   16
      Top             =   4440
      Width           =   495
   End
   Begin VB.CommandButton Arithmetic_operations 
      Caption         =   "×"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   3120
      TabIndex        =   15
      Top             =   4440
      Width           =   495
   End
   Begin VB.CommandButton Arithmetic_operations 
      BackColor       =   &H80000013&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   3840
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3840
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton Arithmetic_operations 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   3120
      TabIndex        =   13
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton Clear 
      BackColor       =   &H000000FF&
      Caption         =   "CA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Clear all"
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton Decimal_point 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   11
      Top             =   5040
      Width           =   495
   End
   Begin VB.CommandButton Number 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   1560
      TabIndex        =   10
      Top             =   5040
      Width           =   495
   End
   Begin VB.CommandButton Number 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   840
      TabIndex        =   9
      Top             =   5040
      Width           =   495
   End
   Begin VB.CommandButton Number 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   120
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   8
      Top             =   5040
      Width           =   495
   End
   Begin VB.CommandButton Number 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   2280
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   7
      Top             =   4440
      Width           =   495
   End
   Begin VB.CommandButton Number 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   1560
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   6
      Top             =   4440
      Width           =   495
   End
   Begin VB.CommandButton Number 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   840
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   5
      Top             =   4440
      Width           =   495
   End
   Begin VB.CommandButton Number 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   120
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   4
      Top             =   4440
      Width           =   495
   End
   Begin VB.CommandButton Number 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   2280
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton Number 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   1560
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   2
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton Number 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   840
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   1
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton Square_root 
      BackColor       =   &H80000018&
      Height          =   375
      Left            =   4680
      Picture         =   "Main.frx":5583
      Style           =   1  'Graphical
      TabIndex        =   17
      Tag             =   "SQRT"
      ToolTipText     =   "Square root"
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox LCD 
      BackColor       =   &H80000006&
      ForeColor       =   &H0000FF00&
      Height          =   975
      Left            =   120
      TabIndex        =   26
      Top             =   120
      Width           =   5175
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Caption         =   "4- SD"
      Height          =   255
      Left            =   3120
      TabIndex        =   61
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "3- Rad"
      Height          =   255
      Left            =   2400
      TabIndex        =   60
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "2- Deg"
      Height          =   255
      Left            =   1680
      TabIndex        =   59
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "1-  Comp"
      Height          =   255
      Left            =   840
      TabIndex        =   58
      Top             =   1200
      Width           =   855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      X1              =   2880
      X2              =   2880
      Y1              =   3600
      Y2              =   5760
   End
   Begin VB.Label Modes 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Modes:"
      Height          =   255
      Left            =   120
      TabIndex        =   57
      Top             =   1200
      Width           =   735
   End
   Begin VB.Line Border2 
      BorderColor     =   &H80000009&
      X1              =   0
      X2              =   5400
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Border1 
      BorderColor     =   &H80000009&
      X1              =   0
      X2              =   5400
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H8000000C&
      Caption         =   "fx-35A"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   3840
      TabIndex        =   62
      Top             =   1200
      Width           =   1455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sorry, I also couldn't come up with a good solution '
' to fractions (a/b/c button), so I left it out. If   '
' anyone has a good solution to this module as well   '
' as all the other modules, please send it to me at   '
' foxdetective007@mailcity.com                        '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Argument As Double

Private Sub Arithmetic_operations_Click(Index As Integer)
    Dot = False
    Sign = True
    Misc.Minus = False
    frmMain.Function.Text = ""
    Call Misc.Exponential_validation
    Operation = Arithmetic_operations(Index).Caption
End Sub

Private Sub Change_sign_Click()
    If Misc.Exponent = False Then
        Call Misc.Change_sign
    ElseIf Misc.Exponent = True Then
        Call Misc.Change_exponential_sign
    End If
End Sub

Private Sub Clear_Click()
            
    Misc.Exponent = False
    Sign = False
    Call Misc.Reset
    frmMain.Exp.Text = "00"
    frmMain.Exp.Visible = False
    Base = 0
    Power = 0
    Initial_length = 0
    New_dot_position = 0
    
End Sub

Private Sub Clear_s_r_Click()
    Call Statistical_functions.Clear_s_r
End Sub

Private Sub Decimal_point_Click()
'If decimal point is allowed as soon as pressed, lock it
    
    If Dot = False Then
        Number_space.Text = Number_space.Text & "."
    ElseIf Dot = True Then

    End If
    Dot = True  'decimal point is locked
        
    Misc.Whole_portion = Len(Number_space.Text) - 1
    
End Sub

Private Sub E_Click()
    Number_space.Text = "2.71828182846"
End Sub

Private Sub Equals_Click()
On Error GoTo solution
    Percentage_indicator.Text = ""
    Call Equal_operation.Equals
    
    Exit Sub
solution:
    If Err.Number = 6 Then
        Number_space.Text = "-OVERFLOW ERROR-"
        MsgBox "The value will be too large or too small.Try another calculation." _
        , vbCritical + vbOKOnly, "Error"
        Number_space.Text = "0"
    Else
        Call Misc.solution
    End If
End Sub

Private Sub Form_Load()
    Call Misc.Reset
    Exp.Text = "00"
    Memory_indicator.Text = " "
    Exp.Visible = False
    Mode_type.Text = "DEG"
    Memory_register = 0
    Clear_s_r = False
    Exponential_minus = False
    Misc.x = 0
    Misc.Loop_counter = 0
    Memory_indicator.Visible = False
    Statistics_mode.Visible = False
    Exponential_value1 = " "
    Exponential_value2 = " "
End Sub

Private Sub Fraction_Click()
    MsgBox "Sorry, fractions do not work.", vbCritical + vbOKOnly, _
    "Error"
End Sub

Private Sub Ln_Click()
    
On Error GoTo solution
    Base = CDbl(Number_space.Text)
    If Base = 0 Then
        Number_space.Text = "-ERROR-"
        MsgBox "There is no natural logaritm of 0.", vbCritical + vbOKOnly, _
        "Error"
        Number_space.Text = 0
    ElseIf Base < 0 Then
        Number_space.Text = "-ERROR-"
        MsgBox "There is no natural logaritm of negative base." _
        , vbCritical + vbOKOnly, "Error"
        Number_space.Text = "0"
    Else
        If Sign = False Then
            Value1 = LgN(Base)
            Number_space.Text = Value1
            First_digit = False
        ElseIf Sign = True Then
            Value2 = LgN(Base)
            Number_space.Text = Value2
            First_digit = False
        End If
    End If
    Exit Sub
solution:
    Call Misc.solution
End Sub

Private Sub Log_Click()

On Error GoTo solution
    Base = CDbl(Number_space.Text)

    If Base = 0 Then
        Number_space.Text = "-ERROR-"
        MsgBox "There is no common logaritm of 0.", vbCritical + vbOKOnly, _
        "Error"
        Number_space.Text = 0
    ElseIf Base < 0 Then
        Number_space.Text = "-ERROR-"
        MsgBox "There is no common logaritm of negative base." _
        , vbCritical + vbOKOnly, "Error"
        Number_space.Text = "0"
    Else
        If Sign = False Then
            Value1 = LogN(Base, 10)
            Number_space.Text = Value1
            First_digit = False
        ElseIf Sign = True Then
            Value2 = LogN(Base, 10)
            Number_space.Text = Value2
            First_digit = False
        End If
    End If
    Exit Sub
solution:
    Call Misc.solution
End Sub

Private Sub Mean_Click()
    
If Statistic_score.Visible = True Then
    Call Statistical_functions.Mean
End If

End Sub

Private Sub Memory_clear_Click()
Dim Response As String
    Response = MsgBox("Are you sure you want to clear the memory?" _
    , vbExclamation + vbYesNo, "Clear memory?")
    If Response = vbNo Then
    
    ElseIf Response = vbYes Then
        Memory_register = 0
        Memory_indicator.Text = ""
        Number_space.Text = "0"
    End If
End Sub

Private Sub Memory_input_Click()
    Memory_indicator.Text = "Min"
    Memory_indicator.Visible = True
    Memory_register = CDbl(Number_space.Text)
    frmMain.Function.Text = ""
    Number_space.Text = "0"
End Sub

Private Sub Memory_recall_Click()
    Number_space.Text = Memory_register
    Sign = True
End Sub

Private Sub Mode_Click()
    
If Mode_value = False Then
    Mode_indicator = "M"
    Mode_value = True
ElseIf Mode_value = True Then
    Mode_indicator = ""
    Mode_value = False
End If
    
End Sub

Private Sub N_root_Click()
    Base = CDbl(Number_space.Text)
    Sign = True
    Misc.Minus = False
    Operation = N_root.Tag
End Sub

Private Sub Number_Click(Index As Integer)

If Mode_indicator.Text = "" Then
  If Misc.Exponent = False Then
    If Sign = True Then
        If First_digit = False Then
            Number_space.Text = " "
            If Exp.Visible = True Then
                Exp.Visible = False
                Exp.Text = ""
            End If
            First_digit = True
        End If
    End If

    If Number_space.Text = "0" Or Number_space.Text = " " Then
        Number_space.Text = Number(Index).Caption
    Else:
        Number_space.Text = Number_space.Text + Number(Index).Caption
    End If
  ElseIf Misc.Exponent = True Then
    If Sign = True Then
        If First_digit = False Then
            Exp.Text = " "
            First_digit = True
        End If
    End If

    If Exp.Text = "00" Or Exp.Text = " " Then
        Exp.Text = Number(Index).Caption
    Else:
        Exp.Text = Exp.Text + Number(Index).Caption
    End If

  End If
  'EXP.Text = ""
ElseIf Mode_indicator.Text = "M" Then
    
    Mode_number = Number(Index).Caption
    Call Modes_code.Mode_Validation
    
End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Dim Index1 As Integer
Dim Index2 As Integer

Select Case KeyAscii
    Case Asc("1")
        Index1 = 0
        Call Number_Click(Index1)
    Case Asc("2")
        Index1 = 1
        Call Number_Click(Index1)
    Case Asc("3")
        Index1 = 2
        Call Number_Click(Index1)
    Case Asc("4")
        Index1 = 3
        Call Number_Click(Index1)
    Case Asc("5")
        Index1 = 4
        Call Number_Click(Index1)
    Case Asc("6")
        Index1 = 5
        Call Number_Click(Index1)
    Case Asc("7")
        Index1 = 6
        Call Number_Click(Index1)
    Case Asc("8")
        Index1 = 7
        Call Number_Click(Index1)
    Case Asc("9")
        Index1 = 8
        Call Number_Click(Index1)
    Case Asc("0")
        Index1 = 9
        Call Number_Click(Index1)
    Case Asc(".")
        Call Decimal_point_Click
    Case Asc("-")
       Index2 = 1
       Call Arithmetic_operations_Click(Index2)
    Case Asc("+")
        Index2 = 0
        Call Arithmetic_operations_Click(Index2)
    Case Asc("*")
        Index2 = 2
        Call Arithmetic_operations_Click(Index2)
    Case Asc("/")
        Index2 = 3
        Call Arithmetic_operations_Click(Index2)
    Case Asc("=")
        Call Equal_operation.Equals
    Case Asc("%")
        Call Percentage_Click
    Case Asc("m")
        Call Mode_Click
    Case Asc("M")
        Call Mode_Click
    End Select

End Sub

Private Sub Number_squared_Click()
Dim Square As Double
On Error GoTo solution
    Square = CDbl(Number_space.Text) * CDbl(Number_space.Text)
    Number_space.Text = Square
    Exit Sub
solution:
    Call Misc.solution
End Sub

Private Sub Percentage_Click()
    Value2 = Number_space.Text
    Percentage_indicator.Text = "%"
End Sub

Private Sub PI_Click()
    Number_space.Text = "3.14159265359"
End Sub

Private Sub Population_standard_deviation_Click()
    Call Statistical_functions.Population_standard_deviation
End Sub

Private Sub Quit_Click()
Dim Response1 As String
Dim Response2 As String

    Response1 = MsgBox("Are you sure you want to quit?", vbExclamation + vbYesNo, "Quit")
    If Response1 = vbYes Then
        If Memory_indicator.Text = "Min" Then
        Response2 = MsgBox("If you quit now, the contents of the memory will be lost. Are you sure you want to quit?" _
        , vbExclamation + vbYesNo, "Memory")
            If Response2 = vbYes Then
                Memory_indicator.Text = ""
                Memory_register = 0
                End
            Else:
                frmMain.Refresh
            End If
        Else:
            End
        End If
        
    Else: frmMain.Refresh
    End If
End Sub

Private Sub Random_number_Click()
    Number_space.Text = Format(Rnd(), "0.############")
End Sub

Private Sub Reciprocal_Click()
    Call Misc.Reciprocal
End Sub

Private Sub Sample_standard_deviation_Click()
    Call Statistical_functions.Sample_standard_deviation
End Sub

Private Sub Scientific_notation_Click()
    Exp.Visible = True
    Exp.Text = "00"
    Misc.Exponent = True
    Exponential_minus = False
End Sub

Private Sub Square_root_Click()
    Call Misc.Square_root
End Sub

Private Sub Statistical_data_Click()
        
        Sign = True
        First_digit = False
    
        If Index > 20 Then
        
        ElseIf Index < 20 Then
            Index = Index + 1
        End If
        
        S_R(Index) = CDbl(Number_space.Text)
        Statistic_score.Text = "n" & Index
        Sign = True
        Dot = False
            
End Sub

Private Sub Sum_Click()
    Call Statistical_functions.Sum
End Sub

Private Sub Trig_Click(Index As Integer)
On Error GoTo solution
    Select Case Index
        
        Case "0"
            frmMain.Function.Text = "sin"
            Call Trig_functions.Sine_validation
        Case "1"
            frmMain.Function.Text = "cos"
            Call Trig_functions.Cosine_validation
        Case "2"
            frmMain.Function.Text = "tan"
            Call Trig_functions.Tangent_validation
        Case "3"
            frmMain.Function.Text = "cosec"
            Call Trig_functions.Cosecant_validation
        Case "4"
            frmMain.Function.Text = "sec"
            Call Trig_functions.Secant_validation
        Case "5"
            frmMain.Function.Text = "cot"
            Call Trig_functions.Cotangent_validation
        Case "6"
            frmMain.Function.Text = "arcsin"
            Call Trig_functions.Arcsine
        Case "7"
            frmMain.Function.Text = "arccos"
            Call Trig_functions.Arccosine
        Case "8"
            frmMain.Function.Text = "arctan"
            Call Trig_functions.Arctangent
    End Select
Exit Sub
solution:
    If Err.Number = "6" Then
        Number_space.Text = "-ERROR-"
        MsgBox "The value is too small or too large. Clear all and try again.", _
        vbCritical + vbOKOnly, "Error"
        Number_space.Text = "0"
        frmMain.Function.Text = ""
    Else
        Number_space.Text = "-ERROR-"
        MsgBox "An error has occurred in the calculation. Clear all and try again.", _
        vbCritical + vbOKOnly, "Error"
        Number_space.Text = "0"
        frmMain.Function.Text = ""
    End If

End Sub

Private Sub x_Power_Click()
    Base = CDbl(Number_space.Text)
    Misc.Minus = False
    Operation = x_Power.Tag
    Sign = True
    First_digit = False
End Sub
