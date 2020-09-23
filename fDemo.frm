VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form fDemo 
   Caption         =   "GradientFill API Demo"
   ClientHeight    =   5985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8220
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   399
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   548
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame fraControls 
      Caption         =   "Controls"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      Begin VB.OptionButton optInterp 
         Caption         =   "HLS"
         Height          =   255
         Index           =   2
         Left            =   2040
         TabIndex        =   14
         Top             =   1920
         Width           =   735
      End
      Begin VB.OptionButton optInterp 
         Caption         =   "Cosine"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   13
         Top             =   1920
         Width           =   855
      End
      Begin VB.OptionButton optInterp 
         Caption         =   "Linear"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   1920
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.CheckBox chkVertical 
         Caption         =   "Vertical"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txtRepeats 
         Height          =   315
         Left            =   2280
         TabIndex        =   9
         Text            =   "1"
         Top             =   1200
         Width           =   375
      End
      Begin VB.CommandButton cmdMultiColor 
         Caption         =   "Draw Gradient"
         Height          =   495
         Left            =   1320
         TabIndex        =   10
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label lblCap 
         Caption         =   "Interpolation Method"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   15
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Label lblCap 
         BackStyle       =   0  'Transparent
         Caption         =   "(left-click to change color, right-click to de/activate)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   720
         TabIndex        =   11
         Top             =   240
         Width           =   2775
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblCap 
         Alignment       =   1  'Right Justify
         Caption         =   "Repeats"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   8
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label lblCap 
         Alignment       =   1  'Right Justify
         Caption         =   "Colors"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblCol 
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   4
         Left            =   2280
         TabIndex        =   6
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lblCol 
         BackColor       =   &H00C000C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   3
         Left            =   1800
         TabIndex        =   5
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lblCol 
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   2
         Left            =   1320
         TabIndex        =   4
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lblCol 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   1
         Left            =   840
         TabIndex        =   3
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lblCol 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   2
         Top             =   720
         Width           =   375
      End
   End
   Begin MSComDlg.CommonDialog cdlColor 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
End
Attribute VB_Name = "fDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pGradientFill.vbp
' 2005 May 12
' redbird77@earthlink.net
' http://home.earthlink.net/~redbird77

' Demo showing how to create multicolor repeating gradients using the
' GradientFill API function or a user defined function (which can
' utilize linear, cosine, or HLS color interpolation).

Option Explicit

Private Sub cmdMultiColor_Click()

Dim bRet    As Boolean
Dim iInterp As Integer

    ' Can be 0, 1, or 2.
    iInterp = IIf(optInterp(0).Value, 0, IIf(optInterp(1).Value, 1, 2))
                     
    bRet = Gradient(Me.hDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, _
                    chkVertical.Value, CLng(txtRepeats.Text), iInterp, _
                    IIf(lblCol(0).BorderStyle, lblCol(0).BackColor, -1), _
                    IIf(lblCol(1).BorderStyle, lblCol(1).BackColor, -1), _
                    IIf(lblCol(2).BorderStyle, lblCol(2).BackColor, -1), _
                    IIf(lblCol(3).BorderStyle, lblCol(3).BackColor, -1), _
                    IIf(lblCol(4).BorderStyle, lblCol(4).BackColor, -1))
                    
    If Not bRet Then MsgBox "Gradient failed.", vbExclamation, "Error"

End Sub

Private Sub lblCol_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error GoTo ErrExit
    
    If Button = vbLeftButton Then
        
        With cdlColor
            .ShowColor
            lblCol(Index).BackColor = .Color
        End With
        
    Else
        ' Toggle border style.
        lblCol(Index).BorderStyle = lblCol(Index).BorderStyle Xor 1
    End If
    
    Exit Sub
    
ErrExit:
End Sub
