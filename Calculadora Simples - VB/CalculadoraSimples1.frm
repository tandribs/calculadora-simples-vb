VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Calculadora Simples"
   ClientHeight    =   4575
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4320
   LinkTopic       =   "Form1"
   ScaleHeight     =   4575
   ScaleWidth      =   4320
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   435
      Left            =   2730
      TabIndex        =   14
      Top             =   3990
      Width           =   825
   End
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "Limpar"
      Height          =   435
      Left            =   1830
      TabIndex        =   13
      Top             =   3990
      Width           =   825
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   765
      Left            =   30
      TabIndex        =   12
      Top             =   3780
      Width           =   4245
      Begin VB.CommandButton cmdSalvar 
         Caption         =   "Salvar"
         Height          =   435
         Left            =   900
         TabIndex        =   17
         Top             =   210
         Width           =   825
      End
   End
   Begin VB.CommandButton cmdSubtrair 
      Caption         =   "Subtrair"
      Height          =   435
      Left            =   1140
      TabIndex        =   8
      Top             =   2310
      Width           =   855
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   30
      TabIndex        =   2
      Top             =   2910
      Width           =   4245
      Begin VB.TextBox txtResultado 
         Height          =   345
         Left            =   2460
         TabIndex        =   18
         Top             =   300
         Width           =   645
      End
      Begin VB.Label lblResultado 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Resultado:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   1110
         TabIndex        =   11
         Top             =   300
         Width           =   1305
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   30
      TabIndex        =   1
      Top             =   2070
      Width           =   4245
      Begin VB.CommandButton cmdDividir 
         Caption         =   "Dividir"
         Height          =   435
         Left            =   3180
         TabIndex        =   10
         Top             =   240
         Width           =   885
      End
      Begin VB.CommandButton cmdMultiplicar 
         Caption         =   "Multiplicar"
         Height          =   435
         Left            =   2130
         TabIndex        =   9
         Top             =   240
         Width           =   885
      End
      Begin VB.CommandButton cmdSomar 
         Caption         =   "Somar"
         Height          =   435
         Left            =   150
         TabIndex        =   7
         Top             =   240
         Width           =   795
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   2085
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4275
      Begin VB.TextBox txtIDCal 
         Height          =   285
         Left            =   3870
         TabIndex        =   15
         Top             =   240
         Width           =   285
      End
      Begin VB.TextBox txtNum2 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   180
         TabIndex        =   6
         Top             =   1440
         Width           =   2025
      End
      Begin VB.TextBox txtNum1 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   150
         TabIndex        =   4
         Top             =   630
         Width           =   2055
      End
      Begin VB.Label lblIDCal 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ID:"
         Height          =   285
         Left            =   3630
         TabIndex        =   16
         Top             =   270
         Width           =   315
      End
      Begin VB.Label lblNum2 
         BackStyle       =   0  'Transparent
         Caption         =   "Segundo Número:"
         Height          =   345
         Left            =   180
         TabIndex        =   5
         Top             =   1110
         Width           =   1665
      End
      Begin VB.Label lblNum1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Primeiro Número:"
         Height          =   315
         Left            =   180
         TabIndex        =   3
         Top             =   270
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type FData
    IDCal As Integer
    Num1 As String * 100
    Num2 As String * 100
    Resultado As String * 100
End Type

Dim Calculadora As FData
Dim FileName As String
Dim IDNum As Integer
Dim FF As Integer

Private Sub Form_Load()
    FF = FreeFile
    FileName = "CalculadoraCopia" & ".txt"
    IDNum = 1
End Sub

Private Sub cmdDividir_Click()
    txtResultado = Val(txtNum1.Text) / Val(txtNum2.Text)
    cmdSalvar.SetFocus
End Sub

Private Sub cmdLimpar_Click()
    txtNum1.Text = ""
    txtNum2.Text = ""
    txtResultado.Text = ""
End Sub

Private Sub cmdMultiplicar_Click()
    txtResultado = Val(txtNum1.Text) * Val(txtNum2.Text)
    cmdSalvar.SetFocus
    End If
End Sub

Private Sub cmdSair_Click()
    End
End Sub

Private Sub cmdSomar_Click()
    txtResultado = Val(txtNum1.Text) + Val(txtNum2.Text)
    cmdSalvar.SetFocus
End Sub

Private Sub cmdSubtrair_Click()
    txtResultado = Val(txtNum1.Text) - Val(txtNum2.Text)
    cmdMultiplicar.SetFocus
    cmdSalvar.SetFocus
End Sub

Private Sub txtIDCal_GotFocus()
    txtNum1.SetFocus
End Sub

Private Sub txtNum1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then
    txtNum2.SetFocus
    End If
End Sub

Private Sub txtNum2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then
    cmdSomar.SetFocus
    End If
End Sub

Private Function verificaNumeros() As Boolean
    verificaNumeros = IsNumeric(txtNum1.Text) And IsNumeric(txtNum2.Text)
    If Not verificaNumeros
End Function
