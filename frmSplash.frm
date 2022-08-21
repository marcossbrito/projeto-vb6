VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4245
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer 
      Interval        =   2000
      Left            =   6000
      Top             =   3480
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Height          =   4050
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   7080
      Begin VB.Label lblCompanhia 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "Viceri - Seidor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   2280
         TabIndex        =   5
         Top             =   1440
         Width           =   4350
      End
      Begin VB.Image imgLogo 
         Height          =   2385
         Left            =   240
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lblIniciando 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Iniciando o Sistema"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   1
         Top             =   3660
         Width           =   6855
      End
      Begin VB.Label lblCandidato 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "Marcos Aurelio Brito da Silva"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3570
         TabIndex        =   2
         Top             =   2700
         Width           =   3285
      End
      Begin VB.Label lblTesteTecnico 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "Teste técnico VB6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4125
         TabIndex        =   3
         Top             =   2340
         Width           =   2730
      End
      Begin VB.Label lblSistemaNome 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Sistema de gerenciamento de Corretores e Clientes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   2280
         TabIndex        =   4
         Top             =   480
         Width           =   4590
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub Timer_Timer()
    On Error GoTo errConexao
    
    cnnSistemaCorretor.ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=db_sistemacorretor;Data Source=.\SQLEXPRESS"
    cnnSistemaCorretor.Open
    
    Unload Me
    frmSistemaCorretor.Show
    Exit Sub
    
errConexao:
    With Err
        If .Number <> 0 Then
            MsgBox "Houve um erro na conexão com o banco de dados." & _
                vbCrLf & "O sistema será encerrado.", _
                vbCritical + vbOKOnly + vbApplicationModal, _
                "Erro na conexão"
            .Number = 0
            Set cnnSistemaCorretor = Nothing
            End
        End If
    End With
End Sub
