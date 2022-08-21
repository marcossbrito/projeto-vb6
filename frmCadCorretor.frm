VERSION 5.00
Begin VB.Form frmCadCorretor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Corretor"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6810
   Icon            =   "frmCadCorretor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   6810
   Begin VB.CommandButton cmdVoltar 
      Caption         =   "Voltar"
      Height          =   855
      Left            =   4080
      Picture         =   "frmCadCorretor.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "Limpar"
      Height          =   855
      Left            =   2760
      Picture         =   "frmCadCorretor.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "Salvar"
      Height          =   855
      Left            =   1440
      Picture         =   "frmCadCorretor.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox txtCodCorretor 
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      Top             =   480
      Width           =   2775
   End
   Begin VB.TextBox txtNome 
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Top             =   960
      Width           =   4335
   End
   Begin VB.TextBox txtCPF 
      Height          =   315
      Left            =   1440
      MaxLength       =   11
      TabIndex        =   3
      Top             =   1440
      Width           =   2775
   End
   Begin VB.Label lblCodCorretor 
      Caption         =   "Código:"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblCliente 
      Caption         =   "Nome:"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblCPF 
      Caption         =   "CPF:"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1440
      Width           =   975
   End
End
Attribute VB_Name = "frmCadCorretor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdLimpar_Click()
    Limpar
End Sub

Private Sub cmdSalvar_Click()
    Dim sql As String
    Dim cnnComando As New ADODB.Command
    Dim rsSelecao As New ADODB.Recordset
    
    If txtCodCorretor.Text = Empty Or txtNome.Text = Empty Or txtCPF.Text = Empty Then
        MsgBox "Todos os campos são obrigatórios!", _
                vbExclamation + vbOKOnly + vbSystemModal, "Erro"
        Exit Sub
    End If
    
    If Len(txtCPF.Text) < 11 Or Len(txtCPF.Text) > 11 Then
        MsgBox "CPF deve conter 11 dígitos", vbExclamation + vbOKOnly + vbApplicationModal, "Erro"
        Exit Sub
    End If
    
    With cnnComando
        .ActiveConnection = cnnSistemaCorretor
        .CommandType = adCmdText
        .CommandText = "SELECT cpf FROM corretor WHERE cpf = '" & txtCPF.Text & "'"
        Set rsSelecao = .Execute
    End With
    
    With rsSelecao
        If Not .EOF And Not .BOF Then
            MsgBox "CPF Já existente!", vbExclamation + vbOKOnly + vbApplicationModal, "Erro"
            Exit Sub
        End If
    End With
    
    
    sql = "INSERT INTO corretor (codigo, nome, cpf) VALUES ('" & txtCodCorretor.Text & "', '" & txtNome.Text & "', '" & txtCPF.Text & "')"
    
    With cnnComando
        .ActiveConnection = cnnSistemaCorretor
        .CommandType = adCmdText
        .CommandText = sql
        .Execute
    End With
    MsgBox "Corretor inserido com sucesso!", _
            vbApplicationModal + vbInformation + vbOKOnly, "Salvar"
    
    Limpar
End Sub

Private Sub cmdVoltar_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{Tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    Me.Left = (frmSistemaCorretor.ScaleWidth - Me.Width) / 2
    Me.Top = (frmSistemaCorretor.ScaleHeight - Me.Height) / 2
End Sub

Private Sub Limpar()
    txtCodCorretor.Text = Empty
    txtNome.Text = Empty
    txtCPF.Text = Empty
End Sub

Private Sub txtCPF_LostFocus()
    Dim cnnComando As New ADODB.Command
    Dim rsSelecao As New ADODB.Recordset
    
End Sub
