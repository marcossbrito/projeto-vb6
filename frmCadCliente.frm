VERSION 5.00
Begin VB.Form frmCadCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Cliente"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8040
   Icon            =   "frmCadCliente.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   8040
   Begin VB.CheckBox chkAtivo 
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "Salvar"
      Height          =   855
      Left            =   1440
      Picture         =   "frmCadCliente.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "Limpar"
      Height          =   855
      Left            =   2760
      Picture         =   "frmCadCliente.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdVoltar 
      Caption         =   "Voltar"
      Height          =   855
      Left            =   4080
      Picture         =   "frmCadCliente.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4080
      Width           =   1095
   End
   Begin VB.ComboBox cmbCidade 
      Height          =   315
      Left            =   1440
      TabIndex        =   6
      Top             =   2880
      Width           =   3255
   End
   Begin VB.ComboBox cmbEstado 
      Height          =   315
      Left            =   1440
      TabIndex        =   5
      Top             =   2400
      Width           =   3255
   End
   Begin VB.TextBox txtEndereco 
      Height          =   315
      Left            =   1440
      TabIndex        =   4
      Top             =   1920
      Width           =   5535
   End
   Begin VB.TextBox txtCPF 
      Height          =   315
      Left            =   1440
      MaxLength       =   11
      TabIndex        =   3
      Top             =   1440
      Width           =   2775
   End
   Begin VB.TextBox txtNome 
      Height          =   315
      Left            =   1440
      TabIndex        =   2
      Top             =   960
      Width           =   4335
   End
   Begin VB.ComboBox cmbCorretor 
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Top             =   480
      Width           =   4335
   End
   Begin VB.Label lblAtivo 
      Caption         =   "Ativo:"
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label lblEstado 
      Caption         =   "Estado:"
      Height          =   255
      Left            =   360
      TabIndex        =   15
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label lblCidade 
      Caption         =   "Cidade:"
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label lblEndereco 
      Caption         =   "Endereço:"
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label lblCPF 
      Caption         =   "CPF:"
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label lblCliente 
      Caption         =   "Nome:"
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblCorretor 
      Caption         =   "Corretor:"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "frmCadCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public vEstado, vCidade, vCorretor As Long
Dim vAtivo

Private Sub chkAtivo_Click()
    If chkAtivo.Value = vbChecked Then
        vAtivo = 1
    Else
        vAtivo = 0
    End If
End Sub

Private Sub cmbCorretor_Click()
    With cmbCorretor
        If .ListIndex <> -1 Then
            vCorretor = .ItemData(.ListIndex)
        Else
            vCorretor = 0
        End If
    End With
End Sub

Private Sub cmbCidade_Click()
    With cmbCidade
        If .ListIndex <> -1 Then
            vCidade = .ItemData(.ListIndex)
        Else
            vCidade = 0
        End If
    End With
End Sub

Private Sub cmbEstado_Click()
    With cmbEstado
        If .ListIndex <> -1 Then
            vEstado = .ItemData(.ListIndex)
            vIdEstado = vEstado
            ComboCidade cmbCidade
        Else
            vEstado = 0
        End If
    End With
End Sub

Private Sub cmdSalvar_Click()
    Dim sql As String
    Dim cnnComando As New ADODB.Command
    Dim rsSelecao As New ADODB.Recordset
    
    If cmbCorretor.Text = Empty Or txtNome.Text = Empty Or txtCPF.Text = Empty Or txtEndereco.Text = Empty Or cmbEstado.Text = Empty Or cmbCidade.Text = Empty Then
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
        .CommandText = "SELECT cpf FROM cliente WHERE cpf = '" & txtCPF.Text & "'"
        Set rsSelecao = .Execute
    End With
    
    With rsSelecao
        If Not .EOF And Not .BOF Then
            MsgBox "CPF Já existente!", vbExclamation + vbOKOnly + vbApplicationModal, "Erro"
            Exit Sub
        End If
    End With
    
    sql = "INSERT INTO cliente (nome, cpf, endereco, ativo, id_uf, id_cidade, id_corretor) "
    sql = sql & "VALUES ('" & txtNome.Text & "', '" & txtCPF.Text & "', '" & txtEndereco.Text & "',"
    sql = sql & "" & vAtivo & "," & vEstado & "," & vCidade & "," & vCorretor & ")"
    
    With cnnComando
        .ActiveConnection = cnnSistemaCorretor
        .CommandType = adCmdText
        .CommandText = sql
        .Execute
    End With
    MsgBox "Cliente inserido com sucesso!", _
            vbApplicationModal + vbInformation + vbOKOnly, "Salvar"
    
    Limpar
End Sub

Private Sub Form_Load()
    Me.Left = (frmSistemaCorretor.ScaleWidth - Me.Width) / 2
    Me.Top = (frmSistemaCorretor.ScaleHeight - Me.Height) / 2
    
    vEstado = 0
    vCidade = 0
    vAtivo = chkAtivo
    
    ComboEstado cmbEstado
    ComboCorretor cmbCorretor
    
    cmbEstado.ListIndex = -1
    cmbCidade.ListIndex = -1
    cmbCorretor.ListIndex = -1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{Tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub cmdLimpar_Click()
    Limpar
End Sub

Private Sub cmdVoltar_Click()
    Unload Me
End Sub

Private Sub Limpar()
    txtNome.Text = Empty
    txtCPF.Text = Empty
    txtEndereco.Text = Empty
    cmbCorretor.Text = Empty
    cmbCidade.Text = Empty
    cmbEstado.Text = Empty
    chkAtivo.Value = 0
End Sub
