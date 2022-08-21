VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmConsultaClientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Clientes"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12000
   Icon            =   "frmConsultaClientes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   12000
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshClientes 
      Height          =   3735
      Left            =   240
      TabIndex        =   11
      Top             =   3000
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   6588
      _Version        =   393216
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.CommandButton cmdCadCorretor 
      Caption         =   "Cadastrar Corretor"
      Height          =   1095
      Left            =   10680
      Picture         =   "frmConsultaClientes.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton cmdCadCliente 
      Caption         =   "Cadastrar Cliente"
      Height          =   1095
      Left            =   10680
      Picture         =   "frmConsultaClientes.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   240
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filtro"
      Height          =   2655
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   10335
      Begin VB.CommandButton cmdPesquisar 
         Caption         =   "Pesquisar"
         DownPicture     =   "frmConsultaClientes.frx":0CC6
         DragIcon        =   "frmConsultaClientes.frx":1108
         Height          =   735
         Left            =   7320
         Picture         =   "frmConsultaClientes.frx":154A
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1800
         Width           =   2775
      End
      Begin VB.ComboBox cmbCidade 
         Height          =   315
         Left            =   7320
         TabIndex        =   7
         Top             =   1320
         Width           =   2775
      End
      Begin VB.ComboBox cmbEstado 
         Height          =   315
         Left            =   7320
         TabIndex        =   6
         Top             =   840
         Width           =   2775
      End
      Begin VB.CheckBox chkAtivo 
         Height          =   285
         Left            =   7320
         TabIndex        =   5
         Top             =   360
         Width           =   255
      End
      Begin VB.TextBox txtCPFCliente 
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   1800
         Width           =   2295
      End
      Begin VB.TextBox txtNomeCliente 
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Top             =   1320
         Width           =   3855
      End
      Begin VB.TextBox txtNomeCorretor 
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Top             =   840
         Width           =   3855
      End
      Begin VB.TextBox txtCodCorretor 
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdLimpar 
         Caption         =   "Limpar filtros"
         Height          =   735
         Left            =   5880
         Picture         =   "frmConsultaClientes.frx":198C
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lblCidade 
         Caption         =   "Cidade:"
         Height          =   255
         Left            =   6360
         TabIndex        =   18
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label lblEstado 
         Caption         =   "Estado:"
         Height          =   255
         Left            =   6360
         TabIndex        =   17
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblAtivo 
         Caption         =   "Ativo:"
         Height          =   255
         Left            =   6360
         TabIndex        =   16
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblCPFCliente 
         Caption         =   "CPF Cliente:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label 
         Caption         =   "Nome Cliente:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label lbNomeCorretor 
         Caption         =   "Nome Corretor:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblCodCorretor 
         Caption         =   "Código Corretor:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmConsultaClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public vEstado, vCidade, vCorretor As Long

Public Enum ConsultaClientes
    nomeCliente = 0
    cpfCliente = 1
    ativo = 2
    nomeCorretor = 3
    codCorretor = 4
    estado = 5
    cidade = 6
End Enum

Private Sub chkAtivo_Click()
    If chkAtivo.Value = vbChecked Then
        vAtivo = 1
    Else
        vAtivo = 0
    End If
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

Private Sub cmdCadCliente_Click()
    frmCadCliente.Show
End Sub

Private Sub cmdCadCorretor_Click()
    frmCadCorretor.Show
End Sub

Private Sub cmdLimpar_Click()
    Limpar
End Sub

Private Sub cmdPesquisar_Click()
Dim sql As String
Dim rs As New ADODB.Recordset

sql = "SELECT c.nome AS 'Nome Cliente', c.cpf AS CPF, c.ativo AS Ativo, co.nome AS 'Nome Corretor', "
sql = sql & "co.codigo AS 'Cod. Corretor', e.uf AS UF, ci.nome AS Cidade "
sql = sql & "FROM cliente AS c "
sql = sql & "INNER JOIN corretor AS co "
sql = sql & "ON c.id_corretor = co.id "
sql = sql & "INNER JOIN estado AS e "
sql = sql & "ON c.id_uf = e.id "
sql = sql & "INNER JOIN cidade AS ci "
sql = sql & "ON c.id_cidade = ci.id "
sql = sql & "WHERE c.id != 0 "
If txtCodCorretor.Text <> Empty Then
    sql = sql & "AND co.codigo = '" & txtCodCorretor.Text & "' "
End If
If txtNomeCorretor.Text <> Empty Then
    sql = sql & "AND co.nome = '" & txtNomeCorretor.Text & "' "
End If
If txtNomeCliente.Text <> Empty Then
    sql = sql & "AND c.nome = '" & txtNomeCliente.Text & "' "
End If
If txtCPFCliente.Text <> Empty Then
    sql = sql & "AND c.cpf = '" & txtCPFCliente.Text & "' "
End If
If chkAtivo.Value = 1 Then
    sql = sql & "AND c.ativo = " & 1 & " "
Else
    sql = sql & "AND c.ativo = " & 0 & " "
End If
If cmbEstado.Text <> Empty Then
    sql = sql & "AND e.uf = '" & cmbEstado.Text & "' "
End If
If cmbCidade.Text <> Empty Then
    sql = sql & "AND ci.nome = '" & cmbCidade.Text & "' "
End If





'sql = sql & "and c.nome = 'testa'"

Set rs = New ADODB.Recordset

rs.Open sql, cnnSistemaCorretor, adOpenStatic

Set mshClientes.DataSource = rs

rs.Close

End Sub

Private Sub Form_Load()
    Me.Left = (frmSistemaCorretor.ScaleWidth - Me.Width) / 2
    Me.Top = (frmSistemaCorretor.ScaleHeight - Me.Height) / 2
    
    vEstado = 0
    vCidade = 0
    chkAtivo.Value = 1
    
    ComboEstado cmbEstado
    
    cmbEstado.ListIndex = -1
    cmbCidade.ListIndex = -1
    
    mshClientes.ColWidth(ConsultaClientes.nomeCliente) = 5000
    mshClientes.ColWidth(ConsultaClientes.cpfCliente) = 2500
    mshClientes.ColWidth(ConsultaClientes.ativo) = 1000
    mshClientes.ColWidth(ConsultaClientes.nomeCorretor) = 5000
    mshClientes.ColWidth(ConsultaClientes.codCorretor) = 2500
    mshClientes.ColWidth(ConsultaClientes.estado) = 1000
    mshClientes.ColWidth(ConsultaClientes.cidade) = 3000
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{Tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub Limpar()
    txtCodCorretor.Text = Empty
    txtNomeCliente.Text = Empty
    txtNomeCorretor.Text = Empty
    txtCPFCliente.Text = Empty
    chkAtivo.Value = 1
    cmbEstado.Text = Empty
    cmbCidade.Text = Empty
End Sub
