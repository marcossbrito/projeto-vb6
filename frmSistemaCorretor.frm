VERSION 5.00
Begin VB.MDIForm frmSistemaCorretor 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Controle de Corretores e Clientes"
   ClientHeight    =   4365
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8655
   Icon            =   "frmSistemaCorretor.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuCadastro 
      Caption         =   "&Cadastro"
      Begin VB.Menu mnuCadCliente 
         Caption         =   "Cliente"
      End
      Begin VB.Menu mnuCadCorretor 
         Caption         =   "Corretor"
      End
   End
   Begin VB.Menu mnuConsulta 
      Caption         =   "Consulta"
      Begin VB.Menu mnuConClientes 
         Caption         =   "Clientes"
      End
   End
   Begin VB.Menu mnuSair 
      Caption         =   "&Sair do Sistema"
   End
End
Attribute VB_Name = "frmSistemaCorretor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Unload(Cancel As Integer)
    Set cnnSistemaCorretor = Nothing
End Sub

Private Sub mnuCadCliente_Click()
    frmCadCliente.Show
End Sub

Private Sub mnuCadCorretor_Click()
    frmCadCorretor.Show
End Sub

Private Sub mnuConClientes_Click()
    frmConsultaClientes.Show
End Sub

Private Sub mnuSair_Click()
Dim vSair As Integer
    vSair = MsgBox("Deseja sair do sistema?", _
    vbYesNo + vbQuestion + vbApplicationModal, "Sair do Sistema")
    
    If vSair = vbYes Then End
End Sub
