Attribute VB_Name = "Module1"
Global cnnSistemaCorretor As New ADODB.Connection
Public vIdEstado As Long

Public Sub ComboEstado(NomeCombo As ComboBox)
    Dim sql As String
    Dim cnnComando As New ADODB.Command
    Dim rsTemp As New ADODB.Recordset
    Dim i As Integer
    
    sql = "SELECT * FROM estado ORDER BY uf"
    
    With cnnComando
        .ActiveConnection = cnnSistemaCorretor
        .CommandType = adCmdText
        .CommandText = sql
        Set rsTemp = .Execute
    End With
    
    With rsTemp
        .MoveFirst
        i = 0
        While Not .EOF
            NomeCombo.AddItem !uf, i
            NomeCombo.ItemData(i) = !id
            .MoveNext
            i = i + 1
        Wend
    End With
    
saida:
    Set cnnComando = Nothing
    Set rsTemp = Nothing
    Exit Sub
End Sub

Public Sub ComboCidade(NomeCombo As ComboBox)
    Dim sql As String
    Dim cnnComando As New ADODB.Command
    Dim rsTemp As New ADODB.Recordset
    Dim i As Integer
    
    sql = "SELECT * FROM cidade WHERE id_uf = " & vIdEstado & " ORDER BY nome"
    
    With cnnComando
        .ActiveConnection = cnnSistemaCorretor
        .CommandType = adCmdText
        .CommandText = sql
        Set rsTemp = .Execute
    End With
    
    With rsTemp
        NomeCombo.Clear
        .MoveFirst
        i = 0
        While Not .EOF
            NomeCombo.AddItem !nome, i
            NomeCombo.ItemData(i) = !id
            .MoveNext
            i = i + 1
        Wend
    End With
    
saida:
    Set cnnComando = Nothing
    Set rsTemp = Nothing
    Exit Sub
End Sub

Public Sub ComboCorretor(NomeCombo As ComboBox)
    Dim sql As String
    Dim cnnComando As New ADODB.Command
    Dim rsTemp As New ADODB.Recordset
    Dim i As Integer
    
    sql = "SELECT * FROM corretor ORDER BY nome"
    
    With cnnComando
        .ActiveConnection = cnnSistemaCorretor
        .CommandType = adCmdText
        .CommandText = sql
        Set rsTemp = .Execute
    End With
    
    With rsTemp
        .MoveFirst
        i = 0
        While Not .EOF
            NomeCombo.AddItem !nome, i
            NomeCombo.ItemData(i) = !id
            .MoveNext
            i = i + 1
        Wend
    End With
    
saida:
    Set cnnComando = Nothing
    Set rsTemp = Nothing
    Exit Sub
End Sub
