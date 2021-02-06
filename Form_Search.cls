VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Abre_Controle_Click()
    DoCmd.OpenTable "Controle"
    'DoCmd.SelectObject acTable, "Controle"
    DoCmd.Requery
End Sub

Private Sub Andamento_Click()
    DoCmd.OpenTable "Controle"
    'DoCmd.Requery
    DoCmd.ApplyFilter , "[Finalizado]=False"
End Sub

Private Sub Cadastrar_Click()
Dim db As Database
Dim rec As Recordset
Dim n_row As Integer

Set db = CurrentDb

MyTable = "Base_Clientes"
valor = Me.Telefone_Input.Value
On Error GoTo ErrorHandler
If IsNull(DLookup("[Cliente]", "Base_Clientes", "[Telefone] =" & valor)) Then
    If Not IsNull(Me.Telefone_Input.Value) And Not IsNull(Me.Cliente_Input.Value) And Not IsNull(Me.Num_Input.Value) And Not IsNull(Me.SearchList.Column(1)) And Len(valor) > 7 Then
        Set rec = db.OpenRecordset("Select * from " & MyTable & "")
        rec.AddNew
        rec("Telefone") = Me.Telefone_Input.Value
        rec("Cliente") = Me.Cliente_Input.Value
        rec("Endereço") = Me!SearchList.Column(1)
        rec("Numero") = Me.Num_Input.Value
        rec("Complemento") = Me.Comp_Input.Value
        rec.Update
        
        rec.Close
        db.Close
        MsgBox ("Cliente Registrado com Sucesso!")
        Me.Telefone_Input.Requery
    Else
ErrorHandler:
        MsgBox ("Favor Preencher Campos Obrigatórios:" + vbCrLf + "• Endereço;" + vbCrLf + "• Cliente;" + vbCrLf + "• Telefone;" + vbCrLf + "• Número de Endereço.")
        Exit Sub
    End If
Else
    MsgBox ("Cliente Já Está Cadastrado!")
    Exit Sub
End If

End Sub

Private Sub Command180_Click()
    On Error GoTo ErrorHandler
    DoCmd.OpenQuery "Finalizados"
    
ErrorHandler:
Exit Sub
End Sub

Private Sub Novo_Pedido_Click()
    Dim ctl As Control
    For Each ctl In Me.Controls
        If ctl.ControlType = acTextBox And ctl.Name <> "Auto_Date" And ctl.Name <> "Auto_Time" Or ctl.Name = "Pedido_Input" Then ctl = Null
    Next
    Set ctl = Nothing
    Resumo_Pedido.RowSource = ""
    Working.RowSource = ""
    Me.Telefone_Input = ""
    
    SearchList.Requery
    SearchList.Selected(1) = True
End Sub

Private Sub Print_Btt_Click()
    DoCmd.OpenReport "Cupom_Fiscal", acViewPreview
End Sub

Private Sub Search_Cliente_Click()
    Dim nome, endereco As String
    Dim num As Integer
    
    valor = Me.Telefone_Input.Value
    If Not IsNull(Me.Telefone_Input.Value) And Len(valor) > 7 And IsNumeric(valor) Then
        If Not IsNull(DLookup("[Cliente]", "Base_Clientes", "[Telefone] =" & valor)) Then
            nome = DLookup("[Cliente]", "Base_Clientes", "[Telefone] =" & valor)
            endereco = DLookup("[Endereço]", "Base_Clientes", "[Telefone] =" & valor)
            num = DLookup("[Numero]", "Base_Clientes", "[Telefone] =" & valor)
            comp = DLookup("[Complemento]", "Base_Clientes", "[Telefone] =" & valor)
            
            Me.Comp_Input.Value = comp
            Me.Cliente_Input.Value = nome
            Me.Num_Input.Value = num
            ListboxSelectString Me!SearchList, endereco
        Else:
            MsgBox ("Cliente não encontrado, Favor Cadastrar!")
        End If
    Else:
        MsgBox ("Favor Preencher o Telefone do Cliente Corretamente")
        Exit Sub
    End If
    
End Sub
Public Sub ListboxSelectString(lstListbox As ListBox, strToSelect As String, Optional SelectDeselect As Boolean = True)
    Dim i As Integer
    'Dim valor As Integer
    For i = 0 To lstListbox.ListCount - 1
        If lstListbox.ItemData(i) = strToSelect Then
            lstListbox.Selected(i) = True
        End If
    Next i
End Sub

Private Sub SearchBox_Change()
    SearchList.Requery
End Sub
Private Sub cmdCopyItem_Click()
 CopySelected Me
 btnSave_OnClick
End Sub
 
Public Sub CopySelected(ByRef frm As Form)
 
 Dim ctlSource As Control
 Dim ctlDest As Control
 Dim strItems As String
 Dim intCurrentRow As Integer
 Dim rowValue As String
 
 Set ctlSource = Me!SearchList
 Set ctlDest = Me!Working
 
 For intCurrentRow = 0 To ctlSource.ListCount - 1
    If ctlSource.Selected(intCurrentRow) Then
        'MsgBox (ctlSource.Column(1))
        strItems = strItems & ctlSource.Column(1, intCurrentRow) & ";"
    End If
 Next intCurrentRow
 
 'Reset destination control's RowSource property.
 ctlDest.RowSource = ""
 ctlDest.RowSource = strItems
 
 Set ctlSource = Nothing
 Set ctlDest = Nothing
 
End Sub

Public Sub btnSave_OnClick()
    Dim db As Database
    Dim rec As Recordset
    Dim n_row As Integer
    
    Set db = CurrentDb
    
    MyTable = "Controle"
    If Not IsNull(Me.Cliente_Input.Value) And Not IsNull(Me.Pedido_Input.Value) And Not IsNull(Me.Num_Input.Value) And Not IsNull(Me.Qtd_Input.Value) And Not IsNull(Me!SearchList.Column(1)) Then
        Set rec = db.OpenRecordset("Select * from " & MyTable & "")
        rec.AddNew
        rec("ID") = rec.RecordCount
        rec("Data") = Me.Auto_Date.Value
        rec("Hora") = Me.Auto_Time.Value
        rec("Cliente") = Me.Cliente_Input.Value
        rec("Pedido") = Me.Pedido_Input.Value
        rec("Quantidade") = Me.Qtd_Input.Value
        rec("Endereço") = Me!SearchList.Column(1)
        rec("Numero_Endereço") = Me.Num_Input.Value
        rec("Complemento_Endereço") = Me.Comp_Input.Value
        rec("Tipo_Endereço") = Me!SearchList.Column(2)
        rec("Tempo Estimado") = Me!SearchList.Column(4)
        rec.Update
    
        rec.Close
        db.Close
        MsgBox ("Pedido Registrado com Sucesso!")
        Resumo_Pedido.RowSource = "SELECT Controle.Pedido, Controle.Quantidade, Controle.Cliente FROM Controle WHERE (((Controle.Cliente)=DLookUp(" & Chr(34) & "[Cliente]" & Chr(34) & "," & Chr(34) & "Controle" & Chr(34) & "," & Chr(34) & "[ID] = DMax('ID','Controle')" & Chr(34) & ")) AND ((Controle.Finalizado)=False));"
        Resumo_Pedido.Requery
    Else
        MsgBox ("Favor Preencher Campos Obrigatórios:" + vbCrLf + "• Endereço;" + vbCrLf + "• Cliente;" + vbCrLf + "• Pedido;" + vbCrLf + "• Quantidade;" + vbCrLf + "• Número de Endereço.")
        Exit Sub
    End If
End Sub

Private Sub Total_Personalizado_Click()
    On Error GoTo ErrorHandler
    DoCmd.OpenQuery "Finalizados_Personalizado"
    
ErrorHandler:
    Exit Sub
End Sub
