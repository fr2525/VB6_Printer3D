VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form frmTipos 
   Caption         =   "Tipos de Pets"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6780
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   6780
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lstTipos 
      Height          =   3675
      Left            =   240
      TabIndex        =   6
      Top             =   1140
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   6482
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.TextBox txtAnimal 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   180
      MaxLength       =   50
      TabIndex        =   1
      Top             =   5010
      Width           =   6390
   End
   Begin Threed.SSCommand cmd_Adicionar 
      Height          =   675
      Left            =   180
      TabIndex        =   0
      Top             =   210
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   1191
      _StockProps     =   78
      Caption         =   " &Novo"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      Picture         =   "frmTipos.frx":0000
   End
   Begin Threed.SSCommand cmd_Limpar 
      Height          =   675
      Left            =   1533
      TabIndex        =   2
      Top             =   210
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   1191
      _StockProps     =   78
      Caption         =   " &Limpar"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      BevelWidth      =   1
      Picture         =   "frmTipos.frx":015A
   End
   Begin Threed.SSCommand cmd_Gravar 
      Height          =   675
      Left            =   2886
      TabIndex        =   3
      Top             =   210
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   1191
      _StockProps     =   78
      Caption         =   "&Gravar"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      BevelWidth      =   1
      Picture         =   "frmTipos.frx":02B4
   End
   Begin Threed.SSCommand cmd_Sair 
      Height          =   675
      Left            =   5595
      TabIndex        =   4
      Top             =   210
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   1191
      _StockProps     =   78
      Caption         =   "&Sair"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      Picture         =   "frmTipos.frx":040E
   End
   Begin Threed.SSCommand cmd_Excluir 
      Height          =   675
      Left            =   4239
      TabIndex        =   5
      Top             =   210
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   1191
      _StockProps     =   78
      Caption         =   "&Excluir"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      BevelWidth      =   1
      Picture         =   "frmTipos.frx":0728
   End
End
Attribute VB_Name = "frmTipos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iTipoOperacao As Integer

Private Sub Carrega_Colunas_Tipos()
    With lstTipos
        .ListItems.Clear
        .ColumnHeaders.Clear
        .View = lvwReport
        .ColumnHeaders.Add 1, , "Id", 300, lvwColumnLeft
        .ColumnHeaders.Add 2, , "Descrição", 4900, lvwColumnLeft
    End With
End Sub

Private Sub MontaColunas_Tipos()
    
    Call sConectaLocal
    strSql = ""
    strSql = strSql & " SELECT ID,DESCRICAO FROM TAB_tipos_an ORDER BY DESCRICAO"
    
    gRs.Close
    gSql = "Select * from tb_clientes "
    gSql = gSql & " WHERE id = " & Val(Me.LblCodclie.Caption)
    gRs.Open gSql, ConDb, adOpenForwardOnly
    
    row = sqlite_get_table(DBz, strSql, minfo) ' query database
    numrows = number_of_rows_from_last_call ' bilangan rows data yang di select
    Call closeDB
    
   ' Set Rstemp = New ADODB.Recordset
   ' Rstemp.Open strSql, CnnLocal, 1, 2
   If numrows > 0 Then
   ' If Rstemp.RecordCount <> 0 Then
   '     Rstemp.MoveLast
   '     Rstemp.MoveFirst
        'fmeListaPedidos.Visible = True
        
   '     For x = 1 To Rstemp.RecordCount
        For X = 1 To numrows
            lstTipos.ListItems.Add X, , row(X, 0)
            
            If Not IsNull(row(X, 1)) Then
                lstTipos.ListItems(X).SubItems(1) = row(X, 1)
            Else
                lstTipos.ListItems(X).SubItems(1) = ""
            End If
'            If Not IsNull(Rstemp!RAZAO_SOCIAL) Then
'                List_Atendimentos.ListItems(X).SubItems(2) = UCase(Rstemp!RAZAO_SOCIAL)
'            Else
'                  List_Atendimentos.ListItems.Add(X).SubItems(2) = "Fornecedor n o Encontrado...!"
'            End If
'            If Not IsNull(Rstemp!VALOR_TOTAL) Then
'                List_Atendimentos.ListItems(X).SubItems(3) = Format(Rstemp!VALOR_TOTAL, "0.00")
'            Else
'                List_Atendimentos.ListItems.Add(X).SubItems(3) = ""
'            End If
            
           ' Rstemp.MoveNext
        Next
        'lstTipos.SetFocus
       
    Else
        MsgBox "Sem registros", vbOKOnly
        'fmeListaPedidos.Visible = False
    End If
    
   ' Rstemp.Close
   Call closeDB
    Set Rstemp = Nothing
    
End Sub

Private Sub cmd_Adicionar_Click()
    txtAnimal.Enabled = True
    txtAnimal.SetFocus
    txtAnimal.Text = ""
    cmd_Adicionar.Enabled = False
    cmd_Gravar.Enabled = True
    cmd_Excluir.Enabled = False
    iTipoOperacao = 1
End Sub

Private Sub cmd_Excluir_Click()
    If Len(txtAnimal.Text) = 0 Or txtAnimal.Text = "" Then
       MsgBox "Tipo de Animal inv lido. Favor corrigir", vbOKOnly
       txtAnimal.SetFocus
       Exit Sub
    End If
    
    If MsgBox("Tem certeza que deseja excluir o tipo de animal: " & Chr(13) & Chr(10) & _
                            Trim(lstTipos.SelectedItem.ListSubItems.Item(1)), vbYesNo) = vbYes Then
        If fExcluir_Tipo_Pet() Then
            cmd_Adicionar.Enabled = True
            cmd_Excluir.Enabled = False
            cmd_Gravar.Enabled = False
            lstTipos.ListItems.Clear
            Call MontaColunas_Tipos
            If lstTipos.ListItems.Count > 0 Then
                lstTipos.ListItems(1).Selected = True
                txtAnimal.Text = Trim(lstTipos.SelectedItem.ListSubItems.Item(1))
            End If
        Else
            MsgBox "Erro ao excluir o tipo de PET: " & Err.Description
        End If
    End If
End Sub

Private Sub cmd_Gravar_Click()
    If Len(txtAnimal.Text) = 0 Or txtAnimal.Text = "" Then
       MsgBox "Tipo de Animal inválido. Favor corrigir", vbOKOnly
       txtAnimal.SetFocus
       Exit Sub
    End If
    If fGravar_Tipo_Pet() Then
        cmd_Adicionar.Enabled = True
        cmd_Gravar.Enabled = False
        cmd_Limpar.Enabled = False
        
        'cmd_Excluir.Enabled = true
        lstTipos.ListItems.Clear
        Call MontaColunas_Tipos
        lstTipos.ListItems(1).Selected = True
        txtAnimal.Text = Trim(lstTipos.SelectedItem.ListSubItems.Item(1))
        'Call cmd_Limpar_Click
    Else
        MsgBox "Erro ao incluir o tipo de PET: " & Err.Description
    End If
End Sub

Private Sub cmd_Limpar_Click()
    txtAnimal.Text = ""
    'txtAnimal.SetFocus
    cmd_Adicionar.Enabled = False
    cmd_Gravar.Enabled = True
End Sub

Private Sub cmd_Sair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call Carrega_Colunas_Tipos
    Call MontaColunas_Tipos
    'lstTipos.ListItems = 1
    If lstTipos.ListItems.Count > 0 Then
        txtAnimal.Text = Trim(lstTipos.SelectedItem.ListSubItems.Item(1))
    End If
    
End Sub

Private Function fGravar_Tipo_Pet()
    
    Dim crows As Variant ' current rows (prive variable)
    
    If Len(txtAnimal.Text) = 0 Or txtAnimal.Text = "" Then
       MsgBox "Tipo de Animal inv lido. Favor corrigir", vbOKOnly
       txtAnimal.SetFocus
       Exit Function
    End If
    
    fGravar_Tipo_Pet = True
    
    On Error GoTo Erro_fGravar_Tipo_Pet
    
    If iTipoOperacao = 1 Then
        strSql = "INSERT INTO tab_tipos_an (DESCRICAO, OPERADOR, DT_ATUALIZA)"
        strSql = strSql + " VALUES( '" & UCase(txtAnimal.Text) & "','" & gOperador & "','" & Format(Now, "yyyy/mm/dd hh:mm:ss") & "')"
        
    Else
        strSql = "UPDATE tab_tipos_an SET DESCRICAO = '" & UCase(txtAnimal.Text) & _
                                          "',OPERADOR = '" & gOperador & _
                                          "', DT_ATUALIZA = '" & Format(Now, "yyyy/mm/dd hh:mm:ss") & _
                                          "' WHERE ID = '" & lstTipos.SelectedItem.Text & "'"
    End If
    
    
    'query = "insert into users (nama,ic) VALUES ('" + Text4.Text + "','" + Text3.Text + "')"
    crows = sqlite_get_table(DBz, strSql, minfo) ' query database
    
    If (minfo = "") Then
        MsgBox "Operação efetuada com sucesso"
    Else
        MsgBox "Error: minfo"
    End If
   Call closeDB
    
    'CnnLocal.Execute strSql
    Exit Function
Erro_fGravar_Tipo_Pet:
    fGravar_Tipo_Pet = False
End Function

Private Function fExcluir_Tipo_Pet()
    
    fExcluir_Tipo_Pet = True
    
    On Error GoTo Erro_fExcluir_Tipo_Pet
    
    strSql = "DELETE from tab_tipoS_an WHERE ID = '" & lstTipos.SelectedItem.Text & "'"
    CnnLocal.Execute strSql
    Exit Function
Erro_fExcluir_Tipo_Pet:
    fExcluir_Tipo_Pet = False
End Function

Private Sub lstTipos_ItemClick(ByVal Item As MSComctlLib.ListItem)
    txtAnimal.Text = Trim(lstTipos.SelectedItem.ListSubItems.Item(1))
    txtAnimal.Enabled = True
    cmd_Gravar.Enabled = True
    cmd_Excluir.Enabled = True
    cmd_Limpar.Enabled = True
    iTipoOperacao = 2
End Sub

Private Sub lstTipos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If lstTipos.ListItems.Count > 0 Then
            SendKeys "{tab}"
        End If
    Else
        Call lstTipos_Click
    End If
End Sub

Private Sub txtAnimal_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    If Len(txtAnimal.Text) = 0 Or txtAnimal.Text = "" Then
       MsgBox "Tipo de Animal inv lido. Favor corrigir", vbOKOnly
       txtAnimal.SetFocus
    Else
       cmd_Gravar.SetFocus
    End If
End If
End Sub
