VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmStatus 
   BackColor       =   &H00808000&
   BorderStyle     =   0  'None
   Caption         =   "Status"
   ClientHeight    =   3915
   ClientLeft      =   -60
   ClientTop       =   -120
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   ScaleHeight     =   3915
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   150
      TabIndex        =   2
      Top             =   2880
      Width           =   4245
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Salvar"
         Enabled         =   0   'False
         Height          =   540
         Left            =   2190
         Picture         =   "frmStatus.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Tag             =   "&Update"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton cmdEditar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Editar"
         Height          =   540
         Left            =   795
         Picture         =   "frmStatus.frx":00FA
         Style           =   1  'Graphical
         TabIndex        =   7
         Tag             =   "&Refresh"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Excluir"
         Height          =   540
         Left            =   1485
         Picture         =   "frmStatus.frx":026C
         Style           =   1  'Graphical
         TabIndex        =   6
         Tag             =   "&Delete"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Incluir"
         Height          =   540
         Left            =   120
         Picture         =   "frmStatus.frx":03DE
         Style           =   1  'Graphical
         TabIndex        =   5
         Tag             =   "&Add"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton CmdSair 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Sair"
         Height          =   540
         Left            =   3555
         Picture         =   "frmStatus.frx":04C8
         Style           =   1  'Graphical
         TabIndex        =   4
         Tag             =   "&Update"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton cmddesfaz 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Desfaz"
         Enabled         =   0   'False
         Height          =   540
         Left            =   2865
         Picture         =   "frmStatus.frx":05C2
         Style           =   1  'Graphical
         TabIndex        =   3
         Tag             =   "&Update"
         Top             =   150
         Width           =   615
      End
   End
   Begin VB.TextBox txtStatus 
      Appearance      =   0  'Flat
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
      Left            =   1290
      MaxLength       =   10
      TabIndex        =   0
      Top             =   630
      Width           =   2175
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexStatus 
      Height          =   1545
      Left            =   1290
      TabIndex        =   9
      Top             =   1140
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   2725
      _Version        =   393216
      Rows            =   5
      FixedCols       =   0
      ScrollBars      =   2
      FormatString    =   "Id | Status                         "
   End
   Begin VB.Label lblId 
      BackStyle       =   0  'Transparent
      Caption         =   "Id"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1290
      TabIndex        =   10
      Top             =   30
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lbl_Animal 
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1290
      TabIndex        =   1
      Top             =   300
      Width           =   1380
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iTipoOperacao As Integer
Private lIncluir As Boolean
Private lPrimeiro As Boolean

Private Sub Dados_Colunas()
    
    Call sConectaLocal
    strSql = ""
    strSql = strSql & " SELECT ID_situacao,DESCRICAO FROM TB_situacao"
    Set Rstemp = New ADODB.Recordset
    Rstemp.Open strSql, cnnLocal, 1, 2
    If Rstemp.RecordCount <> 0 Then
        Rstemp.MoveLast
        Rstemp.MoveFirst
        'fmeListaPedidos.Visible = True
        lblId = Rstemp!id_situacao
        txtStatus = Rstemp!descricao
        With Rstemp
            '.MoveLast
            'nItem = .RecordCount
            '.MoveFirst
            MSFlexStatus.Rows = 1
            Do While Not .EOF
               MSFlexStatus.Rows = MSFlexStatus.Rows + 1
               MSFlexStatus.Row = MSFlexStatus.Rows - 1
             
               MSFlexStatus.Col = 0: MSFlexStatus.Text = f_nulo(Rstemp!id_situacao, "")
               MSFlexStatus.Col = 1: MSFlexStatus.Text = f_nulo(Rstemp!descricao, "")
               Rstemp.MoveNext
               
             Loop
             MSFlexStatus.FixedRows = 1
        End With
        
  Else
       MsgBox "Sem registros", vbOKOnly
  End If
  
  Rstemp.Close
  Set Rstemp = Nothing
    
End Sub

Private Sub cmd_Adicionar_Click()
    txtServico.Enabled = True
    txtServico.SetFocus
    txtServico.Text = ""
    cmd_Adicionar.Enabled = False
    cmd_Gravar.Enabled = True
    cmd_Excluir.Enabled = False
    iTipoOperacao = 1

End Sub

Private Sub cmd_Excluir_Click()
    If Len(txtServico.Text) = 0 Or txtServico.Text = "" Then
       MsgBox "Descrição de status inválida. Favor corrigir", vbOKOnly
       txtServico.SetFocus
       Exit Sub
    End If
    
    If MsgBox("Tem certeza que deseja excluir o Status: " & Chr(13) & Chr(10) & _
                            Trim(LstServicos.SelectedItem.ListSubItems.Item(1)), vbYesNo) = vbYes Then
        If fExcluir_Servico() Then
            cmd_Adicionar.Enabled = True
            cmd_Excluir.Enabled = False
            cmd_Gravar.Enabled = False
            Call Dados_Colunas
        Else
            MsgBox "Erro ao excluir o Status: " & Err.Description
        End If
    End If

End Sub

Private Sub cmd_Gravar_Click()
    
    If Len(txtServico.Text) = 0 Or txtServico.Text = "" Then
        MsgBox "Descrição de status inválida. Favor corrigir", vbOKOnly
        txtServico.SetFocus
        Exit Sub
    End If
    
    If fGravar_Servico() Then
        cmd_Adicionar.Enabled = True
        cmd_Gravar.Enabled = False
        cmd_Limpar.Enabled = False
        'cmd_Excluir.Enabled = true
        'LstServicos.ListItems.Clear
        Call Dados_Colunas
        LstServicos.ListItems(1).Selected = True
        txtServico.Text = Trim(LstServicos.SelectedItem.ListSubItems.Item(1))
        'Call cmd_Limpar_Click
    Else
        MsgBox "Erro ao incluir o Status: " & Err.Description
    End If

End Sub

Private Sub cmd_Limpar_Click()
    txtServico.Text = ""
    'txtServico.SetFocus
    cmd_Adicionar.Enabled = False
    cmd_Gravar.Enabled = True

End Sub

Private Sub cmd_Sair_Click()
    Unload Me
End Sub

Private Sub cmdAdd_Click()
   lIncluir = True
   limpa_tela Me
   
   Me.lblId.Caption = ""
   Me.txtStatus.SetFocus
   Me.cmdUpdate.Enabled = True
   Me.cmdDesfaz.Enabled = True
   Me.cmdEditar.Enabled = False
   Me.cmdAdd.Enabled = False
   Me.cmdSair.Enabled = False
   Me.cmdDelete.Enabled = False
End Sub

Private Sub cmdDelete_Click()
  On Error GoTo ErroDelete
    
    'this may produce an error if you delete the last
    'record or the only record in the recordset
    If MsgBox("Deseja realmente apagar este Status ? ", vbYesNo, "Atenção") = vbYes Then
        gSql = "delete from tb_situacao where id_situacao = " & Val(Me.lblId.Caption)
        cnnLocal.Execute gSql
        'Rstemp.Close
        Call Dados_Colunas
        Desabilita Me
        
     End If
     Exit Sub
     
ErroDelete:
     MsgBox "Deu erro na exclusao do Status " & Chr(13) & "Instrucao Sql = '" & _
            gSql & "'  "

End Sub

Private Sub cmdEditar_Click()
   
   Habilita Me
        
   Me.cmdUpdate.Enabled = True
   Me.cmdDesfaz.Enabled = True
   Me.cmdEditar.Enabled = False
   Me.cmdAdd.Enabled = False
   Me.cmdSair.Enabled = False
   Me.cmdDelete.Enabled = False
   Me.txtStatus.SetFocus
   
End Sub



Private Sub CmdSair_Click()
   Unload Me
End Sub

Private Sub cmdUpdate_Click()
   'gRs.Close
   If lIncluir Then
      gSql = "INSERT INTO tb_situacao (descricao,operador,datatual) "
      gSql = gSql & "VALUES ('" & Me.txtStatus.Text & "',"
      gSql = gSql & f_nulo(gncodoperador, 99) & ",'" & Format(Date, "yyyy-mm-dd") & "')"
      cnnLocal.Execute gSql
      lIncluir = False
      
   Else
      gSql = "UPDATE tb_situacao SET descricao = '" & Me.txtStatus.Text & "',"
      gSql = gSql & " operador = " & f_nulo(gncodoperador, 99) & ", datatual = '" & Format(Date, "yyyy-mm-dd") & "'"
      gSql = gSql & " WHERE id_situacao = " & Val(lblId.Caption)
      cnnLocal.Execute gSql
            
   End If
      
   Call Dados_Colunas
   
   Desabilita Me
   
   Me.cmdUpdate.Enabled = False
   Me.cmdDesfaz.Enabled = False
   Me.cmdEditar.Enabled = True
   Me.cmdAdd.Enabled = True
   Me.cmdSair.Enabled = True
   Me.cmdDelete.Enabled = True

End Sub

Private Sub Form_Load()
    'Call Nomes_Colunas
    
    Call Dados_Colunas
    'lstServicos.ListItems = 1
'    If LstServicos.ListItems.Count > 0 Then
'        txtServico.Text = Trim(LstServicos.SelectedItem.ListSubItems.Item(1))
'    End If
End Sub

Private Function fGravar_Servico()
    
    If Len(txtServico.Text) = 0 Or txtServico.Text = "" Then
       MsgBox "Descrição do status inválida. Favor corrigir", vbOKOnly
       txtServico.SetFocus
       Exit Function
    End If
    
    fGravar_Servico = True
    
    On Error GoTo Erro_fGravar_Servico
    
    'ID,DESCRICAO,valor,tempo_est
    If iTipoOperacao = 1 Then
        strSql = "INSERT INTO tb_situacao (DESCRICAO, OPERADOR, DaTATUAL)"
        strSql = strSql + " VALUES( '" & UCase(txtServico.Text) & "',"
        strSql = strSql + f_nulo(gncodoperador, 99) & "','" & Format(Now, "yyyy/mm/dd hh:mm:ss") & "')"
        
    Else
        strSql = "UPDATE tb_situacao SET DESCRICAO = '" & UCase(txtServico.Text) & _
                                          ",OPERADOR = '" & f_nulo(gncodoperador, 99) & _
                                          "', DT_ATUALIZA = '" & Format(Now, "yyyy/mm/dd hh:mm:ss") & _
                                          "' WHERE id_situacao = '" & lblId.Caption & "'"
                                          
    End If
    cnnLocal.Execute strSql
    Exit Function
    
Erro_fGravar_Servico:
    fGravar_Servico = False
End Function

Private Function fExcluir_Servico()
    
    fExcluir_Servico = True
    
    On Error GoTo Erro_fExcluir_Servico
    
    strSql = "DELETE from tb_situacao WHERE ID = " & lblId.Caption
    cnnLocal.Execute strSql
    
    Exit Function
Erro_fExcluir_Servico:
    fExcluir_Servico = False
End Function

Private Sub MSFlexStatus_Click()
Dim oldrow As Long
  
  oldrow = MSFlexStatus.Row
  
  MSFlexStatus.Row = 0
  
  With MSFlexStatus
    .Redraw = False
    Do While True
       .Row = .Row + 1
       For ix = 0 To .Cols - 1
           .Col = ix: .CellBackColor = vbWhite
       Next
       If .Row = .Rows - 1 Then
          Exit Do
       End If
    Loop
    .Redraw = True
    
    .Row = oldrow
    
    .Col = 0:   lblId.Caption = .Text: .CellBackColor = vbYellow
    .Col = 1:   txtStatus.Text = .Text: .CellBackColor = vbYellow
    .TopRow = .Row
   
End With

 Desabilita Me
   Me.cmdUpdate.Enabled = False
   Me.cmdDesfaz.Enabled = False
   Me.cmdEditar.Enabled = True
   Me.cmdAdd.Enabled = True
   Me.cmdSair.Enabled = True
End Sub

Private Sub txtStatus_GotFocus()
     Call SelText(txtStatus)
End Sub

Private Sub txtStatus_KeyPress(KeyAscii As Integer)
    'Char = Chr(KeyAscii)
    'KeyAscii = Asc(UCase(Char))
    If KeyAscii = 13 Then
        If Len(Trim(txtStatus.Text)) = 0 Then
            MsgBox "Obrigatório Informar Descrição do Status.", vbInformation, "Aviso"
            txtServico.SetFocus
            Exit Sub
        End If

        Sendkeys "{tab}"
    End If
End Sub
