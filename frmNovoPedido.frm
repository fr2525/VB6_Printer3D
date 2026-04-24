VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Begin VB.Form frmNovoPedido 
   BackColor       =   &H00808000&
   BorderStyle     =   0  'None
   Caption         =   "Pedido de Impressão "
   ClientHeight    =   7785
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12180
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7785
   ScaleWidth      =   12180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   840
      Left            =   4920
      TabIndex        =   21
      Top             =   6750
      Width           =   2145
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Salvar"
         Height          =   540
         Left            =   75
         Picture         =   "frmNovoPedido.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   24
         Tag             =   "&Update"
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdSair 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Sair"
         CausesValidation=   0   'False
         Height          =   540
         Left            =   1440
         Picture         =   "frmNovoPedido.frx":00FA
         Style           =   1  'Graphical
         TabIndex        =   23
         Tag             =   "&Update"
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton cmddesfaz 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Desfaz"
         Height          =   540
         Left            =   750
         Picture         =   "frmNovoPedido.frx":01F4
         Style           =   1  'Graphical
         TabIndex        =   22
         Tag             =   "&Update"
         Top             =   180
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Impressoras e filamentos escolhidos "
      Height          =   2265
      Left            =   255
      TabIndex        =   19
      Top             =   4335
      Width           =   11655
      Begin ComctlLib.TreeView Tvw1 
         Height          =   1830
         Left            =   225
         TabIndex        =   20
         Top             =   315
         Width           =   11205
         _ExtentX        =   19764
         _ExtentY        =   3228
         _Version        =   327682
         Style           =   7
         Appearance      =   0
         Enabled         =   0   'False
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF8080&
      Height          =   2670
      Left            =   240
      TabIndex        =   7
      Top             =   1575
      Width           =   11685
      Begin VB.TextBox txQtPedida 
         Height          =   405
         Left            =   10920
         MaxLength       =   10
         TabIndex        =   18
         Text            =   "txqtpedida"
         Top             =   240
         Visible         =   0   'False
         Width           =   675
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   2160
         Left            =   180
         TabIndex        =   15
         Top             =   345
         Width           =   3810
         _ExtentX        =   6720
         _ExtentY        =   3810
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         ScrollBars      =   2
         Appearance      =   0
         FormatString    =   "    | Marca - Modelo                                     | Rolos"
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   2115
         Left            =   4260
         TabIndex        =   16
         Top             =   360
         Width           =   7260
         _ExtentX        =   12806
         _ExtentY        =   3731
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         SelectionMode   =   1
         Appearance      =   0
         FormatString    =   " Marca                                              | Modelo                                         | Qt.rolos    |      Usar "
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Filamento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   270
         Left            =   4305
         TabIndex        =   25
         Top             =   150
         Width           =   1065
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Impressora"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   270
         Left            =   225
         TabIndex        =   17
         Top             =   135
         Width           =   1065
      End
   End
   Begin VB.TextBox txtDataPrevista 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   10800
      TabIndex        =   6
      Top             =   1170
      Width           =   1095
   End
   Begin VB.TextBox txtDataPedido 
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   3
      EndProperty
      Height          =   300
      Left            =   6105
      MaxLength       =   10
      TabIndex        =   5
      Top             =   1170
      Width           =   1095
   End
   Begin VB.TextBox txtTotal 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   300
      Left            =   4020
      MaxLength       =   20
      TabIndex        =   4
      Top             =   1170
      Width           =   1050
   End
   Begin VB.TextBox txtQtde 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   735
      MaxLength       =   6
      TabIndex        =   2
      Top             =   1170
      Width           =   555
   End
   Begin VB.ComboBox cmbCliente 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   765
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   735
      Width           =   4335
   End
   Begin VB.TextBox txtProjeto 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   6105
      MaxLength       =   100
      TabIndex        =   1
      Top             =   720
      Width           =   5790
   End
   Begin VB.TextBox txtUnitario 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   2235
      MaxLength       =   20
      TabIndex        =   3
      Top             =   1170
      Width           =   840
   End
   Begin VB.Label lblPedido 
      BackStyle       =   0  'Transparent
      Caption         =   "Novo Pedido"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   435
      Left            =   4860
      TabIndex        =   26
      Top             =   120
      Width           =   2100
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dta.Prevista:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   9825
      TabIndex        =   14
      Top             =   1245
      Width           =   915
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dta.Pedido:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5220
      TabIndex        =   13
      Top             =   1200
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Qtde."
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   270
      TabIndex        =   12
      Top             =   1245
      Width           =   390
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   135
      TabIndex        =   11
      Tag             =   "NOME:"
      Top             =   780
      Width           =   525
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Projeto:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   5130
      TabIndex        =   10
      Tag             =   "SALARIO:"
      Top             =   825
      Width           =   855
   End
   Begin VB.Label lblcgc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vlr.unitário"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1395
      TabIndex        =   9
      Top             =   1230
      Width           =   735
   End
   Begin VB.Label LblTelefone 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Vlr.Total:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3225
      TabIndex        =   8
      Top             =   1200
      Width           =   705
   End
End
Attribute VB_Name = "frmNovoPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const Marcado = "þ"
Const Desmarcado = "q"
Const Excluir = "X"
Const COL_NOME = 1
Const COL_STATUS = 3
Const COL_EDITAVEL = 4

'Dim WithEvents TxtMarca As TextBox
'Dim WithEvents TxtTipo As TextBox
'Dim WithEvents txtCor As TextBox
'Dim WithEvents txtEstoque As TextBox
'Dim WithEvents txtQtde As TextBox
Dim nodx As Node

Private Sub cmbCliente_Change()
    Me.txtProjeto.SetFocus
End Sub

Private Sub cmbCliente_Click()
   If cmbCliente.ListIndex <> -1 Then
        Sendkeys "{TAB}"
    End If
End Sub

Private Sub cmbCliente_Validate(Cancel As Boolean)
    If cmbCliente.ListIndex = -1 Then
        MsgBox "Favor escolher um cliente!"
        KeepFocus = True ' Mantém o cursor no ComboBox até que um item seja escolhido
    End If

End Sub

Private Sub Form_Load()
  If gNumPedido = 0 Then
    Call CarregarClientes
    Call carregarImpressoras
    Me.txtDataPedido = Format(Now(), "dd/mm/yyyy")
  Else
    'vai carregar o pedido que está selecionado
    Call CarregarPedido
  End If
End Sub
 
Private Sub MostrarCheckbox(iRow As Integer, iCol As Integer)
    With MSFlexGrid1
        If .TextMatrix(iRow, 0) = Desmarcado Then
            .TextMatrix(iRow, 0) = Marcado
             coluna_ant = MSFlexGrid1.col
             MSFlexGrid1.col = 2
             Call CarregarFilamentos(MSFlexGrid1.Text)
             MSFlexGrid1.col = coluna_ant
        Else
            .TextMatrix(iRow, 0) = Desmarcado
            MSFlexGrid2.Clear
            MSFlexGrid2.Redraw = False
            MSFlexGrid2.Redraw = True
        End If
    End With
    Call CriaTreeView
End Sub

Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Or KeyAscii = 32 Then 'Enter ou Barra de Espaço
     With MSFlexGrid1
        Call MostrarCheckbox(.row, .col)
     End With
 End If
End Sub

Private Sub MSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = 1 Then
        With MSFlexGrid1
            If .MouseRow <> 0 And .MouseCol = 0 Then
                Call MostrarCheckbox(.MouseRow, .MouseCol)
            
            End If
        End With
    End If
End Sub

Private Sub MSFlexGrid2_DblClick()
    If MSFlexGrid2.row = 0 Then
        Exit Sub
    End If
    MSFlexGrid2.col = 4
    EditarCelula
    
End Sub

Private Sub MSFlexGrid2_KeyPress(KeyAscii As Integer)
' If KeyAscii = 13 Or KeyAscii = 32 Then 'Enter ou Barra de Espaço
'        With MSFlexGrid2
'            Call MostrarCheckbox2(.row, .Col)
'            'Call gravar_filamentos
'        End With
'    End If
'
'Select Case KeyAscii
'Case vbKeyReturn, vbKeyTab
''move para a proxima celula.
'
'With MSFlexGrid2
'
'  If .Col + 1 <= .Cols - 1 Then
'     .Col = .Col + 1
'  Else
'     If .row + 1 <= .Rows - 1 Then
'         .row = .row + 1
'         .Col = 0
'     Else
'         .row = 1
'         .Col = 0
'     End If
'  End If
'End With
'
'Case vbKeyBack
'
'   With MSFlexGrid2
'   'remove o ultimo caractere
'      If Len(.Text) Then
'         .Text = Left(.Text, Len(.Text) - 1)
'      End If
'   End With
'
'Case Is < 32
'
'Case Else
'    With MSFlexGrid2
'       .Text = .Text & Chr(KeyAscii)
'    End With
'End Select

End Sub

Private Sub MSFlexGrid2_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
'    If Button = 1 Then
'        With MSFlexGrid2
'            If .MouseRow <> 0 And .MouseCol = 0 Then
'                Call MostrarCheckbox2(.MouseRow, .MouseCol)
'                'Call gravar_filamentos
'            End If
'        End With
'    End If
End Sub
' Exemplo de manipulação de evento para a nova TextBox
Private Sub NewTextBox_Change()
    ' Você pode adicionar código aqui para responder a alterações de texto
    ' Debug.Print NewTextBox.Text
End Sub
Private Sub processaFilamentos(id_impressora)
    Dim iRow2
    With MSFlexGrid2
       iRow2 = 1
       .row = iRow2
       Do While iRow2 < .Rows
          If .TextMatrix(iRow2, 0) = Marcado Then
              'insert pedxfilamento
          End If
          iRow2 = iRow2 + 1
       Loop
     End With
     'insert pedximpressora
End Sub

Private Sub CarregarClientes()

    Call sConectaLocal
  
    strSql = ""
    strSql = strSql & "select id,nome from tb_clientes order by nome asc"
    Set Rstemp = New ADODB.Recordset
    Rstemp.Open strSql, cnnLocal, 1, 2
    Rstemp.MoveFirst
    Do While Not Rstemp.EOF
              
        Me.cmbCliente.AddItem (Rstemp!nome)
        Me.cmbCliente.ItemData(Me.cmbCliente.NewIndex) = Rstemp!id
        Rstemp.MoveNext
    Loop
    Rstemp.Close
    Set Rstemp = Nothing
    
    For i = 0 To cmbCliente.ListCount - 1
        If cmbCliente.ItemData(i) = gNumPedido Then
            cmbCliente.ListIndex = i
            Exit For ' Para o loop ao encontrar
        End If
    Next i
    
End Sub

Private Sub carregarImpressoras()

    Call sConectaLocal

    strSql = ""
    strSql = strSql & "SELECT id, marca, modelo, qtrolos FROM TB_impressoras "
    strSql = strSql & "WHERE id NOT in (select id_impressora FROM tb_pedximpres) "
    strSql = strSql & " order by marca"
    
    RsTemp1.Open strSql, cnnLocal, 1, 2
    RsTemp1.MoveFirst
    MSFlexGrid1.Redraw = False
    MSFlexGrid1.Rows = 1
    MSFlexGrid1.Cols = 4
    MSFlexGrid1.ColWidth(0) = 400
    MSFlexGrid1.ColWidth(1) = 2800
    MSFlexGrid1.ColWidth(2) = 490
    MSFlexGrid1.ColWidth(3) = 0
    'y = 1
    'fmeListaPedidos.Visible = True
    Do While Not RsTemp1.EOF
           
       With MSFlexGrid1
         .Rows = .Rows + 1
         .row = .Rows - 1
         .col = 0
         .CellFontName = "Wingdings"
         .CellFontSize = 14
         .CellAlignment = flexAlignCenterCenter
         .Text = Desmarcado
         .col = 1
         .CellFontName = "arial"
         .Text = Left(RsTemp1!Marca, 20) & " - " & Left(RsTemp1!modelo, 20)
         .col = 2
         .CellFontName = "arial"
         .Text = RsTemp1!qtrolos
         .col = 3
         .CellFontName = "arial"
         .Text = RsTemp1!id
    
      End With
      RsTemp1.MoveNext
        
    Loop
    MSFlexGrid1.FixedRows = 1
    MSFlexGrid1.Redraw = True
    
    RsTemp1.Close
    Set RsTemp1 = Nothing
    
End Sub

Sub CarregarFilamentos(idImpressora As Long)
Dim rs As New ADODB.Recordset
Dim sql As String

sql = "SELECT id, marca,tipo,cor,qtde_estoque FROM tb_filamentos"

rs.Open sql, cnnLocal, 1, 2
With MSFlexGrid2
    .Clear
    .Redraw = False
    .Rows = 1
    .Cols = 6
    .ColWidth(0) = 2500
    .ColWidth(1) = 2000
    .ColWidth(2) = 1000
    .ColWidth(3) = 800
    .ColWidth(4) = 600
    .ColWidth(5) = 0
    .TextMatrix(0, 0) = "Marca"
    .TextMatrix(0, 1) = "Tipo"
    .TextMatrix(0, 2) = "Cor"
    .TextMatrix(0, 3) = "Estoque"
    .TextMatrix(0, 4) = "Pedido"
    .TextMatrix(0, 5) = "Id"
End With

Do While Not rs.EOF
  'Preenche a grid dos filamentos e abre a janela dos filamentos
  With MSFlexGrid2
     .Rows = .Rows + 1
     .row = .Rows - 1
     .col = 0
     .Text = rs!Marca
     .col = 1
     .Text = rs!tipo
     .col = 2
     .Text = rs!cor
     .col = 3
     .Text = rs!qtde_Estoque
     .col = 4
     .Text = 0
     .col = 5
     .Text = rs!id
     
  End With
  rs.MoveNext
Loop
MSFlexGrid2.Redraw = True
rs.Close

End Sub
Sub CarregarPedido()

    Call CarregarClientes
          
    Call sConectaLocal
    strSql = ""
    strSql = strSql & " SELECT A.id_venda,a.id_cliente,a.descricao,a.preco,a.quantidade,a.total_venda,"
    strSql = strSql & " a.situacao,a.dataCompra,a.dataInicioProd,a.dataPrevisao,a.dataFinaliza,a.dataEntrega,"
    strSql = strSql & " b.nome, c.descricao as desc_situacao "
    strSql = strSql & " FROM tb_pedidos A , TB_clientes B, tb_situacao C "
    strSql = strSql & " WHERE A.id_cliente = b.id "
    strSql = strSql & " AND A.situacao = c.id_situacao "
    strSql = strSql & " AND A.id_venda = " & gNumPedido
        
    gRs.Open strSql, cnnLocal, adOpenKeyset
    If gRs.EOF Then
        MsgBox "Pedido não encontrado"
        Return
       
    Else
        Me.txtProjeto.Text = gRs!descricao
        Me.txtDataPedido = gRs!datacompra
        Me.txtDataPrevista = gRs!dataprevisao
        Me.txtQtde = gRs!quantidade
        Me.txtTotal = gRs!total_venda
        Me.txtUnitario = gRs!preco
        
        'gRs.MoveLast
        'gRs.MoveFirst
    
          
     End If
     If gRs.State = adStateOpen Then
        gRs.Close
     End If
     Set gRs = Nothing
    
End Sub
Sub EditarCelula()
    With Me.txQtPedida
        .Text = MSFlexGrid2.Text
        .Move MSFlexGrid2.Left + MSFlexGrid2.CellLeft, _
              MSFlexGrid2.Top + MSFlexGrid2.CellTop, _
              MSFlexGrid2.CellWidth, _
              MSFlexGrid2.CellHeight
        .Visible = True
        .ZOrder
        .SetFocus
    End With
End Sub

Private Sub txQtPedida_GotFocus()
    Call SelText(txQtPedida)
End Sub

Private Sub txQtPedida_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
      MSFlexGrid2.Text = txQtPedida.Text
      MSFlexGrid2.col = 3
      MSFlexGrid2.Text = MSFlexGrid2.Text - txQtPedida.Text
      txQtPedida.Visible = False
      Call CriaTreeView
  ElseIf KeyCode = vbKeyEscape Then
      txQtPedida.Visible = False
  End If
End Sub

Private Sub CriaTreeView()

'limpa qualquer nó criado
Tvw1.Nodes.Clear

'Inserindo o nó raiz
Set nodx = Tvw1.Nodes.Add(, , "Root", "Nó Raiz")

'Expandindo o nó raiz para exibir as ramificações
'nodx.ExpandedImage = "Aberto"
nodx.Expanded = True

If MSFlexGrid1.Rows = 1 Then
    Exit Sub
End If

linha1 = 1

Do While linha1 < MSFlexGrid1.Rows
    'Inserindo as impressoras
    MSFlexGrid1.row = linha1
    MSFlexGrid1.col = 0
    
    If MSFlexGrid1.Text = Marcado Then
        MSFlexGrid1.col = 1
        'Set nodx = Tvw1.Nodes.Add(, , "Root", MSFlexGrid1.Text)
        Child = "Child" & linha1
        Set nodx = Tvw1.Nodes.Add("Root", tvwChild, Child, MSFlexGrid1.Text)
        
        'Expandindo o nó para exibir as ramificações
        'nodx.ExpandedImage = "Aberto"
        nodx.Expanded = True
        
        'Pegar os filamentos
        
        noFilho = 1
        linha2 = 1
        Do While linha2 < MSFlexGrid2.Rows
           MSFlexGrid2.row = linha2
           Child_A = "Child_A" & noFilho
           'Criando dois nós filhos subordinado a ao primeiro nó filho da raiz - Child1
           MSFlexGrid2.col = 4
           If Val(MSFlexGrid2.Text) > 0 Then
              MSFlexGrid2.col = 0
              strfilamento = MSFlexGrid2.Text
              MSFlexGrid2.col = 1
              strfilamento = strfilamento & " - " & MSFlexGrid2.Text
              MSFlexGrid2.col = 2
              strfilamento = strfilamento & " - " & MSFlexGrid2.Text
              
              Set nodx = Tvw1.Nodes.Add(Child, tvwChild, Child_A, strfilamento)
              noFilho = noFilho + 1
           End If
           linha2 = linha2 + 1
        Loop
    End If
    linha1 = linha1 + 1
Loop

End Sub

Private Sub CmdSair_Click()
    Unload Me
End Sub

Private Sub cmdUpdate_Click()

Dim novoID
 On Error GoTo TrataErro

'***** Inicio da transação para gravação em todas as tabelas envolvidas ***
If Not ConsistePedido() Then
    Exit Sub
End If

Call sConectaLocal

'Salvar o pedido primeiro para pegar o numero dele e atualizar as outras tabelas
' O comando OUTPUT inserted.* traz todos os campos do registro recém-criado
         
'***** Inicio da transação para gravação em todas as tabelas envolvidas ***

cnnLocal.BeginTrans
         
Set rs = New ADODB.Recordset
         
strSql = "INSERT INTO tb_pedidos (id_cliente,descricao,preco,quantidade,total_venda,situacao,dataCompra,dataprevisao)" & _
        " VALUES (" & cmbCliente.ItemData(cmbCliente.ListIndex) & _
         ",'" & txtProjeto.Text & "'," & SoNumero(Me.txtUnitario.Text) & "," & Me.txtQtde.Text & "," & _
         SoNumero(Me.txtTotal.Text) & "," & 1 & ",'" & Format(Me.txtDataPedido.Text, "yyyy-mm-dd") & "','" & Format(Me.txtDataPrevista.Text, "yyyy-mm-dd") & "')"

Set rs = cnnLocal.Execute(strSql)

strSql = "SELECT id_venda FROM tb_pedidos ORDER BY ID_venda desc limit 1"
Set rs = cnnLocal.Execute(strSql)

If Not rs.EOF Then
    novoID = rs.Fields("ID_venda").Value
End If

' 2. Executa as atualizações

With MSFlexGrid1
    iRow = 1
    .row = iRow
    Do While iRow < .Rows
       If .TextMatrix(iRow, 0) = Marcado Then
          strSql = "UPDATE tb_impressoras SET ocupada = 'S' WHERE id = " & .TextMatrix(iRow, 3)
          Set rs = cnnLocal.Execute(strSql)
          strSql = "INSERT INTO tb_pedximpres (id_pedido,id_impressora) VALUES (" & novoID & "," & .TextMatrix(iRow, 3) & ")"
          Set rs = cnnLocal.Execute(strSql)
       End If
       iRow = iRow + 1
    Loop
 End With
  
With MSFlexGrid2
    iRow = 1
    .row = iRow
    Do While iRow < .Rows
       If .TextMatrix(iRow, 4) > 0 Then
          strSql = "UPDATE tb_filamentos SET qtde_estoque = " & .TextMatrix(iRow, 4) & " WHERE id = " & .TextMatrix(iRow, 5)
          Set rs = cnnLocal.Execute(strSql)
          strSql = "INSERT INTO tb_pedxfilamento (id_pedido,id_filamento) VALUES (" & novoID & "," & .TextMatrix(iRow, 5) & ")"
          Set rs = cnnLocal.Execute(strSql)
       End If
       iRow = iRow + 1
    Loop
 End With

' 3. Se não houve erro, confirma as alterações permanentemente
cnnLocal.CommitTrans
MsgBox "Operação concluída com sucesso!", vbInformation
Unload Me

Sair:
    If cnnLocal.State = adStateOpen Then cnnLocal.Close
    Set cnnLocal = Nothing
    Exit Sub

TrataErro:
    ' 4. Em caso de erro, desfaz todas as alterações desde o BeginTrans
    If Not cnnLocal Is Nothing Then
        If cnnLocal.State = adStateOpen Then cnnLocal.RollbackTrans
    End If
    MsgBox "Erro na transação: " & Err.Description, vbCritical
    Resume Sair

'*****
End Sub

Private Sub txQtPedida_KeyPress(KeyAscii As Integer)
    Dim separador As String
    separador = "," ' Ou "." dependendo da região
    
    ' Permite números, backspace e a vírgula decimal
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = Asc(separador)) Then
        KeyAscii = 0 ' Bloqueia a tecla
    End If
    
    ' Evita mais de uma vírgula
    If KeyAscii = Asc(separador) And InStr(Text1.Text, separador) > 0 Then
        KeyAscii = 0
    End If
End Sub

End Sub

Private Sub txtDataPedido_GotFocus()
    Call SelText(txtDataPedido)
End Sub

Private Sub txtDataPedido_KeyPress(KeyAscii As Integer)
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
    If KeyAscii = 13 Then
        Sendkeys "{tab}"
    End If
    
     ' Permite apenas números e Backspace
    If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
        KeyAscii = 0
        Exit Sub
    End If

    ' Adiciona as barras automaticamente nas posições 3 e 6
    Select Case Len(txtDataPedido.Text)
        Case 2, 5
            If KeyAscii <> 8 Then ' Se não for Backspace
                txtDataPedido.Text = txtDataPedido.Text & "/"
                txtDataPedido.SelStart = Len(txtDataPedido.Text)
            End If
    End Select

End Sub

Private Sub txtDataPedido_LostFocus()
    'txtDataPedido.Text = Format(txtDataPedido.Text, "dd/mm/yyyy")
End Sub

Private Sub txtDataPedido_Validate(Cancel As Boolean)
  If Not f_ValidaData(txtDataPedido.Text) Then
      MsgBox "Data Inválida.", vbInformation, "Atenção"
      txtDataPedido.SetFocus
      Cancel = True
   End If
    
End Sub

Private Sub txtDataPrevista_KeyPress(KeyAscii As Integer)
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
    If KeyAscii = 13 Then
        Sendkeys "{tab}"
    End If
    
     ' Permite apenas números e Backspace
    If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
        KeyAscii = 0
        Exit Sub
    End If

    ' Adiciona as barras automaticamente nas posições 3 e 6
    Select Case Len(txtDataPrevista.Text)
        Case 2, 5
            If KeyAscii <> 8 Then ' Se não for Backspace
                txtDataPrevista.Text = txtDataPrevista.Text & "/"
                txtDataPrevista.SelStart = Len(txtDataPrevista.Text)
            End If
    End Select

End Sub

Private Sub txtDataPrevista_Validate(Cancel As Boolean)
  If Not f_ValidaData(txtDataPrevista.Text) Then
      MsgBox "Data Inválida.", vbInformation, "Atenção"
      txtDataPrevista.SetFocus
      Cancel = True
   End If

End Sub

Private Sub txtProjeto_GotFocus()
    Call SelText(txtProjeto)
End Sub

Private Sub txtProjeto_KeyPress(KeyAscii As Integer)
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
    If KeyAscii = 13 Then
        Sendkeys "{tab}"
    End If
End Sub

Private Sub txtQtde_KeyPress(KeyAscii As Integer)
  Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
    If KeyAscii = 13 Then
        Sendkeys "{tab}"
    End If
End Sub

Private Sub txtQtde_LostFocus()
   ' Verifica se o campo não está vazio
    If txtQtde.Text <> "" Then
        ' Converte o texto para valor numérico e formata como moeda
        ' A função FormatCurrency usa as configurações regionais do sistema
    End If

End Sub

Private Sub txtUnitario_KeyPress(KeyAscii As Integer)
  Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
    If KeyAscii = 13 Then
        Sendkeys "{tab}"
    End If

End Sub

Private Sub txtUnitario_LostFocus()
   ' Verifica se o campo não está vazio
    If txtUnitario.Text <> "" Then
        ' Converte o texto para valor numérico e formata como moeda
        ' A função FormatCurrency usa as configurações regionais do sistema
        txtTotal.Text = FormatCurrency(CDbl(Val(txtUnitario.Text) * Val(txtQtde.Text)))
        txtUnitario.Text = FormatCurrency(CDbl(txtUnitario.Text))

    End If
    
End Sub

Private Function ConsistePedido()

Dim nMarcados

If Len(Trim(txtProjeto.Text)) = 0 Then
    MsgBox "Obrigatório Informar Descrição do projeto.", vbInformation, "Aviso"
    txtProjeto.SetFocus
    ConsistePedido = False
    Exit Function
End If
If Len(Trim(txtQtde.Text)) = 0 Then
    MsgBox "Obrigatório Informar uma quantidade.", vbInformation, "Aviso"
    txtQtde.SetFocus
    ConsistePedido = False
    Exit Function
End If
If Len(Trim(Me.txtUnitario.Text)) = 0 Then
    MsgBox "Obrigatório Informar Valor unitário.", vbInformation, "Aviso"
    txtUnitario.SetFocus
    ConsistePedido = False
    Exit Function
End If
If Len(Trim(Me.txtDataPedido.Text)) = 0 Then
    MsgBox "Obrigatório Informar a data do pedido.", vbInformation, "Aviso"
    txtDataPedido.SetFocus
    ConsistePedido = False
    Exit Function
End If
If Len(Trim(Me.txtDataPrevista.Text)) = 0 Then
    MsgBox "Obrigatório Informar a data prevista para entrega do pedido.", vbInformation, "Aviso"
    txtDataPrevista.SetFocus
    ConsistePedido = False
    Exit Function
End If

'**** Consistindo as impressoras *****
iRow = 0
nMarcados = 0
With MSFlexGrid1
    Do While iRow < MSFlexGrid1.Rows
      If .TextMatrix(iRow, 0) = Marcado Then
          nMarcados = nMarcados + 1
      End If
      iRow = iRow + 1
    Loop
End With

If nMarcados = 0 Then
    MsgBox "Sem impressoras escolhidas para o pedido.", vbInformation, "Aviso"
    MSFlexGrid1.SetFocus
    ConsistePedido = False
    Exit Function
End If

'**** Consistindo os filamentos *****
iRow = 0
nMarcados = 0
With MSFlexGrid2
    Do While iRow < .Rows
      If Val(.TextMatrix(iRow, 4)) > 0 Then
          nMarcados = nMarcados + 1
      End If
      iRow = iRow + 1
    Loop
End With

If nMarcados = 0 Then
    MsgBox "Sem filamentos escolhidos para o pedido.", vbInformation, "Aviso"
    MSFlexGrid2.SetFocus
    ConsistePedido = False
    Exit Function
End If

ConsistePedido = True

End Function
