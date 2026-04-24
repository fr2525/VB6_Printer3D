VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form frmPedidos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impressão 3D"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12210
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPedidos.frx":0000
   ScaleHeight     =   7965
   ScaleWidth      =   12210
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdParametros 
      BackColor       =   &H00FFFFC0&
      Caption         =   "&Parâmetros"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Sitka Banner"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   2070
      Picture         =   "frmPedidos.frx":6758E
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   870
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.CommandButton cmdUsuarios 
      BackColor       =   &H00FFFFC0&
      Caption         =   "&Usuários"
      Default         =   -1  'True
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Sitka Banner"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   3270
      Picture         =   "frmPedidos.frx":67810
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   870
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.Frame frmSituacao 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Caption         =   "Atualização"
      Enabled         =   0   'False
      Height          =   1290
      Left            =   4050
      TabIndex        =   8
      Top             =   2865
      Visible         =   0   'False
      Width           =   4080
      Begin VB.CommandButton cmdVoltar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Voltar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   3015
         Picture         =   "frmPedidos.frx":6790A
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   675
         Width           =   885
      End
      Begin VB.TextBox txtNovaData 
         Height          =   285
         Left            =   855
         TabIndex        =   11
         Top             =   750
         Width           =   1395
      End
      Begin VB.CommandButton cmdSituacao 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   3045
         Picture         =   "frmPedidos.frx":67A04
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   60
         Width           =   840
      End
      Begin VB.ComboBox cmbSituacao 
         Height          =   315
         Left            =   225
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Data"
         Height          =   240
         Left            =   240
         TabIndex        =   10
         Top             =   795
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdSair 
      BackColor       =   &H00C0E0FF&
      Cancel          =   -1  'True
      Caption         =   "Sai&r"
      BeginProperty Font 
         Name            =   "Sitka Banner"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   10620
      Picture         =   "frmPedidos.frx":67F36
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   870
      Width           =   1200
   End
   Begin VB.CommandButton cmdStatus 
      BackColor       =   &H00FFFFC0&
      Caption         =   "&Status"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Sitka Banner"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   4470
      Picture         =   "frmPedidos.frx":68030
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   870
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.CommandButton cmdImpressoras 
      BackColor       =   &H00FFFFC0&
      Caption         =   "&Impressoras"
      BeginProperty Font 
         Name            =   "Sitka Banner"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   9240
      Picture         =   "frmPedidos.frx":6812A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   870
      Width           =   1110
   End
   Begin VB.CommandButton cmdFilamentos 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Fila&mentos"
      BeginProperty Font 
         Name            =   "Sitka Banner"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   8040
      Picture         =   "frmPedidos.frx":68D6C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   870
      Width           =   1110
   End
   Begin VB.CommandButton cmdFornece 
      BackColor       =   &H00FFFFC0&
      Caption         =   "&Fornecedores"
      BeginProperty Font 
         Name            =   "Sitka Banner"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   6870
      Picture         =   "frmPedidos.frx":699AE
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   870
      Width           =   1110
   End
   Begin VB.CommandButton cmdClientes 
      BackColor       =   &H00FFFFC0&
      Caption         =   "&Clientes"
      BeginProperty Font 
         Name            =   "Sitka Banner"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   5670
      Picture         =   "frmPedidos.frx":6A5F0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   870
      Width           =   1110
   End
   Begin VB.CommandButton cmdNovoPed 
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Novo Pedido"
      BeginProperty Font 
         Name            =   "Sitka Banner"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   120
      Picture         =   "frmPedidos.frx":6AEFA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   870
      Width           =   1200
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexPedidos 
      Height          =   5355
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "   "
      Top             =   2115
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   9446
      _Version        =   393216
      Rows            =   5
      Cols            =   10
      FixedCols       =   0
      FormatString    =   $"frmPedidos.frx":6AFE4
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gestão de Impressão 3D"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   510
      Left            =   3795
      TabIndex        =   12
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "frmPedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim aLista_Horario(48) As String
Public gnNivel As Integer
Public gcUsuario As String

Dim oldCol, situacao, lDataCompra, lDataInicio, lDataFinaliza, lDataEntrega
Dim lDataAlterada
Dim oldsituacao As Integer

Private Sub MontaAtendimentos()

    gNumPedido = 0
    Dim indCol
    Call sConectaLocal
    strSql = ""
    strSql = strSql & " SELECT A.*, b.nome, c.descricao as desc_situacao FROM tb_pedidos A , TB_clientes B, tb_situacao C "
    strSql = strSql & " WHERE A.id_cliente = b.id "
    strSql = strSql & " AND A.situacao = c.id_situacao "
    
    gRs.Open strSql, cnnLocal, adOpenKeyset
    If gRs.EOF Then
        With MSFlexPedidos
            .row = 0
            .Rows = .Rows + 1
            .row = .Rows - 1
            .ColAlignment(-1) = flexAlignLeftCenter
            For indCol = 0 To 10
                .col = indCol: .Text = ""
            Next
            .FixedRows = 1
        End With
      
    Else
        gRs.MoveLast
        gRs.MoveFirst
    
        MSFlexPedidos.row = 0
        
       '.MoveLast
       'nItem = .RecordCount
       '.MoveFirst
       With MSFlexPedidos
          .Rows = 1
          .ColWidth(9) = 0
          Do While Not gRs.EOF
              .Rows = .Rows + 1
              .row = .Rows - 1
        
              .col = 0: .Text = f_nulo(gRs!nome, "")
              .col = 1: .Text = f_nulo(gRs!descricao, "")
              .col = 2: .Text = f_nulo(Format(gRs!total_venda, "0.00"), "")
              .col = 3: .Text = f_nulo(gRs!desc_situacao, "")
              .col = 4: .Text = f_nulo(Format(gRs!datacompra, "DD/MM/YYYY"), "")
              .col = 5: .Text = f_nulo(Format(gRs!dataInicioProd, "DD/MM/YYYY"), "")
              .col = 6: .Text = f_nulo(Format(gRs!dataprevisao, "DD/MM/YYYY"), "")
              .col = 7: .Text = f_nulo(Format(gRs!dataFinaliza, "DD/MM/YYYY"), "")
              .col = 8: .Text = f_nulo(Format(gRs!dataEntrega, "DD/MM/YYYY"), "")
              .col = 9: .Text = f_nulo(gRs!id_venda, 0)
              gRs.MoveNext
          Loop
          .FixedRows = 1
          
      End With
   End If
   If gRs.State = adStateOpen Then
      gRs.Close
  End If
  Set gRs = Nothing
   
End Sub

Private Sub cmd_Adicionar_Click()

End Sub

Private Sub cmd_clientes_Click()
    FrmClientes.Show vbModal
End Sub

Private Sub cmd_Sair_Click()
   ' Call closeDB
    Unload Me
End Sub

Private Sub cmd_Serviços_Click()
    Frmfilamento.Show vbModal
End Sub


Private Sub cmdClientes_Click()
    FrmClientes.Show vbModal
End Sub

Private Sub cmdFilamentos_Click()
    Frmfilamento.Show vbModal
End Sub

Private Sub cmdFornece_Click()
    frmfornec.Show vbModal
End Sub

Private Sub cmdImpressoras_Click()
    Frmimpressora.Show vbModal
End Sub

Private Sub cmdNovoPed_Click()
    frmNovoPedido.Show vbModal
    Call MontaAtendimentos
    
End Sub

Private Sub cmdParametros_Click()
    frmlojas.Show vbModal
    
End Sub

Private Sub CmdSair_Click()
'   If MsgBox("Vai embora mesmo ?", 32 + 4 + 256) <> 6 Then
'      Cancel = True
'   Else
      End
'   End If
End Sub

Private Sub cmdSituacao_Click()
    Call suConsisteSituacao
    'Call suGravaSituacao
'    Me.frmSituacao.Visible = False
'    Me.frmSituacao.Enabled = False
End Sub

Private Sub cmdStatus_Click()
   frmStatus.Show vbModal
End Sub

Private Sub cmdUsuarios_Click()
    frmUsuarios.Show vbModal
End Sub

Private Sub cmdVoltar_Click()
    Me.frmSituacao.Visible = False
    Me.frmSituacao.Enabled = False
    
End Sub

Private Sub Form_Load()
   Dim i As Integer
   Dim sHora As String
   
    ' Define o tamanho do formulário para 800x600 pixels (convertidos para twips)
    Me.Width = 800 * Screen.TwipsPerPixelX
    Me.Height = 600 * Screen.TwipsPerPixelY
    
    ' Centraliza o formulário na tela
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
   
   Call sConectaLocal
   
   gSql = "select nome,nivel"
   gSql = gSql & " FROM tb_usuarios"
   gRs.Open gSql, cnnLocal, adOpenKeyset
   If gRs.BOF And gRs.EOF Then
      gnNivel = 1
      gcUsuario = "Master"
   Else
      gnNivel = gRs!nivel
      gcUsuario = gRs!nome
   End If

   If gRs.State = adStateOpen Then
      gRs.Close
  End If
  Set gRs = Nothing
  
  If gnNivel = 1 Then
    Me.cmdParametros.Enabled = True
    Me.cmdParametros.Visible = True
    Me.cmdUsuarios.Enabled = True
    Me.cmdUsuarios.Visible = True
    Me.cmdStatus.Enabled = True
    Me.cmdStatus.Visible = True
  End If

  'sHora = "07:00"   ' Estabelecemos um horario inicial que depopis pode ser parametrizado

   'aLista_Horario(0) = sHora
  ' For i = 0 To 25   ' Vai até as 20:00 - Podemos ver parametrização depois
  '   sHora = DateAdd("n", 30, CDate(sHora))
  '   aLista_Horario(i) = Mid(sHora, 1, 5)
  '   CmbHorario.AddItem (aLista_Horario(i))
  ' Next
  Call MontaAtendimentos
'  If gnNivel = 1 Then
'     Me.cmdParametros.Visible = True
'     Me.cmdParametros.Enabled = True
'     Me.cmdUsuarios.Visible = True
'     Me.cmdUsuarios.Enabled = True
'  Else
'     Me.cmdParametros.Visible = False
'     Me.cmdParametros.Enabled = False
'     Me.cmdUsuarios.Visible = False
'     Me.cmdUsuarios.Enabled = False
'  End If
  'DTPicker1.Value = Date
End Sub

Private Sub MSFlexPedidos_Click()
    If Me.MSFlexPedidos.col = 3 Then
        Me.frmSituacao.Visible = True
        Me.frmSituacao.Enabled = True
        Call suCarregarSituacao
    Else
        Call suEditarPedido
    End If
End Sub

Private Sub MSFlexPedidos_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
 Dim row As Long
    Dim col As Long
    
    ' Get the current row and column under the mouse
    row = MSFlexPedidos.MouseRow
    col = MSFlexPedidos.MouseCol
    
    ' Check if it's a valid cell (not fixed headers, optional)
    If row >= MSFlexPedidos.FixedRows And col >= MSFlexPedidos.FixedCols Then
        ' Set the tooltip text based on cell content
        If col = 3 Then
           MSFlexPedidos.ToolTipText = "Clique aqui se quiser alterar a situação do pedido"
        Else
            MSFlexPedidos.ToolTipText = "Row: " & row & ", Col: " & col & _
                                  " | Data: " & MSFlexPedidos.TextMatrix(row, col)
        End If
    Else
        ' No tooltip for headers
        MSFlexPedidos.ToolTipText = ""
    End If
End Sub

Private Sub txtNovaData_GotFocus()
   Call SelText(txtNovaData)
End Sub

Private Sub txtNovaData_KeyPress(KeyAscii As Integer)
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
    Select Case Len(txtNovaData.Text)
        Case 2, 5
            If KeyAscii <> 8 Then ' Se não for Backspace
                txtNovaData.Text = txtNovaData.Text & "/"
                txtNovaData.SelStart = Len(txtNovaData.Text)
            End If
    End Select
End Sub

Private Sub txtNovaData_Validate(Cancel As Boolean)
  If Not f_ValidaData(txtNovaData.Text) Then
      MsgBox "Data Inválida.", vbInformation, "Atenção"
      txtNovaData.SetFocus
      Cancel = True
   End If
End Sub

Private Sub suCarregarSituacao()

    oldCol = Me.MSFlexPedidos.col
    MSFlexPedidos.col = 3
    oldsituacao = IIf(MSFlexPedidos.Text = "A iniciar", 1, IIf(MSFlexPedidos.Text = "Iniciado", 2, IIf(MSFlexPedidos.Text = "Finalizado", 3, 4)))

    MSFlexPedidos.col = oldCol

    Call sConectaLocal
  
    strSql = ""
    strSql = strSql & "select id_situacao,descricao from tb_situacao WHERE id_situacao >= " & oldsituacao
    Set Rstemp = New ADODB.Recordset
    Rstemp.Open strSql, cnnLocal, 1, 2
    Rstemp.MoveFirst
    Me.cmbSituacao.Clear
    
    Indcmb = 0
    Do While Not Rstemp.EOF
      
        Me.cmbSituacao.AddItem (Rstemp!descricao)
        Me.cmbSituacao.ItemData(Me.cmbSituacao.NewIndex) = Rstemp!id_situacao
        If Trim(Me.MSFlexPedidos.Text) = Trim(Rstemp!descricao) Then
            Me.cmbSituacao.ListIndex = Indcmb
        End If
        Rstemp.MoveNext
        Indcmb = Indcmb + 1
    Loop
    Rstemp.Close
    Set Rstemp = Nothing
    'Me.cmbSituacao.ListIndex = 0

End Sub

Private Sub suConsisteSituacao()

oldCol = Me.MSFlexPedidos.col
MSFlexPedidos.col = 3

' IIf invertido porque o usuario pode cadastrar "A Iniciar" ou "A iniciar" (i minúsculo) no status
oldsituacao = IIf(MSFlexPedidos.Text = "Iniciado", 2, _
IIf(MSFlexPedidos.Text = "Finalizado", 3, _
IIf(MSFlexPedidos.Text = "Entregue", 4, 1)))
' partes(0) será "A iniciar"
    
    'Consistir a data entrada com a data de compra.
    
    situacao = cmbSituacao.ItemData(cmbSituacao.ListIndex)
    With MSFlexPedidos
        '.Col = 5:  situacao = .Text
        .col = 4:  lDataCompra = Format(.Text, "yyyy-mm-dd")
        .col = 5:  lDataInicio = Format(.Text, "yyyy-mm-dd")
        .col = 7:  lDataFinaliza = Format(.Text, "yyyy-mm-dd")
        .col = 8: lDataEntrega = Format(.Text, "yyyy-mm-dd")
    End With
    
    If situacao <= oldsituacao Then
        If MsgBox("Situação escolhida já definida. " & vbCrLf & "Deseja alterá-la? ", vbYesNo, "Atenção") = vbNo Then
            Me.cmbSituacao.SetFocus
            Exit Sub
        End If
    End If
    
    Select Case situacao
       Case 1  ' "A iniciar"
            'Data Compra
            If lDataCompra > Format(Me.txtNovaData.Text, "yyyy-mm-dd") Then
                If MsgBox("Data informada menor que a data da compra." & vbCrLf & "Deseja alterá-la? ", vbYesNo, "Atenção") = vbNo Then
                    Me.txtNovaData.SetFocus
                    Exit Sub
                Else
                    lDataCompra = Format(txtNovaData.Text, "yyyy-mm-dd")
                End If
            End If
       
       Case 2   '"Iniciado"
            If Len(lDataInicio) > 0 Then
                If MsgBox("Data de inicio de produção já definida. " & vbCrLf & "Deseja alterá-la? ", vbYesNo, "Atenção") = vbNo Then
                    Me.txtNovaData.SetFocus
                    Exit Sub
                Else
                   lDataInicio = Format(txtNovaData.Text, "yyyy-mm-dd")
               End If
            Else
               lDataInicio = Format(txtNovaData.Text, "yyyy-mm-dd")
            End If
            'Data Inicio de produção
            If lDataInicio > Format(Me.txtNovaData.Text, "yyyy-mm-dd") Then
                If MsgBox("Data menor que a data de inicio de produção já informada. " & vbCrLf & "Deseja alterá-la", vbYesNo, "Atenção") = vbNo Then
                    Me.txtNovaData.SetFocus
                    Exit Sub
                Else
                    lDataInicio = Format(txtNovaData.Text, "yyyy-mm-dd")
                End If
            End If
       
       Case 3    ' "Finalizado"
            If Len(lDataFinaliza) > 0 Then
                If MsgBox("Data de finalização já definida. " & vbCrLf & "Deseja alterá-la? ", vbYesNo, "Atenção") = vbNo Then
                    Me.txtNovaData.SetFocus
                    Exit Sub
                Else
                   lDataFinaliza = Format(txtNovaData.Text, "yyyy-mm-dd")
                End If
            Else
                lDataFinaliza = Format(txtNovaData.Text, "yyyy-mm-dd")
            End If
       
            If lDataFinaliza > Format(Me.txtNovaData.Text, "yyyy-mm-dd") Then
                If MsgBox("Data informada menor que a data de Finalização." & vbCrLf & "Deseja alterá-la?", vbYesNo, "Atenção") = vbNo Then
                    Me.txtNovaData.SetFocus
                    Exit Sub
                Else
                   lDataFinaliza = Format(txtNovaData.Text, "yyyy-mm-dd")
                End If
            End If
       
       Case 4    ' "Entregue"
'            Me.MSFlexPedidos.Col = 10  'Data de entrega
            If Len(lDataEntrega) > 0 Then
                If MsgBox("Data de entrega já definida. " & vbCrLf & "Deseja alterá-la? ", vbYesNo, "Atenção") = vbNo Then
                    Me.txtNovaData.SetFocus
                    Exit Sub
                Else
                   lDataEntrega = Format(txtNovaData.Text, "yyyy-mm-dd")
               End If
            Else
               lDataEntrega = Format(txtNovaData.Text, "yyyy-mm-dd")
            End If
            
            If lDataEntrega > Format(Me.txtNovaData.Text, "yyyy-mm-dd") Then
                If MsgBox("Data informada menor que a data de entrega." & vbCrLf & "Deseja alterá-la?", vbYesNo, "Atenção") = vbNo Then
                    Me.txtNovaData.SetFocus
                    Exit Sub
                Else
                    lDataEntrega = Format(txtNovaData.Text, "yyyy-mm-dd")
                End If
            End If
        
    End Select
    
    'Grava a nova situação se alterada e também a data da mesma
    MSFlexPedidos.col = 9
    
    strSql = ""
    strSql = strSql & "UPDATE tb_pedidos set situacao =  " & cmbSituacao.ItemData(cmbSituacao.ListIndex) & ","
    strSql = strSql & "dataCompra = '" & lDataCompra & "'"
    If Len(lDataInicio) > 0 Then
        strSql = strSql & ",dataInicioProd =  '" & lDataInicio & "'"
    End If
    If Len(lDataFinaliza) > 0 Then
        strSql = strSql & ",dataFinaliza =   '" & lDataFinaliza & "'"
    End If
    If Len(lDataEntrega) > 0 Then
        strSql = strSql & ",dataEntrega =  '" & lDataEntrega & "' "
    End If
    strSql = strSql & " WHERE id_venda = " & Me.MSFlexPedidos.Text
    
    cnnLocal.Execute (strSql)
    
    Me.frmSituacao.Visible = False
    Me.frmSituacao.Enabled = False

    Call MontaAtendimentos

    Me.MSFlexPedidos.col = oldCol
   
   'Print Me.cmbSituacao.ItemData(Me.cmbSituacao.ListIndex)
   'Me.txtNovaData.Text
End Sub

Private Sub suEditarPedido()
    
    Me.MSFlexPedidos.col = 9
    
    If MsgBox("Deseja mesmo alterar o pedido num." & Me.MSFlexPedidos.Text & " ? ", vbYesNo, "Atenção") = vbYes Then
        gNumPedido = Val(Me.MSFlexPedidos.Text)
        frmNovoPedido.Show vbModal
    End If
    
End Sub

