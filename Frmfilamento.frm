VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form Frmfilamento 
   BackColor       =   &H00808000&
   BorderStyle     =   0  'None
   Caption         =   "Cadastro de Filamentos"
   ClientHeight    =   7065
   ClientLeft      =   1605
   ClientTop       =   1635
   ClientWidth     =   9915
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   9915
   ShowInTaskbar   =   0   'False
   Tag             =   "cadvend"
   Begin VB.TextBox txtPeso 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6240
      MaxLength       =   7
      TabIndex        =   21
      Top             =   1320
      Width           =   720
   End
   Begin VB.TextBox txtEstoque 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   8520
      MaxLength       =   7
      TabIndex        =   5
      Top             =   1320
      Width           =   900
   End
   Begin VB.TextBox txtPreco 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4230
      MaxLength       =   20
      TabIndex        =   4
      Top             =   1320
      Width           =   1050
   End
   Begin VB.TextBox txtCor 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1110
      MaxLength       =   50
      TabIndex        =   3
      Top             =   1320
      Width           =   2265
   End
   Begin VB.TextBox TxtTipo 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4980
      MaxLength       =   100
      TabIndex        =   2
      Top             =   720
      Width           =   3705
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   2490
      TabIndex        =   6
      Top             =   6030
      Width           =   4095
      Begin VB.CommandButton cmddesfaz 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Desfaz"
         Enabled         =   0   'False
         Height          =   540
         Left            =   2820
         Picture         =   "Frmfilamento.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         Tag             =   "&Update"
         Top             =   120
         Width           =   615
      End
      Begin VB.CommandButton CmdSair 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Sair"
         Height          =   540
         Left            =   3495
         Picture         =   "Frmfilamento.frx":00FA
         Style           =   1  'Graphical
         TabIndex        =   11
         Tag             =   "&Update"
         Top             =   120
         Width           =   615
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Incluir"
         Height          =   540
         Left            =   150
         Picture         =   "Frmfilamento.frx":01F4
         Style           =   1  'Graphical
         TabIndex        =   10
         Tag             =   "&Add"
         Top             =   120
         Width           =   615
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Excluir"
         Height          =   540
         Left            =   1485
         Picture         =   "Frmfilamento.frx":02DE
         Style           =   1  'Graphical
         TabIndex        =   9
         Tag             =   "&Delete"
         Top             =   120
         Width           =   615
      End
      Begin VB.CommandButton cmdEditar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Editar"
         Height          =   540
         Left            =   825
         Picture         =   "Frmfilamento.frx":0450
         Style           =   1  'Graphical
         TabIndex        =   8
         Tag             =   "&Refresh"
         Top             =   120
         Width           =   615
      End
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Salvar"
         Enabled         =   0   'False
         Height          =   540
         Left            =   2160
         Picture         =   "Frmfilamento.frx":05C2
         Style           =   1  'Graphical
         TabIndex        =   7
         Tag             =   "&Update"
         Top             =   120
         Width           =   615
      End
   End
   Begin VB.TextBox TxtMarca 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1110
      MaxLength       =   100
      TabIndex        =   1
      Top             =   720
      Width           =   3150
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexFila 
      Height          =   3945
      Left            =   315
      TabIndex        =   13
      Top             =   1830
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   6959
      _Version        =   393216
      Rows            =   5
      Cols            =   7
      FixedCols       =   0
      ScrollBars      =   2
      FormatString    =   $"Frmfilamento.frx":06BC
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Peso:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   5670
      TabIndex        =   20
      Top             =   1350
      Width           =   495
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Cadastro de Filamentos"
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
      Left            =   2385
      TabIndex        =   19
      Top             =   75
      Width           =   4395
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Estoque:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   7695
      TabIndex        =   18
      Top             =   1350
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Preço:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   3555
      TabIndex        =   17
      Top             =   1350
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cor:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   660
      TabIndex        =   16
      Top             =   1320
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Marca:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   480
      TabIndex        =   15
      Top             =   765
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   4470
      TabIndex        =   14
      Top             =   780
      Width           =   450
   End
   Begin VB.Label Lbltipovend 
      BackStyle       =   0  'Transparent
      Caption         =   "id"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   105
      TabIndex        =   0
      Top             =   765
      Width           =   465
   End
End
Attribute VB_Name = "Frmfilamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private lIncluir As Boolean
Private lPrimeiro As Boolean

Private Sub Abre_Le_rst()
   If gRs.State = adStateOpen Then
      gRs.Close
   End If
   
   gSql = "select * FROM tb_filamentos"
   gRs.Open gSql, cnnLocal, adOpenKeyset
   
End Sub

Private Sub Carrega_Grid()
  
  MSFlexFila.row = 0
  
  With gRs
      '.MoveLast
      'nItem = .RecordCount
      '.MoveFirst
      MSFlexFila.Redraw = False
      MSFlexFila.Rows = 1
      Do While Not .EOF
         MSFlexFila.Rows = MSFlexFila.Rows + 1
         MSFlexFila.row = MSFlexFila.Rows - 1
         MSFlexFila.col = 0: MSFlexFila.Text = f_nulo(!id, "")
         MSFlexFila.col = 1: MSFlexFila.Text = f_nulo(!Marca, "")
         MSFlexFila.col = 2: MSFlexFila.Text = f_nulo(!tipo, "")
         MSFlexFila.col = 3: MSFlexFila.Text = f_nulo(!cor, "")
         MSFlexFila.col = 4: MSFlexFila.Text = f_nulo(Format(!VALOR, "0.00"), "")
         MSFlexFila.col = 5: MSFlexFila.Text = f_nulo(Format(!peso, "0.00"), "")
         MSFlexFila.col = 6: MSFlexFila.Text = f_nulo(Format(!qtde_Estoque, "0"), "")
         .MoveNext
         
       Loop
       MSFlexFila.FixedRows = 1
       MSFlexFila.Redraw = True
  End With
  
  End Sub

Private Sub Carrega_tela()
   'Limpa as variaveis da tela se caso ficarem com dados da outra tela
   limpa_tela Me
   'Carrega a tela com os dados do registro
   Me.Lbltipovend.Caption = gRs("id")
   Me.TxtMarca.Text = gRs("marca")
   Me.TxtTipo.Text = gRs("tipo")
   Me.txtCor = gRs!cor
   Me.txtPreco = gRs!VALOR
   Me.txtEstoque = gRs!qtde_Estoque
   
End Sub
Private Sub cmdAdd_Click()
   
   lIncluir = True
   limpa_tela Me
   
   Me.Lbltipovend.Caption = ""
   Me.TxtMarca.SetFocus
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
    If MsgBox("Deseja realmente apagar este Filamento ? ", vbYesNo, "Atenção") = vbYes Then
        gSql = "delete from tb_filamentos where id = " & Val(Me.Lbltipovend.Caption)
        cnnLocal.Execute gSql
        Abre_Le_rst
        Carrega_Grid
        gRs.MoveFirst
        Carrega_tela
        Desabilita Me
     End If
     Exit Sub
     
ErroDelete:
     MsgBox "Deu erro na exclusao do Filamento " & Chr(13) & "Instrucao Sql = '" & _
            cSql & "'  "
End Sub


Private Sub cmddesfaz_Click()
  
  lIncluir = False
  ' Carrega_tela
  Desabilita Me
  MSFlexFila_Click
   
  Me.cmdUpdate.Enabled = False
  Me.cmdDesfaz.Enabled = False
  Me.cmdEditar.Enabled = True
  Me.cmdAdd.Enabled = True
  Me.cmdSair.Enabled = True
  Me.cmdDelete.Enabled = True
  
End Sub

Private Sub cmdEditar_Click()
   
   Habilita Me
        
   Me.cmdUpdate.Enabled = True
   Me.cmdDesfaz.Enabled = True
   Me.cmdEditar.Enabled = False
   Me.cmdAdd.Enabled = False
   Me.cmdSair.Enabled = False
   Me.cmdDelete.Enabled = False
   Me.TxtMarca.SetFocus
   
End Sub

Private Sub CmdSair_Click()
   Unload Me
End Sub

Private Sub cmdUpdate_Click()
   If gRs.State = adStateOpen Then
      gRs.Close
   End If
   If lIncluir Then
      gSql = "INSERT INTO tb_filamentos (marca,tipo,cor,valor,qtde_estoque,operador,datatual) "
      gSql = gSql & "VALUES ('" & Me.TxtMarca.Text & "','"
      gSql = gSql & Me.TxtTipo.Text & "','" & Me.txtCor.Text & "',"
      gSql = gSql & SoNumero(Me.txtPreco.Text) & ","
      gSql = gSql & Me.txtEstoque.Text & ","
      gSql = gSql & f_nulo(gnCodOperador, 99) & ",'" & Format(Date, "yyyy-mm-dd") & "')"
      cnnLocal.Execute gSql
      lIncluir = False
      
   Else
      gSql = "UPDATE tb_filamentos SET marca = '" & Me.TxtMarca.Text & "',"
      gSql = gSql & "tipo = '" & Me.TxtTipo.Text & "',"
      gSql = gSql & "cor = '" & Me.txtCor.Text & "',"
      gSql = gSql & "valor = " & SoNumero(Me.txtPreco.Text) & ","
      gSql = gSql & "qtde_estoque = " & Val(Me.txtEstoque.Text) & ","
      gSql = gSql & " operador = " & f_nulo(gnCodOperador, 99) & ", datatual = '" & Format(Date, "yyyy-mm-dd") & "'"
      gSql = gSql & " WHERE id = " & Val(Lbltipovend.Caption)
      cnnLocal.Execute gSql
            
   End If
      
   Abre_Le_rst
   Carrega_Grid
   gRs.MoveFirst
   Carrega_tela
   
   Desabilita Me
   
   Me.cmdUpdate.Enabled = False
   Me.cmdDesfaz.Enabled = False
   Me.cmdEditar.Enabled = True
   Me.cmdAdd.Enabled = True
   Me.cmdSair.Enabled = True
   Me.cmdDelete.Enabled = True
     
 End Sub




Private Sub Form_Activate()
  
  Abre_Le_rst
  limpa_tela Me
   
  Me.Lbltipovend.Caption = ""
  If gRs.BOF And gRs.EOF Then
      If MsgBox("Arquivo vazio. Incluir dados agora ?", vbYesNo, "Atenção ") = vbYes Then
      gSql = "INSERT INTO tb_filamentos (marca,tipo,cor,valor,qtde_estoque,operador,datatual) "
      gSql = gSql & "VALUES ('" & Me.TxtMarca.Text & "','"
      gSql = gSql & Me.TxtTipo.Text & "','" & Me.txtCor.Text & "',"
      gSql = gSql & f_nulo(Me.txtPreco.Text, 0) & ","
      gSql = gSql & f_nulo(Me.txtEstoque.Text, 0) & ","
      gSql = gSql & "'" & f_nulo(gnCodOperador, 99) & "','" & Format(Date, "yyyy-mm-dd") & "')"
      cnnLocal.Execute gSql
         
         Abre_Le_rst
         Me.Lbltipovend.Caption = gRs!id
         cmdEditar_Click
         lPrimeiro = True
      Else
         Desabilita Me
      End If
      
   Else
      gRs.MoveFirst
      Carrega_tela
    
      Desabilita Me
      lIncluir = False
      lPrimeiro = False
   End If
   Carrega_Grid
   
   lIncluir = False

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then Sendkeys "{TAB}"
     
End Sub

Private Sub Form_Load()
   
   'Centraliza a tela no video
   Me.Move (Screen.Width - Me.Width) / 2, _
           (Screen.Height - Me.Height) / 2
   
  End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
    If gRs.State = adStateOpen Then
      gRs.Close
   End If
    
End Sub


Private Sub MSFlexFila_Click()
Dim oldrow As Long
  
  oldrow = MSFlexFila.row
  
  MSFlexFila.row = 0
  
  With MSFlexFila
    .Redraw = False
    Do While True
       .row = .row + 1
       For ix = 0 To .Cols - 1
           .col = ix: .CellBackColor = vbWhite
       Next
       If .row = .Rows - 1 Then
          Exit Do
       End If
    Loop
    .Redraw = True
    
    .row = oldrow
    
    .col = 0:   Lbltipovend.Caption = .Text: .CellBackColor = vbYellow
    .col = 1:   TxtMarca.Text = .Text: .CellBackColor = vbYellow
    .col = 2:   TxtTipo.Text = .Text: .CellBackColor = vbYellow
    .col = 3:   txtCor.Text = .Text: .CellBackColor = vbYellow
    .col = 4:   txtPreco.Text = .Text: .CellBackColor = vbYellow
    .col = 5:   txtEstoque.Text = .Text: .CellBackColor = vbYellow
    .TopRow = .row
    
   Me.cmdUpdate.Enabled = False
   Me.cmdDesfaz.Enabled = False
   Me.cmdEditar.Enabled = True
   Me.cmdAdd.Enabled = True
   Me.cmdSair.Enabled = True
   Me.cmdDelete.Enabled = True

   
End With

End Sub

Private Sub TxtEntrada_Validate(Cancel As Boolean)
    TxtEntrada.Text = UCase(TxtEntrada.Text)
    If UCase(TxtEntrada.Text) <> "S" And UCase(TxtEntrada.Text) <> "N" Then
       MsgBox "Digite somente 'S' ou 'N' por favor", vbOKOnly, "Atenção: " + gOperador
       Cancel = True
    End If
    
End Sub

Private Sub txtCor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub txtEstoque_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub TxtMarca_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub txtPreco_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub txtPreco_LostFocus()
    ' Verifica se o campo não está vazio
    If txtPreco.Text <> "" Then
        ' Converte o texto para valor numérico e formata como moeda
        ' A função FormatCurrency usa as configurações regionais do sistema
        txtPreco.Text = FormatCurrency(CDbl(txtPreco.Text))
    End If
End Sub

Private Sub TxtTipo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub
