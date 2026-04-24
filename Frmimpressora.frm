VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Frmimpressora 
   BackColor       =   &H00808000&
   BorderStyle     =   0  'None
   Caption         =   "Cadastro de Impressoras"
   ClientHeight    =   7020
   ClientLeft      =   7545
   ClientTop       =   3195
   ClientWidth     =   6990
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   Tag             =   "cadvend"
   Begin VB.TextBox txtRolos 
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Top             =   1665
      Width           =   885
   End
   Begin VB.TextBox TxtModelo 
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Top             =   1245
      Width           =   5175
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   1380
      TabIndex        =   4
      Top             =   6015
      Width           =   4245
      Begin VB.CommandButton cmddesfaz 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Desfaz"
         Enabled         =   0   'False
         Height          =   540
         Left            =   2865
         Picture         =   "Frmimpressora.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         Tag             =   "&Update"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton CmdSair 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Sair"
         Height          =   540
         Left            =   3555
         Picture         =   "Frmimpressora.frx":00FA
         Style           =   1  'Graphical
         TabIndex        =   9
         Tag             =   "&Update"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Incluir"
         Height          =   540
         Left            =   105
         Picture         =   "Frmimpressora.frx":01F4
         Style           =   1  'Graphical
         TabIndex        =   8
         Tag             =   "&Add"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Excluir"
         Height          =   540
         Left            =   1485
         Picture         =   "Frmimpressora.frx":02DE
         Style           =   1  'Graphical
         TabIndex        =   7
         Tag             =   "&Delete"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton cmdEditar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Editar"
         Height          =   540
         Left            =   795
         Picture         =   "Frmimpressora.frx":0450
         Style           =   1  'Graphical
         TabIndex        =   6
         Tag             =   "&Refresh"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Salvar"
         Enabled         =   0   'False
         Height          =   540
         Left            =   2190
         Picture         =   "Frmimpressora.frx":05C2
         Style           =   1  'Graphical
         TabIndex        =   5
         Tag             =   "&Update"
         Top             =   150
         Width           =   615
      End
   End
   Begin VB.TextBox TxtMarca 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   795
      Width           =   5190
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexFila 
      Height          =   3735
      Left            =   480
      TabIndex        =   11
      Top             =   2205
      Width           =   6045
      _ExtentX        =   10663
      _ExtentY        =   6588
      _Version        =   393216
      Rows            =   5
      Cols            =   4
      FixedCols       =   0
      ScrollBars      =   2
      Appearance      =   0
      FormatString    =   "Id |   Marca                                       | Modelo                                          |   Rolos "
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Cadastro de Impressoras"
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
      Left            =   1425
      TabIndex        =   15
      Top             =   135
      Width           =   4395
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rolos:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   690
      TabIndex        =   14
      Top             =   1725
      Width           =   450
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Marca:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   645
      TabIndex        =   13
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Modelo:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   570
      TabIndex        =   12
      Top             =   1275
      Width           =   570
   End
   Begin VB.Label Lbltipovend 
      BackStyle       =   0  'Transparent
      Caption         =   "id"
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   225
      TabIndex        =   0
      Top             =   825
      Width           =   465
   End
End
Attribute VB_Name = "Frmimpressora"
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
   
   gSql = "select * FROM tb_impressoras"
   gRs.Open gSql, cnnLocal, adOpenKeyset
   
End Sub

Private Sub Carrega_Grid()
'Teste do MSFlexFila
  MSFlexFila.Row = 0
  MSFlexFila.ColAlignment(-1) = flexAlignLeftCenter
  
  With gRs
      '.MoveLast
      'nItem = .RecordCount
      '.MoveFirst
      MSFlexFila.Rows = 1
      MSFlexFila.Redraw = False
      
      Do While Not .EOF
         MSFlexFila.Rows = MSFlexFila.Rows + 1
         MSFlexFila.Row = MSFlexFila.Rows - 1
         MSFlexFila.Col = 0: MSFlexFila.Text = f_nulo(!id, "")
         MSFlexFila.Col = 1: MSFlexFila.Text = f_nulo(!Marca, "")
         MSFlexFila.Col = 2: MSFlexFila.Text = f_nulo(!modelo, "")
         MSFlexFila.Col = 3: MSFlexFila.Text = f_nulo(!qtrolos, "") & "  "
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
   Me.TxtModelo.Text = gRs("modelo")
   Me.txtRolos = gRs!qtrolos
   
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
    If MsgBox("Deseja realmente apagar essa impressora ? ", vbYesNo, "Atenção") = vbYes Then
        gSql = "delete from tb_impressoras where id = " & Val(Me.Lbltipovend.Caption)
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
   'gRs.Close
   If lIncluir Then
      gSql = "INSERT INTO tb_impressoras (marca,modelo,qtrolos,ocupada,operador,datatual) "
      gSql = gSql & "VALUES ('" & Me.TxtMarca.Text & "','"
      gSql = gSql & Me.TxtModelo.Text & "'," & Me.txtRolos.Text & ",'N',"
      gSql = gSql & f_nulo(gncodoperador, 99) & ",'" & Format(Date, "yyyy-mm-dd") & "')"
      cnnLocal.Execute gSql
      lIncluir = False
      
   Else
      gSql = "UPDATE tb_impressoras SET marca = '" & Me.TxtMarca.Text & "',"
      gSql = gSql & "modelo = '" & Me.TxtModelo.Text & "',"
      gSql = gSql & "qtrolos = '" & Val(Me.txtRolos.Text) & "',ocupada = 'N',"
      gSql = gSql & " operador = " & f_nulo(gncodoperador, 99) & ", datatual = '" & Format(Date, "yyyy-mm-dd") & "'"
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
      gSql = "INSERT INTO tb_impressoras (marca,modelo,qtrolos,ocupada,operador,datatual) "
      gSql = gSql & "VALUES ('" & Me.TxtMarca.Text & "','"
      gSql = gSql & Me.TxtModelo.Text & "'," & Me.txtRolos.Text & ",'NÃO',"
      gSql = gSql & f_nulo(gncodoperador, 99) & ",'" & Format(Date, "yyyy-mm-dd") & "')"
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
  
  oldrow = MSFlexFila.Row
  
  MSFlexFila.Row = 0
    
    With MSFlexFila
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
    
    .Col = 0:   Lbltipovend.Caption = .Text: .CellBackColor = vbYellow
    .Col = 1:   TxtMarca.Text = .Text: .CellBackColor = vbYellow
    .Col = 2:   TxtModelo.Text = .Text: .CellBackColor = vbYellow
    .Col = 3:   txtRolos.Text = .Text: .CellBackColor = vbYellow
    .TopRow = .Row
   
End With
   
   Desabilita Me
   Me.cmdUpdate.Enabled = False
   Me.cmdDesfaz.Enabled = False
   Me.cmdEditar.Enabled = True
   Me.cmdAdd.Enabled = True
   Me.cmdSair.Enabled = True
   Me.cmdDelete.Enabled = True

End Sub


Private Sub TxtMarca_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub TxtModelo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub txtRolos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub
