Attribute VB_Name = "Mod_geral"
' Exemplo simplificado de declaração - Para SqLite
'****************************************************
Public cnnLocal As New ADODB.Connection
Public cmd As New ADODB.Command

'
'*************************
'variaveis para recordsets
Public Rstemp       As New ADODB.Recordset
Public RsTemp1      As New ADODB.Recordset
Public Rstemp2      As New ADODB.Recordset
Public gRs          As New ADODB.Recordset
'
'variaveis pra controle de registro
Global Situacao_Registro As String
Global Dias_Uso_Sistema As Integer
Global ConsultaProd_Ped As Integer
Global flagConsultaPedProd As Boolean
Global gNumPedido As Integer

Public gTransacao As Boolean
'
Public gMensagem As String
Public strSql  As String
Public strSql1 As String
Public strSql2 As String
Public strSql3 As String
Public gSql  As String
Public tmpSQL As String
Public strConn As String

Public strPesqProdProv As Boolean
Public strFormaPgto As String

Global sysNomeAcesso As String
Global gOperador As Integer
'
'*************************************************************************************
'*** Fabio Reinert ( Alemao) 06/2017 - Inclusão de captura de IP do cliente - Inicio *
'*************************************************************************************
'
Public STR_IP_COMPUTADOR As String

Public Function BuscaIP() As String
Dim NIC As Variant
Dim NICs As Object

sysNomeAcesso = "MASTER"

On Error GoTo errError

Set NICs = GetObject("winmgmts:").InstancesOf("Win32_NetworkAdapterConfiguration")

For Each NIC In NICs
   If NIC.IPEnabled Then
        BuscaIP = NIC.IpAddress(0)
    End If
Next NIC

'ou
'Dim IPConfig As Variant
'Dim IPConfigSet As Object
'Set IPConfigSet = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery("SELECT IPAddress FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = TRUE")
'
'For Each IPConfig In IPConfigSet
' If Not IsNull(IPConfig.IPAddress) Then MsgBox IPConfig.IPAddress(0), vbInformation
'Next IPConfig

Exit Function
    
errError:
    
    If Err.Number <> 0 Then
        Err.Clear
    End If
    BuscaIP = ""

End Function
'
'*************************************************************************************
'*** Fabio Reinert ( Alemao) 06/2017 - Inclusão de captura de IP do cliente - Fim    *
'*************************************************************************************
'
Public Sub sConectaLocal()
         
  On Error GoTo Erro_sConectaLocal

  '*** Conexão com SQL Server
  'cnnLocal.ConnectionString = "Sql_Server_Express"
  'cnnLocal.Open
  
  'ou
  'connstring = "Provider=SQL Server Native Client 11.0;Server=DESKTOP-TN96NN4\\SQLEXPRESS;Database=db_printer3d;trusted_connection=yes"
  
  'Conexão com Mysql que vou deixar comentada para testar a conexão com Sqlite ---> Qualquer coisa eu volto atrás
  Set cnnLocal = New ADODB.Connection
  cnnLocal.Open "Driver={MySQL ODBC 8.0 Unicode Driver};Server=localhost;Database=db_Printer3d;User=root;Password=oyster;"
 
  
Exit Sub

Erro_sConectaLocal:
    Call sMostraErro("sConectaLocal", Err.Number, Err.Description)
    'Call Fecha_Formularios
    End

End Sub
'
''tutup database
'Public Function closeDB()
'
'  '  sqlite3_close (DBz)
'
'End Function

Public Sub sMostraAviso(Optional ByVal pTitulo As String, Optional ByVal pTexto1 As String, _
                        Optional ByVal pTexto2 As String, _
                        Optional ByVal pTexto3 As String, _
                        Optional ByVal pTexto4 As String)
                        
    Dim fAviso As Form
    If IsMissing(pTexto2) Then
        pTexto2 = ""
    End If
    If IsMissing(pTexto3) Then
        pTexto3 = ""
    End If
    If IsMissing(pTexto4) Then
        pTexto4 = ""
    End If
    If IsMissing(pTitulo) Then
        pTitulo = "Aviso:"
    End If
    Set fAviso = New frmAviso
    fAviso.lblAviso1.Caption = pTexto1
    fAviso.lblAviso2.Caption = pTexto2
    fAviso.lblAviso3.Caption = pTexto3
    fAviso.lblAviso4.Caption = pTexto4
    fAviso.Caption = pTitulo
    fAviso.Show vbModal
    Unload fAviso
    Set fAviso = Nothing
End Sub

Public Sub sMostraErro(Optional ByVal pModulo, Optional ByVal pErroNumero, Optional ByVal pErroDesc)
        
    If pModulo = "" Then
        pModulo = "Geral"
    End If
    If pErroNumero = "" Then
       pErroNumero = Err.Number
    End If
    If pErroDesc = "" Then
       pErroDesc = Err.Description
    End If
    Call sMostraAviso("Atenção - Erro: ", "Contate a Info Sistemas informando o erro abaixo:", _
                      "No.erro: " & pErroNumero & " Descr.: " & pErroDesc, _
                      "Módulo do erro: " & pModulo, "Sistema será encerrado")
    'Call Fecha_Formularios
    End
End Sub

Sub SelText(object As Control)
    
    With object
        .SelStart = 0
        .SelLength = Len(object)
    End With

End Sub

Public Sub Sendkeys(Text As Variant, Optional wait As Boolean = False)
   Dim WshShell As Object
   Set WshShell = CreateObject("wscript.shell")
   WshShell.Sendkeys CStr(Text), wait
   Set WshShell = Nothing
End Sub
