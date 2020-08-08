Sub btExecuta_Click()

Application.ScreenUpdating = False

Dim vErro           As String
Dim IElocation      As String
Dim Resultado(1 To 15) As String

Dim vNome           As String
Dim vDados          As String
Dim vSituacao       As String

Dim W               As Worksheet

Dim Ie              As Object

Dim UltCel          As Range

Dim A               As Integer
Dim col             As Integer
Dim vSegundos       As Integer

Dim ln              As Long

Set W = Planilha1

vSegundos = 3

W.Range("A2").Select
W.Range("B2:d1000").Clear

W.Range("A1").Value = "num_cpf"
W.Range("b1").Value = "nome_pessoa_física"
W.Range("c1").Value = "situação"
W.Range("d1").Value = "informações complementares"

Set Ie = CreateObject("InternetExplorer.Application")
Set UltCel = W.Cells(W.Rows.Count, 1).End(xlUp)

With Ie
    .navigate "https://www.situacaocadastral.com.br/"
    .Visible = False
End With

Do While Ie.busy
Loop

ln = 2
col = 1

Application.Wait TimeSerial(Hour(Now()), Minute(Now()), Second(Now()) + vSegundos)

Do While ln <= UltCel.Row

    Ie.Document.getelementbyid("doc").Value = W.Cells(ln, col)
    Ie.Document.getelementbyid("consultar").Click
    
    Application.Wait TimeSerial(Hour(Now()), Minute(Now()), Second(Now()) + vSegundos)
    
    On Error Resume Next
        vErro = Ie.Document.getelementbyid("mensagem").innertext
        
    On Error GoTo 0
    
    If vErro = "Informe um termo válido! " Then
        Ie.Document.getelementbyid("consultar").Click
        Application.Wait TimeSerial(Hour(Now()), Minute(Now()), Second(Now()) + vSegundos)
    ElseIf vErro = "Informe um termo válido! " Then
        W.Cells(ln, col + 1).Value = "'" & vErro
    Else
        vErro = vbNullString
    End If
    
    Do While Ie.busy
    Loop
    
    If vErro = vbNullString Then
    
        vNome = Ie.Document.getelementsbyclassname("dados nome")(0).innertext
        vDados = Ie.Document.getelementsbyclassname("dados texto")(0).innertext
        vSituacao = Ie.Document.getelementsbyclassname("dados situacao")(0).innertext
        
        W.Cells(ln, col + 1) = vNome
        W.Cells(ln, col + 2) = vSituacao
        W.Cells(ln, col + 3) = vDados
        
        vNome = vbNullString
        vDados = vbNullString
        vSituacao = vbNullString
    
        Ie.Document.getelementbyid("btnVoltar").Click
        
    Else
    
        Ie.navigate "https://www.situacaocadastral.com.br/"
        W.Cells(ln, col + 1) = "Dados inválidos para consulta"
        
    End If
    
    ln = ln + 1
    
    Application.Wait TimeSerial(Hour(Now()), Minute(Now()), Second(Now()) + vSegundos)
    
Loop

Ie.Quit

W.UsedRange.EntireColumn.AutoFit

Application.ScreenUpdating = True

DoEvents
MsgBox "Consulta realizada com sucesso!"

Set Ie = Nothing
Set UltCel = Nothing
Set W = Nothing
End Sub
