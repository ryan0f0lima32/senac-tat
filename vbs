'Atividade do Leandro

https://www.youtube.com/watch?v=ecKij3K6jYs

'PARAMOS EM 9:20

'--------------------------------------------------

'Fiz mapeamento de pastas
'Início de programas
'Montei uma impressora fictícia
'Mapeamento da impressora fictícia

'Apagar os scripts feitos dos grupos Alunos e Professores
'Talvez apagar a impressora fictícia

'Fazer isso via VBS
'Para isso, assistir ao vídeo acima

'--------------------------------------------------

'Ignora erros no script
On Error Resume Next
Err.Clear

'Cria objetos necessários
Set objNetwork = CreateObject("WScript.Network")
Set FSODrive = CreateObject("Scripting.FileSystemObject")

'Obtém domínio e usuário
strDom = objNetwork.UserDomain
strUser = objNetwork.UserName

'Obtém informações do usuário
Set objUser = GetObject("WinNT://" & strDom & "/" & strUser & ",user")

'Verifica os grupos do usuário
For Each objGroup In objUser.Groups

    Select Case objGroup.Name

        Case "Alunos"

            If Not FSODrive.DriveExists("G:") Then
                objNetwork.MapNetworkDrive "G:", "\\ryan\Aulas", True
            End If

            If Not FSODrive.DriveExists("H:") Then
                objNetwork.MapNetworkDrive "H:", "\\ryan\Alunos", True
            End If

            If Not FSODrive.DriveExists("I:") Then
                objNetwork.MapNetworkDrive "I:", "\\ryan\Trabalhos", True
            End If

            MsgBox "Bem-vindo ao sistema, Aluno(a)! Bons estudos.", 64, "Logon Alunos"

            objNetwork.AddWindowsPrinterConnection "\\ryan\HP"

        Case "Professores"

            If Not FSODrive.DriveExists("G:") Then
                objNetwork.MapNetworkDrive "G:", "\\ryan\Aulas", True
            End If

            If Not FSODrive.DriveExists("H:") Then
                objNetwork.MapNetworkDrive "H:", "\\ryan\Alunos", True
            End If

            If Not FSODrive.DriveExists("I:") Then
                objNetwork.MapNetworkDrive "I:", "\\ryan\Trabalhos", True
            End If

            If Not FSODrive.DriveExists("J:") Then
                objNetwork.MapNetworkDrive "J:", "\\ryan\Provas", True
            End If

            MsgBox "Bem-vindo(a), Professor(a)! O sistema de aulas está pronto.", 64, "Logon Professores"

            objNetwork.AddWindowsPrinterConnection "\\ryan\Impressora-Da-Rede"

    End Select

Next

'Mapeia pasta pública
objNetwork.MapNetworkDrive "P:", "\\ryan\Publica", True
