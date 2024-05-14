' Função para criar um diretório e sobrescrever, se necessário
Sub CriarOuSobrescreverDiretorio(caminho)
    Dim objFSO
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    ' Verifique se o diretório existe
    If objFSO.FolderExists(caminho) Then
        ' Se o diretório existir, exclua-o
        objFSO.DeleteFolder caminho
    End If
    
    ' Crie o diretório
    objFSO.CreateFolder caminho
End Sub

' Caminho para a área de trabalho do usuário
Dim diretorio
diretorio = CreateObject("WScript.Shell").SpecialFolders("Desktop")

' Chame a função para criar ou sobrescrever o diretório
CriarOuSobrescreverDiretorio diretorio & "\pastaVBS"

' Mude o diretório atual para o novo diretório
Set objShell = CreateObject("WScript.Shell")
objShell.CurrentDirectory = diretorio & "\pastaVBS"

' Crie um arquivo de texto no diretório com o conteúdo
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.CreateTextFile("texto.txt")
objFile.Write "mensagem do arquivo"
objFile.Close

MsgBox "Voce foi vitima de um hacker", vbExclamation, "Aviso"
