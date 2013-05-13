'Option Explicit'

' Nomes de arquivos
Const first_exec = "first_run.lock"
Const semaphore = "running.lock"
Const newest_date = "newest_date.txt"
Const failed_attempts = "failed_attempts.txt"
Const executions = "executions_count.txt"
Const logFile = "log.txt"
Const movedFiles = "arquivos_movidos_FTP.txt"

' variáveis
Const attempts_to_delete_semaphore = 6
Const daily_executions = 144

' email
Const email_to = "geraldo.junior@gonow.com.br,renato.berce@ftd.com.br,haline.camasmie@ftd.com.br"
Const email_from = "no-reply@nfe.ftd.com.br"
Const smtp_server = "ftdmail.ftd.com.br"
Const smtp_port = 25

' Paths dos arquivos
' Const origin_path = "E:\nfes\"
Const origin_path = "E:\VBScript\destino\"
Const destiny_path = "E:\VBScript\destino\"
'On Error Resume Next'

exec()

Function exec ()
    debug "call exec()"
    If Not existsFile(semaphore) Then
        ' executa primeira vez para carregar todos os arquivos
        If isFirstTime() Then
            firstExec()
            ' cria o marcador de primeira execuÃ§Ã£o
            createFile(first_exec)
        Else
            othersExecs()
        End If
        increaseFileCounter(executions)
        resetFailedAttempts()
    Else
        If existsFile(first_exec) Then
            debug "other process already executing"
            increaseFileCounter(failed_attempts)
            errorRecover() 
        Else
            debug "first process already executing"
        End If
    End If
    debug "end exec()"
    debug "-----------------------------------------------------------------"
End Function

Function isFirstTime ()
    ' debug "call isFirstTime"
    ' checa se existe o arquivo first_time
    isFirstTime = Not existsFile(first_exec)
End Function

Function firstExec ()
    debug "call firstExec"
    ' cria semÃ¡foro
    createSemaphore()
    ' processamento das notas
    fileProcess(origin_path)
    ' apaga semÃ¡foro
    deleteSemaphore()
    ' manda email com resultado
    sendMail()
End Function

Function othersExecs ()
    debug "call othersExecs"
    createSemaphore()
    ' processamento das notas
    fileProcess(origin_path)
    ' apaga o semáforo de execução
    deleteSemaphore()
End Function

Function fileProcess (filePath)
    ' debug "call fileProcess(" & filePath & ")"
    Dim folder, file, fileCollection, folderCollection, subFolder, fso, strTempSource, intCompare
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(filePath)
    Set strTempSource = fso.GetFolder(folder)
    Set fileCollection = folder.Files
    For Each file In fileCollection
        ' só trabalha com os arquivos dentro da pasta "Producao"
        intCompare = StrComp("Producao", folder.name, vbTextCompare)
        If intCompare = 0 Then
            ' Se arquivo encontrado for mais novo que a data no arquivo então atualiza o arquivo
            If file.DateCreated >= getNewestDate() Then
                debug "encontrada uma data de criação mais recente: " & file.DateCreated & " do arquivo " & file.Path
                setNewestDate(file.DateCreated)
                copyFile file.Path, destiny_path & file.ShortName
            End If
            ' Se for primeira vez copia todos os arquivos
            If isFirstTime Then
                copyFile file.Path, destiny_path & file.ShortName
            End If
        End If
    Next
    Set folderCollection = strTempSource.SubFolders
    ' recursividade para sub pastas
    For Each subFolder In folderCollection
        fileProcess(subFolder.Path)
    Next
End Function

Function copyFile (origin, destiny)
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not existsFile(destiny) Then
        fso.CopyFile origin, destiny_path, TRUE
        trace Now & " - arquivo copiado: " & origin
    End If
End Function

Function getNewestDate ()
    ' debug "call getNewestDate"
    ' WScript.echo "nome do arquivo: " & newest_date
    If Not existsFile(newest_date) Then
        createFile(newest_date)
        setNewestDate("01/01/1970 00:00:00")
    End If
    getNewestDate = loadDateFromFile(newest_date)
End Function

Function setNewestDate (newestDate)
    ' debug "call setNewestDate (" & newestDate & ")"
    If existsFile(newest_date) Then
        deleteFile(newest_date)
    End If
    ' WScript.echo "nome do arquivo: " & newest_date
    writeFile newest_date, newestDate
End Function

Function loadDateFromFile (fileName)
    ' debug "call loadDateFromFile (" & fileName & ")"
    Const fsoForReading = 1
    Dim fso, f
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.OpenTextFile(fileName, fsoForReading)
    strDate = CDate(f.ReadAll)
    ' Wscript.echo "conteúdo do arquivo: " & strDate
    loadDateFromFile = strDate
End Function

Function createSemaphore ()
    debug "call createSemaphore"
    createFile(semaphore)
End Function

Function deleteSemaphore ()
    debug "call deleteSemaphore"
    deleteFile(semaphore)
End Function

Function createFile (fileName)
    ' debug "call createFile (" & fileName & ")"
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFSO.CreateTextFile(fileName)
End Function

Function existsFile (fileName)
    ' debug "call existsFile (" & fileName & ")"
    Set filesys = CreateObject("Scripting.FileSystemObject")
    ' debug "file " & fileName & " exists: " & filesys.FileExists(fileName)
    existsFile = filesys.FileExists(fileName)
End Function

Function deleteFile (fileName)
    ' debug "call deleteFile (" & fileName & ")"
    Set filesys = CreateObject("Scripting.FileSystemObject")
    If filesys.FileExists(fileName) Then
        filesys.DeleteFile fileName
        ' debug "file " & fileName & " deleted"
    Else
        debug "file " & fileName & " doesn't exists"
    End If
End Function

Function writeFile (fileName, content)
    ' debug "call writeFile (" & fileName & "," & content & ")"
    Const ForAppending = 8

    If Not existsFile(fileName) Then
        createFile(fileName)
    End if

    Set objFileSystem = CreateObject("Scripting.FileSystemObject")
    Set objOutputFile = objFileSystem.OpenTextFile(fileName, ForAppending)

    objOutputFile.WriteLine(content)

    objOutputFile.Close
End Function

Function readFile (fileName)
    Const fsoForReading = 1
    ' debug "call readFile (" & fileName & ")"
    Dim fso, f
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.OpenTextFile(fileName, fsoForReading)
    readFile = f.ReadAll
End Function

Function increaseFileCounter (fileName)
    If Not existsFile(fileName) Then
        createFile(fileName)
        writeFile fileName, "1"
    Else
        count = readFile(fileName)
        count = count + 1
        deleteFile(fileName)
        writeFile fileName, count
    End If
End Function

Function resetFailedAttempts ()
    deleteFile(failed_attempts)
End Function

Function sendMail ()
    debug "call sendMail()"
    Const cdoSendUsingPickup = 1 'Send message using the local SMTP service pickup directory. 
    Const cdoSendUsingPort = 2 'Send the message using the network (SMTP over the network). 

    Const cdoAnonymous = 0 'Do not authenticate
    Const cdoBasic = 1 'basic (clear-text) authentication
    Const cdoNTLM = 2 'NTLM
    ' Const email_to = "geraldo.junior@gonow.com.br,geraldo.guiesi@gmail.com"
    ' Const email_from = "no-reply@nfe.ftd.com.br"
    ' Const smtp_server = "ftdmail.ftd.com.br"
    ' Const smtp_port = 25

    Set fso = CreateObject("Scripting.FileSystemObject") 
    fullPath = fso.GetParentFolderName(wscript.ScriptFullName)

    Set objMessage = CreateObject("CDO.Message") 
    objMessage.Subject = "Relatório diário de notas processadas" 
    objMessage.From = email_from
    objMessage.To = email_to
    objMessage.TextBody = "Seguem os logs de processamento das notas ficais."
    If existsFile(fullPath & "\" & logFile) Then
        objMessage.AddAttachment fullPath & "\" & logFile
    End If
    If existsFile(fullPath & "\" & movedFiles) Then
        objMessage.AddAttachment fullPath & "\" & movedFiles
    End If

    objMessage.Configuration.Fields.Item _
    ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 

    'Name or IP of Remote SMTP Server
    objMessage.Configuration.Fields.Item _
    ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = smtp_server

    'Type of authentication, NONE, Basic (Base64 encoded), NTLM
    objMessage.Configuration.Fields.Item _
    ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic

    'Server port (typically 25)
    objMessage.Configuration.Fields.Item _
    ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25

    'Use SSL for the connection (False or True)
    objMessage.Configuration.Fields.Item _
    ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False

    'Connection Timeout in seconds (the maximum time CDO will try to establish a connection to the SMTP server)
    objMessage.Configuration.Fields.Item _
    ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60

    objMessage.Configuration.Fields.Update

    '==End remote SMTP server configuration section==
    objMessage.Send
End Function

Function trace (message)
    ' WScript.echo message
    writeFile movedFiles, message
End Function

Function debug (message)
    ' WScript.echo message
    writeFile logFile, Now & " - " & message
End Function

Function errorRecover ()
    failures = readFile(failed_attempts)
    If (failures Mod attempts_to_delete_semaphore) = 0 And Not isFirstTime() Then
        debug "error recover - erase semaphore"
        deleteSemaphore()
    End If
End Function

Function dailyReport ()
    debug "send daily Report"
    If (readFile(executions) Mod daily_executions) = 0 Then
        sendMail()
    End If
End Function