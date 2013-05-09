'Option Explicit'
Const fsoForReading = 1
Const fsoForWriting = 2
Const newest_path = "newest_date.txt"
Const semaphore_file = "running.txt"
Const first_run_file = "first_run.txt"
Const logProcess = "backtrack.log"
Const logFiles = "arquivos_movidos_FTP.log"
Const failed_attempts = "failed_attempts.txt"
Const attempts_to_delete_semaphore = 12
Const executions = "executions.txt"
Const executions_to_send_email = 288

Const email_to = "geraldo.junior@gonow.com.br,renato.berce@ftd.com.br,haline.camasmie@ftd.com.br"
Const email_from = "no-reply@nfe.ftd.com.br"
Const smtp_server = "ftdmail.ftd.com.br"
Const smtp_port = 25

Const origin_path = "E:\VBScript\files\"    'caminho da pasta onde estão os xmls'
Const destiny_path = "E:\VBScript\files\"     'caminho da pasta onde devem ser enviados os arquivos'
'On Error Resume Next'

'Execução de script'
'incrementa número de execuções'
executionsCounter()

'condição para execução de processamento das notas'
If Not isRunning(semaphore_file) Then
    writeFileText logProcess, "-----------------------------------------------------------------------------------------------------------"
    writeFileText logProcess, "start script"
    'zera as tentativas falhas'
    resetAttempts()
    'criar semáforo para não permitir que execute de novo'
    createSemaphore(semaphore_file)
    'condição para execução da primeira vez'
    If Not alreadyRunOnce(first_run_file) Then
        writeFileText logProcess, "first time execution"
        firstMove origin_path, destiny_path, load_time_from_file(newest_path)
        firstRun first_run_file
    'condição para execução das demais vezes'
    Else
        writeFileText logProcess, "other times executions"
        otherMoves origin_path, destiny_path, load_time_from_file(newest_path)
    End If
    'apaga semáforo'
    deleteSemaphore(semaphore_file)
    writeFileText logProcess, "end script"
'condição de bloqueio de processamento das notas'
Else
    'incrementa a tentativa falha'
    incrementAttempt()
    'apaga os semáforos se 12 tentativas falhas e zera tentativas falhas'
    ' Wscript.echo "failed_attempts" & loadDataFile(failed_attempts)
    If ((loadDataFile(failed_attempts) Mod attempts_to_delete_semaphore) = 0 And alreadyRunOnce(first_run_file)) Then
        ' Wscript.echo "apaga semáforo"
        writeFileText logProcess, "Running yet"
        deleteSemaphore(semaphore_file)
        writeFileText logProcess, "force semaphore delete"
        resetAttempts()
    End If
End If
'envia diariamente log de arquivos processados'
dailyReportSender()


Function dailyReportSender()
    If (loadDataFile(executions) Mod 10) = 0 Then
        sendMail()
    End If 
End Function

Function executionsCounter()
    Dim count
    If Not existsFile(executions) Then
        createFile(executions)
        writeCountFile executions, "1"
    Else
        count = loadDataFile(executions)
        count = count + 1
        deleteFile(executions)
        writeCountFile executions, count
    End If
End Function

Function resetAttempts()
    deleteFile(failed_attempts)
    createFile(failed_attempts)
    writeCountFile failed_attempts, "0"
End Function

Function incrementAttempt()
    Dim count
    ' Wscript.echo "entrou na função"
    If Not existsFile(failed_attempts) Then
        createFile(failed_attempts)
        writeCountFile failed_attempts, "1"
    Else
        count = loadDataFile(failed_attempts)
        count = count + 1
        deleteFile(failed_attempts)
        writeCountFile failed_attempts, countsssss
    End If
End Function

Function loadDataFile(filename)
    Dim fso, f
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.OpenTextFile(filename, fsoForReading)
    loadDataFile = f.ReadAll
End Function

Function save_data_file(filename, data)
    Dim fso, f
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.OpenTextFile(filename, fsoForWriting)
    f.Write data
End Function

Function writeFileText(sFilePath, sText)
    Const ForAppending = 8

    If Not existsFile(sFilePath) Then
        createFile(sFilePath)
    End if

    Set objFileSystem = CreateObject("Scripting.FileSystemObject")
    Set objOutputFile = objFileSystem.OpenTextFile(sFilePath, ForAppending)

    objOutputFile.WriteLine(Now & " - " & sText)

    objOutputFile.Close
End Function

Function writeCountFile(sFilePath, sText)
    If Not existsFile(sFilePath) Then
        createFile(sFilePath)
    End if

    Set objFileSystem = CreateObject("Scripting.FileSystemObject")
    Set objOutputFile = objFileSystem.OpenTextFile(sFilePath, 8)

    objOutputFile.WriteLine(sText)
    objOutputFile.Close
End Function

Function alreadyRunOnce(filename)
    alreadyRunOnce = existsFile(filename)
End Function

Function deleteSemaphore(filename)
    deleteFile filename
    writeFileText logProcess, "Semaphore deleted"
End Function

Function isRunning(filename)
    isRunning = existsFile(filename)
End Function

Function createSemaphore(filename)
    createFile filename
    writeFileText logProcess, "semaphore created"
End Function

Function firstRun(filename)
    createFile filename
    writeFileText logProcess, "first time semaphore created"
End Function

Function createFile(filename)
    Dim objFSO, objFile
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFSO.CreateTextFile(filename)
End Function

Function deleteFile(filename)
    dim filesys
    Set filesys = CreateObject("Scripting.FileSystemObject")
    If filesys.FileExists(filename) Then
        filesys.DeleteFile filename
    End If
End Function

Function existsFile(filename)
    dim filesys
    Set filesys = CreateObject("Scripting.FileSystemObject")
    existsFile = filesys.FileExists(filename)
End Function

'Salva no arquivo last_run.txt o tempo de início da execução'
Function save_datetime_in_file(filename, run_time)
    Dim fso, f
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.OpenTextFile(filename, fsoForWriting)
    f.Write run_time
End Function

'Lê qual foi a última vez que executou o script'
Function load_time_from_file(filename)
    Dim fso, f
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.OpenTextFile(filename, fsoForReading)
    load_time_from_file = CDate(f.ReadAll)
End Function

Function otherMoves(origin_path, destiny_path, newest)
    Dim folder, file, fileCollection, folderCollection, subFolder, fso, strTempSource, intCompare
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(origin_path)
    Set strTempSource = fso.GetFolder(folder)
    Set fileCollection = folder.Files
    For Each file In fileCollection
        intCompare = StrComp("Producao", folder.name, vbTextCompare)
        If intCompare = 0 Then
            If file.DateCreated >= load_time_from_file(newest_path) Then
                If Not existsFile(destiny_path & file.ShortName) Then
                    ' Wscript.echo "destiny_path: " & destiny_path & file.ShortName
                    writeFileText logProcess, "data de modificação mais recente: " & file.DateCreated
                    writeFileText logFiles, "arquivo movido: " & file.Name
                    newest = file.DateCreated
                    fso.CopyFile file.Path, destiny_path, TRUE
                    save_datetime_in_file newest_path, newest
                End If
            End If
        End If
    Next
    Set folderCollection = strTempSource.SubFolders
    For Each subFolder In folderCollection
        otherMoves subFolder.Path, destiny_path, newest
    Next
End Function

'primeira execução para saber qual arquivo é o mais antigo'
Function firstMove(origin_path, destiny_path, newest)
    Dim folder, file, fileCollection, folderCollection, subFolder, fso, strTempSource, intCompare
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(origin_path)
    Set strTempSource = fso.GetFolder(folder)
    Set fileCollection = folder.Files
    For Each file In fileCollection
        intCompare = StrComp("Producao", folder.name, vbTextCompare)
        If intCompare = 0 Then
            If file.DateCreated >= newest Then
                ' writeFileText logProcess, "------------------------------------------------------------------------------------------------------------------------"
                ' writeFileText logProcess, "data arquivo: " & file.DateCreated
                ' writeFileText logProcess, "data base:    " & newest
                newest = file.DateCreated
                writeFileText logProcess, "novo newest: " & newest & " from file: " & file.Path
                ' writeFileText logProcess, "------------------------------------------------------------------------------------------------------------------------"
                'salva a data do mais novo no arquivo de ref'
                save_datetime_in_file newest_path, FormatDateTime(newest)
            End If
            If Not existsFile(destiny_path & file.ShortName) Then
                ' Wscript.echo "destiny_path: " & destiny_path & file.ShortName
                fso.CopyFile file.Path, destiny_path, TRUE
                writeFileText logFiles, "arquivo movido: " & file.Name
            End If
        End If
    Next
    Set folderCollection = strTempSource.SubFolders
    For Each subFolder In folderCollection
        firstMove subFolder.Path, destiny_path, newest
    Next
End Function

Function sendMail()
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
    objMessage.AddAttachment fullPath & "\" & logFiles

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