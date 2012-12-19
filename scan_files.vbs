Option Explicit
Const fsoForReading = 1
Const fsoForWriting = 2
Const origin_path= "C:\Documents and Settings\Administrador\Desktop\xml_path\"    'caminho da pasta onde estão os xmls'
Const destiny_path= "C:\Documents and Settings\Administrador\Desktop\to_transfer\"     'caminho da pasta onde devem ser enviados os arquivos'
Const newest_path = "newest_file.txt"
Const semaphore_file = "running.txt"
Const first_run_file = "first_run.txt"
On Error Resume Next

'Teste para validar se está rodando ou não'
Wscript.Echo "start script"
If Not isRunning(semaphore_file) Then
    'criar semáforo para não permitir que execute de novo'
    createSemaphore(semaphore_file)
    If Not almostRunOnce(first_run_file) Then
        firstRun first_run_file
        firstMove origin_path, destiny_path, load_time_from_file(newest_path)
        Wscript.Echo "first time execution"
    Else 
        otherMoves origin_path, destiny_path, load_time_from_file(newest_path)
        Wscript.Echo "other times executions"
    End If
    deleteSemaphore(semaphore_file)
Else 
    WScript.Echo "Running yet"
End If
Wscript.Echo "end script"

Function almostRunOnce(filename)
    almostRunOnce = existsFile(filename)
End Function

Function deleteSemaphore(filename)
    deleteFile filename
    Response.Write("Semaphore deleted")
End Function

Function isRunning(filename)
    isRunning = existsFile(filename)
End Function

Function createSemaphore(filename)
    createFile filename
    Response.Write("semaphore created")
End Function

Function firstRun(filename)
    createFile filename
    Response.Write("first time semaphore created")
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
    Dim folder, file, fileCollection, folderCollection, subFolder, fso, strTempSource
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(origin_path)
    Set strTempSource = fso.GetFolder(folder)
    Set fileCollection = folder.Files
    For Each file In fileCollection
        If file.DateLastModified > load_time_from_file(newest_path) Then
            Wscript.Echo "data de modificação mais recente: " & file.DateLastModified
            Wscript.Echo "arquivo movido"
            newest = file.DateLastModified
            fso.CopyFile file.Path, destiny_path, True
        End If
    Next
    Set folderCollection = strTempSource.SubFolders
    For Each subFolder In folderCollection
        otherMoves subFolder.Path, destiny_path, newest
    Next
    save_datetime_in_file newest_path, newest
End Function

'primeira execução para saber qual arquivo é o mais antigo'
Function firstMove(origin_path, destiny_path, newest)
    'Wscript.Echo "first newest: " & newest'
    Dim folder, file, fileCollection, folderCollection, subFolder, fso, strTempSource
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(origin_path)
    Set strTempSource = fso.GetFolder(folder)
    Set fileCollection = folder.Files
    For Each file In fileCollection
        If file.DateLastModified > newest Then
            Wscript.Echo "------------------------------------------------------------------------------------------------------------------------"
            Wscript.Echo "data arquivo: " & file.DateLastModified 
            Wscript.Echo "data base:    " & newest
            newest = file.DateLastModified
            Wscript.Echo "novo newest: " & newest & " from file: " & file.Path
            Wscript.Echo "------------------------------------------------------------------------------------------------------------------------"
        End If
        fso.CopyFile file.Path, destiny_path, True
    Next
    Set folderCollection = strTempSource.SubFolders
    For Each subFolder In folderCollection
        'Wscript.Echo "entrou na recursão - firstMove("& subFolder.Path & "," & destiny_path & "," & newest & ")"'
        firstMove subFolder.Path, destiny_path, newest
    Next
    'salva a data do mais novo no arquivo de ref'
    'Wscript.Echo "data mais recente: " & newest'
    save_datetime_in_file newest_path, FormatDateTime(newest)
End Function
