'***************************************************************************************
Option Explicit
#Const blnDeveloperMode = False
Private Const strModuleName As String = "CONC_FILES"
'**** Author  : Robert M Kreegier
'**** Purpose : Procedures for creating and handling .conc files
'**** Notes   : If you store your workbook on a shared drive or cloud drive, it's
'****           possible for two people to open up the file at the same time. As a
'****           result, concurrency becomes an issue. To resolve this these procedures
'****           can be used to write a small file with your username and a timestamp.
'****           That file (.conc) is then synced to the shared drive. When someone
'****           tries to open your workbook, you can have it check for the .conc file.
'****           If there's one in the same directory as the workbook and it has a
'****           different username than the current user, then you know someone else
'****           is already in the file.
'****           There can be occations where a workbook crashes and leaves the .conc
'****           file behind. In that case, the code here looks to ignore and clean up
'****           .conc files older than 12 hours (720 minutes).
'****           That time limit can be changed directly below with the constant
'****           "intConcTimeout".
'***************************************************************************************

Private Const intConcTimeout = 720

'***************************************************************************************************
' Author  : Robert Kreegier
' Purpose : This checks to see if there's a concurrency file that exists for this workbook and user.
'           It will also throw True if the file doesn't exist at all. Basically, all we want to know
'           is if we're good to edit and save. If there's a conc file with someone else's name on
'           it, then we have to throw False.
' In other words:
'   True = CheckConcFile = Good to edit and save
'   False = CheckConcFile = Someone else is in the file
' Params  :
'   strFileName     This should be the filename of the workbook we want to check
'***************************************************************************************************
Function CheckConcFile(Optional ByVal strFileName As String = vbNullString) As Boolean
    #If Not blnDeveloperMode Then
        On Error GoTo Whoa
    #End If
    '********************************************************************************
    
    CheckConcFile = True
    
    ' Sharepoint makes ThisWorkbook return a web address for FullName. When that's the case, the namespace is obviously different and we can't reliably
    ' write conc files. It's all good, however, because Sharepoint takes care of concurency issues.
    ' So, if FullName has a ":\" in the second position, then we can safetly bet that we're looking at a file on disk.
    If InStr(1, ThisWorkbook.FullName, ":\") = 2 Then
        ' if a file name wasn't given, then we'll use the filename for the current workbook
        If strFileName = vbNullString Then
            strFileName = ThisWorkbook.FullName
        End If
    
        ' remove the file extension and replace it with ".conc"
        strFileName = Left(strFileName, InStrRev(strFileName, ".") - 1) & ".conc"
        
        ' check if the conc file exists
        If FileExists(strFileName) Then
            Open strFileName For Input As #1
            
            Dim strData As String
            While Not EOF(1)
                Line Input #1, strData             ' get the file name
            Wend
            Close #1
            
            ' separate the username from the timestamp
            Dim strUserName As String: strUserName = Trim(Left(strData, Len(strData) - 22)) ' knock 22 chars off the end and trim the whitespace to get the username
            Dim strTimeStamp As String: strTimeStamp = Trim(Right(strData, 22))             ' get the last 22 chars and trim the whitespace to get the timestamp
            
            ' if the file is younger than 12 hours
            If DateDiff("n", TimeStampToDate(strTimeStamp), GetTimeStamp) < 720 Then
                ' and the username is not the current user
                If InStr(strUserName, Application.UserName) <= 0 Then
                    CheckConcFile = False
                End If
                
            ' if the file is older than 12 hours, then let's delete it
            Else
                Kill strFileName
            End If
        End If
    End If
    
    '********************************************************************************
    #If Not blnDeveloperMode Then
Letscontinue:
        Exit Function
Whoa:
        InfoBox Err.Description & Chr(10) & "thrown from " & strModuleName & ": CheckConcFile", True
        Resume Letscontinue
    #End If
End Function

'***************************************************************************************************
' Author  : Robert Kreegier
' Purpose : This creates a concurrency file in the same directory as this workbook using this
'           workbook's filename.
'   True = CreateConcFile = Good to edit and save, conc file created
'   False = CreateConcFile = Someone else is in the file, no conc file created
'***************************************************************************************************
Function CreateConcFile(Optional ByVal strFileName As String = vbNullString) As Boolean
    #If Not blnDeveloperMode Then
        On Error GoTo Whoa
    #End If
    '********************************************************************************
    
    CreateConcFile = False
    
    ' Sharepoint makes ThisWorkbook return a web address for FullName. When that's the case, the namespace is obviously different and we can't reliably
    ' write conc files. It's all good, however, because Sharepoint takes care of concurency issues...supposedly
    ' So, if FullName has a ":\" in the second position, then we can safetly bet that we're looking at a file on disk.
    If InStr(1, ThisWorkbook.FullName, ":\") = 2 Then
        ' if a file name wasn't given, then we'll use the file name for the current workbook
        If strFileName = vbNullString Then
            strFileName = ThisWorkbook.FullName
        End If
        
        ' remove the file extension and replace it with ".conc"
        strFileName = Left(strFileName, InStrRev(strFileName, ".") - 1) & ".conc"
        
        ' if this file already exists, then we know someone is in this AHR. let's read out who it is...
        If FileExists(strFileName) Then
            Dim strData As String
            
            ' get the string
            Open strFileName For Input As #1
            While Not EOF(1)
                Line Input #1, strData
            Wend
            Close #1
            
            ' separate the username from the timestamp
            Dim strUserName As String: strUserName = Trim(Left(strData, Len(strData) - 22)) ' knock 22 chars off the end and trim the whitespace to get the username
            Dim strTimeStamp As String: strTimeStamp = Trim(Right(strData, 22))             ' get the last 22 chars and trim the whitespace to get the timestamp
            
            ' if the file is younger than 12 hours
            If DateDiff("n", TimeStampToDate(strTimeStamp), GetTimeStamp) < intConcTimeout Then
                ' and the username is not the current user
                If InStr(strUserName, Application.UserName) <= 0 Then
                    InfoBox "It looks like " & strUserName & " is already in this AHR. " & strUserName & " must close out their session before you can edit this file. Reopen this file after they're out."
                
                Else
                    ' create and open the file name
                    Open strFileName For Output As #1
                    
                    ' overwrite the user's name and time
                    Print #1, Application.UserName & "   " & GetTimeStamp
                    
                    ' close the file
                    Close #1
                    
                    CreateConcFile = True
                End If
    
            ' if the file has been around longer than 12 hours, we're going to presume the person is out of this AHR
            Else
                ' create and open the file name
                Open strFileName For Output As #1
                
                ' overwrite the user's name and time
                Print #1, Application.UserName & "   " & GetTimeStamp
                
                ' close the file
                Close #1
                
                CreateConcFile = True
            End If
    
        ' file doesn't exist, so let's make one
        Else
            ' create and open the file name
            Open strFileName For Output As #1
            
            ' write the user's name and time
            Print #1, Application.UserName & "   " & GetTimeStamp
            
            ' close the file
            Close #1
            
            CreateConcFile = True
        End If
        
    Else
        CreateConcFile = True
    End If
    
    '********************************************************************************
    #If Not blnDeveloperMode Then
Letscontinue:
        Exit Function
Whoa:
        InfoBox Err.Description & Chr(10) & "thrown from " & strModuleName & ": CreateConcFile", True
        Resume Letscontinue
    #End If
End Function

'***************************************************************************************************
' Author  : Robert Kreegier
' Purpose : This deletes the concurrency file associated with this workbook
'***************************************************************************************************
Sub DeleteConcFile(Optional ByVal strFileName As String = vbNullString)
    #If Not blnDeveloperMode Then
        On Error GoTo Whoa
    #End If
    '********************************************************************************
    
    ' Sharepoint makes ThisWorkbook return a web address for FullName. When that's the case, the namespace is obviously different and we can't reliably
    ' write conc files. It's all good, however, because Sharepoint takes care of concurency issues.
    ' So, if FullName has a ":\" in the second position, then we can safetly bet that we're looking at a file on disk.
    If InStr(1, ThisWorkbook.FullName, ":\") = 2 Then
        ' if a file name wasn't given, then we'll use the filename for the current workbook
        If strFileName = vbNullString Then
            strFileName = ThisWorkbook.FullName
        End If
        
        ' remove the file extension and replace it with ".conc"
        strFileName = Left(strFileName, InStrRev(strFileName, ".") - 1) & ".conc"
        
        If FileExists(strFileName) Then
            Open strFileName For Input As #1
            Dim strData As String
            While Not EOF(1)
                Line Input #1, strData             ' get the user name
            Wend
            Close #1
            
            ' separate the username from the timestamp
            Dim strUserName As String: strUserName = Trim(Left(strData, Len(strData) - 22)) ' knock 22 chars off the end and trim the whitespace to get the username
            Dim strTimeStamp As String: strTimeStamp = Trim(Right(strData, 22))             ' get the last 22 chars and trim the whitespace to get the timestamp
            
            ' if the file is younger than 12 hours
            If DateDiff("n", TimeStampToDate(strTimeStamp), GetTimeStamp) < 720 Then
                ' and the username is the current user
                If InStr(strUserName, Application.UserName) > 0 Then
                    ' delete the conc file
                    Kill strFileName
                End If
            
            ' the file was older than 12 hours, so let's delete it anyway
            Else
                Kill strFileName
            End If
        End If
    End If
    
    '********************************************************************************
    #If Not blnDeveloperMode Then
Letscontinue:
        Exit Sub
Whoa:
        InfoBox Err.Description & Chr(10) & "thrown from " & strModuleName & ": DeleteConcFile", True
        Resume Letscontinue
    #End If
End Sub
