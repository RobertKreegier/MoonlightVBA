VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLinkedCache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************
Option Explicit
Private Const blnDeveloperMode As Boolean = False
Private Const strModuleName As String = "CLinkedCache"
'**** Author  : Robert Kreegier
'**** Purpose : A class for managing a linked list cache of records comprised of a name
'               and a value.
'**** Notes   : This uses a linked list to implement a cache. When a record is
'               retrieved, we look down through the list until we find it, then we move
'               the record to the head so it's easier to find next time. When a new
'               record is added, it's linked at the head of the list, assuming we'll be
'               asking for it again soon. With these two behaviors, commonly used
'               records will stay at the top of the list, reducing query time. Also,
'               because this is a dynamically expanding linked list, it will continue
'               to keep track of any previous queries for as long as the program is
'               running.
'
'               The CLinkedCache class relies on a companion class called SNode, which
'               has all of four lines:
'                Option Explicit
'                Public strHash As String
'                Public strName As String
'                Public strValue As String
'                Public ndNextNode As SNode
'
'               If this class is not included in your project, you must create it.
'
'**** Use     : Create an instance of the cache object:
'****
'****               Dim cCache As CLinkedCache
'****
'****           Make sure it's initialized:
'****
'****               If cCache Is Nothing Then Set cCache = New CLinkedCache
'****
'****           Add a record to the cache:
'****
'****               cCache.AddRecord "Company Name", "D25"
'****
'****           Later in the procedure, use it to look up a record:
'****
'****               Dim strInfo As String
'****               strInfo = Sheets(1).Range(cCache.GetRecord("Company Name")).Value
'***************************************************************************************

#Const blnUseHashing = False    ' This is an experimental feature to use hashes for keys to avoid matching large strings and potentially decrease lookup times.
Private ndHead As SNode

'*******************************************************************************
' Author  : Robert Kreegier
' Purpose : Add a new node to the list at the head.
'*******************************************************************************
Public Sub AddRecord(ByVal strName As String, ByVal strValue As String)
    On Error GoTo ProcException
    '********************************************************************************
    
    Dim ndNewNode As SNode: Set ndNewNode = New SNode
    
    If Not RecordExists(strName) Then
        ' Fill in the data for this new node.
        #If blnUseHashing Then
            ndNewNode.strHash = Hash4(strName)
        #End If
        
        ndNewNode.strName = strName
        ndNewNode.strValue = strValue
        
        ' Set the new node as ndHead and the old head as ndNextNode
        Set ndNewNode.ndNextNode = ndHead
        Set ndHead = ndNewNode
    End If
    
    '********************************************************************************
ExitProc:
    Exit Sub
ProcException:
    CacheThrowError strModuleName, "AddRecord"
    Resume ExitProc
End Sub

'*******************************************************************************
' Author  : Robert Kreegier
' Purpose : Our error handling procedure.
'*******************************************************************************
Private Sub CacheThrowError(ByVal strModule As String, ByVal strProcedure As String)
    If Not blnDeveloperMode Then
        If strModule <> vbNullString Then
            strModule = strModule & ": "
        End If
        
        If strProcedure <> vbNullString Then
            strModule = strModule & strProcedure & ", "
            
            If Erl > 0 Then
                strModule = strModule & "Line: " & Erl
            End If
        End If
        
        If strModule <> vbNullString Then
            strModule = strModule & Chr(10) & Chr(10)
        End If
        
        Dim strErrorMessage As String
        MsgBox strModule & Err.Description
    Else
        Err.Raise vbObjectError + 1000, "ERROR", strErrorMessage
    End If
End Sub

'*******************************************************************************
' Author  : Robert Kreegier
' Purpose : Returns a total count of all the nodes in the cache.
'*******************************************************************************
Public Function Count() As Long
    On Error GoTo ProcException
    '********************************************************************************
    
    Dim ndNode As SNode
    
    Set ndNode = ndHead
    
    While Not ndNode Is Nothing
        Count = Count + 1
        Set ndNode = ndNode.ndNextNode
    Wend
    
    '********************************************************************************
ExitProc:
    Exit Function
ProcException:
    CacheThrowError strModuleName, "Count"
    Resume ExitProc
End Function

'*******************************************************************************
' Author  : Robert Kreegier
' Purpose : Prints out a list of the nodes and their values to the Immediate
'           Window for debugging purposes.
'*******************************************************************************
Public Sub DebugRecords()
    On Error GoTo ProcException
    '********************************************************************************
    
    Dim ndNode As SNode: Set ndNode = ndHead
    Dim lngCount As Long: lngCount = 0
    
    While Not ndNode Is Nothing
        lngCount = lngCount + 1
        Debug.Print lngCount & ": " & ndNode.strHash & " - " & ndNode.strName & " - " & ndNode.strValue
        Set ndNode = ndNode.ndNextNode
    Wend
    
    '********************************************************************************
ExitProc:
    Exit Sub
ProcException:
    CacheThrowError strModuleName, "DebugRecords"
    Resume ExitProc
End Sub

'*******************************************************************************
' Author  : Robert Kreegier
' Purpose : Flushes the cache back to zero nodes.
'*******************************************************************************
Public Sub Flush()
    On Error GoTo ProcException
    '********************************************************************************
    
    Set ndHead = Nothing
    
    '********************************************************************************
ExitProc:
    Exit Sub
ProcException:
    CacheThrowError strModuleName, "Flush"
    Resume ExitProc
End Sub

'*******************************************************************************
' Author  : Robert Kreegier
' Purpose : Gets a record from the node list.
'*******************************************************************************
Public Function GetRecord(ByVal strNameQuery As String) As String
    On Error GoTo ProcException
    '********************************************************************************
    
    ' Return vbNullString by default
    GetRecord = vbNullString
    
    Dim ndNode As SNode: Set ndNode = ndHead
    Dim ndPrevNode As SNode: Set ndPrevNode = Nothing
    Dim strHashQuery As String
    
    #If blnUseHashing Then
        strHashQuery = Hash4(strNameQuery)
    #End If
    
    ' Find the record we're looking for and swap it with the head so it's easier to
    ' find next time.
    While Not ndNode Is Nothing
        #If blnUseHashing Then
        If ndNode.strHash = strHashQuery Then
        #End If
            If ndNode.strName = strNameQuery Then
                ' Return the value we found
                GetRecord = ndNode.strValue
                
                ' Move the node up to the head if it's not already the head.
                If Not ndNode Is ndHead And Not ndPrevNode Is Nothing Then
                    ' Remove ndNode from where it was in the list by making the previous node link to
                    ' the next node.
                    Set ndPrevNode.ndNextNode = ndNode.ndNextNode
    
                    ' Set ndNode as the head
                    Set ndNode.ndNextNode = ndHead
                    Set ndHead = ndNode
                End If
    
                GoTo ExitProc
            End If
        #If blnUseHashing Then
        End If
        #End If
        
        Set ndPrevNode = ndNode
        Set ndNode = ndNode.ndNextNode
    Wend
    
    '********************************************************************************
ExitProc:
    Exit Function
ProcException:
    CacheThrowError strModuleName, "GetRecord"
    Resume ExitProc
End Function

'*******************************************************************************
' Author  : Robert Kreegier
' Purpose : Hash function used internally to speed up average seek time by
'           by reducing the comparison time.
'*******************************************************************************
Private Function Hash4(ByVal txt As String) As String
    On Error GoTo ProcException
    '********************************************************************************
    
    Dim mask, j, nC, crc As Integer
    Dim C As String
    
    crc = &HFFFF
    
    For nC = 1 To Len(txt)
        j = asc(Mid(txt, nC))
        crc = crc Xor j
        For j = 1 To 8
            mask = 0
            If crc / 2 <> Int(crc / 2) Then mask = &HA001
            crc = Int(crc / 2) And &H7FFF: crc = crc Xor mask
        Next j
    Next nC
    
    C = Hex$(crc)
    
    While Len(C) < 4
      C = "0" & C
    Wend
    
    Hash4 = C
    
    '********************************************************************************
ExitProc:
    Exit Function
ProcException:
    CacheThrowError strModuleName, "Hash4"
    Resume ExitProc
End Function

'*******************************************************************************
' Author  : Robert Kreegier
' Purpose : Returns true/false if the specified record exists.
'*******************************************************************************
Public Function RecordExists(ByVal strNameQuery As String) As Boolean
    On Error GoTo ProcException
    '********************************************************************************
    
    Dim ndNode As SNode: Set ndNode = ndHead
    Dim strHashQuery As String
    
    #If blnUseHashing Then
    strHashQuery = Hash4(strNameQuery)
    #End If
    
    While Not ndNode Is Nothing
        #If blnUseHashing Then
        If ndNode.strHash = strHashQuery Then
        #End If
            If ndNode.strName = strNameQuery Then
                RecordExists = True
                GoTo ExitProc
            End If
        #If blnUseHashing Then
        End If
        #End If
        
        Set ndNode = ndNode.ndNextNode
    Wend
    
    '********************************************************************************
ExitProc:
    Exit Function
ProcException:
    CacheThrowError strModuleName, "RecordExists"
    Resume ExitProc
End Function

'*******************************************************************************
' Author  : Robert Kreegier
' Purpose : Remove a node from the list.
'*******************************************************************************
Public Sub RemoveRecord(ByVal strName As String)
    On Error GoTo ProcException
    '********************************************************************************
    
    Dim ndNode As SNode: Set ndNode = ndHead
    Dim prev As SNode
    
    While Not ndNode Is Nothing
        ' If the node we're looking at is the node we're looking for, then...
        If ndNode.strName = strName Then
            ' Remove the node.
            If ndNode Is ndHead Then
                Set ndHead = ndNode.ndNextNode
            Else
                Set prev.ndNextNode = ndNode.ndNextNode
            End If
            
            Set ndNode = Nothing
            
            GoTo ExitProc
        End If
        
        ' Move on to the next node in the list.
        Set prev = ndNode
        Set ndNode = ndNode.ndNextNode
    Wend
    
    '********************************************************************************
ExitProc:
    Exit Sub
ProcException:
    CacheThrowError strModuleName, "RemoveRecord"
    Resume ExitProc
End Sub
