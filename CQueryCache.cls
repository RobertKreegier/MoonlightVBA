VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CQueryCache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************
Option Explicit
#Const blnDeveloperMode = False
Private Const strModuleName As String = "CQueryCache"
'**** Author  : Robert M Kreegier
'**** Purpose : A class for managing a cache of records comprised of a name and a value.
'**** Notes   : In this class, the cache is represented with a two dimensional array of
'****           strings. One dimension is for the name of a record, and the second
'****           dimension is the value associated with that name.
'****           Every time a record is fetched from Cache with GetRecord, the record
'****           is promoted up the array (closer to position zero). This ensures that
'****           the fetching process will take one less loop the next time it needs to
'****           fetch the record. Therefore variables that are used more accumilate
'****           closer to the zeroth array position and are thus accessed quicker.
'****
'**** Use     : At the top of your module, create an instance of the cache object:
'****
'****               Public cCache As CQueryCache
'****
'****           This ensures that the cache will persist globally for awhile.
'****
'****           In the procedure it'll be used in, make sure it's initialized:
'****
'****               If cCache Is Nothing Then Set cCache = New CQueryCache
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

Public CacheDepth As Long

Private Const DefaultCacheDepth As Long = 80

Private Cache() As String

' Returns the depth of the cache
Public Property Get Depth() As Long
    #If Not blnDeveloperMode Then
        On Error GoTo ProcException
    #End If
    '********************************************************************************
    
    Depth = CacheDepth
    
    '********************************************************************************
    #If Not blnDeveloperMode Then
ExitProc:
        Exit Property
ProcException:
        MsgBox Err.Description & Chr(10) & "thrown from " & strModuleName & ": Get Depth"
        Resume ExitProc
    #End If
End Property

' Sets the depth of the cache
Public Property Let Depth(ByVal value As Long)
    #If Not blnDeveloperMode Then
        On Error GoTo ProcException
    #End If
    '********************************************************************************
    
    CacheDepth = value
    
    ' you can only ReDim Preserve the last dimension of an array, so if we want this Cache to be dynamic, then we'll have to rework this
    ReDim Cache(CacheDepth, 1) As String
    
    '********************************************************************************
    #If Not blnDeveloperMode Then
ExitProc:
        Exit Property
ProcException:
        MsgBox Err.Description & Chr(10) & "thrown from " & strModuleName & ": Let Depth"
        Resume ExitProc
    #End If
End Property

' Retrieve a record from the cache. If no record is found, then vbNullString is
' returned. Otherwise, the value associated with the name is returned.
Public Function GetRecord(ByVal strName As String) As String
    #If Not blnDeveloperMode Then
        On Error GoTo ProcException
    #End If
    '********************************************************************************
    
    ' Return vbNullString by default.
    GetRecord = vbNullString
    
    ' Look down through the Cache for our record.
    Dim lngIndex As Long
    For lngIndex = 0 To CacheDepth
        If Cache(lngIndex, 0) = strName Then
            ' we found it, so lets return it
            GetRecord = Cache(lngIndex, 1)
            
            ' since the variable was found, float it up a position
            PromoteRecord lngIndex
            Exit Function
        End If
    Next
    
    '********************************************************************************
    #If Not blnDeveloperMode Then
ExitProc:
        Exit Function
ProcException:
        MsgBox Err.Description & Chr(10) & "thrown from " & strModuleName & ": GetRecord"
        Resume ExitProc
    #End If
End Function

' Adds a new record to the next available spot in the cache closes to the top
Public Sub AddRecord(ByVal strName As String, ByVal strValue As String)
    #If Not blnDeveloperMode Then
        On Error GoTo ProcException
    #End If
    '********************************************************************************
    
    ' We want to add this variable in the first open spot in the cache closest to the top
    Dim intIndex As Integer
    For intIndex = 0 To CacheDepth
        If Cache(intIndex, 0) = vbNullString Then
            Exit For
        End If
    Next
    
    ' for some reason the For loop is allowing intIndex to go over CacheDepth
    If intIndex > CacheDepth Then intIndex = CacheDepth
    
    Cache(intIndex, 0) = strName
    Cache(intIndex, 1) = strValue

    '********************************************************************************
    #If Not blnDeveloperMode Then
ExitProc:
        Exit Sub
ProcException:
        MsgBox Err.Description & Chr(10) & "thrown from " & strModuleName & ": Add"
        Resume ExitProc
    #End If
End Sub

' Moves a record in the cache closer to position zero by one
Private Sub PromoteRecord(ByVal lngIndex As Long)
    #If Not blnDeveloperMode Then
        On Error GoTo ProcException
    #End If
    '********************************************************************************
    
    If lngIndex > 0 And lngIndex <= CacheDepth Then
        Dim lngPrev As Long: lngPrev = lngIndex - 1

        ' If Cache(lngIndex,0) is the same as the one above it, then we can just delete it
        If Cache(lngPrev, 0) = Cache(lngIndex, 0) Then
            Cache(lngIndex, 0) = vbNullString
            Cache(lngIndex, 1) = vbNullString
            
        ' Otherwise, swap it with the one above
        Else
            Dim strTemp0 As String: strTemp0 = Cache(lngIndex, 0)
            Dim strTemp1 As String: strTemp1 = Cache(lngIndex, 1)

            Cache(lngIndex, 0) = Cache(lngPrev, 0)
            Cache(lngIndex, 1) = Cache(lngPrev, 1)
            
            Cache(lngPrev, 0) = strTemp0
            Cache(lngPrev, 1) = strTemp1
        End If
    End If
    
    '********************************************************************************
    #If Not blnDeveloperMode Then
ExitProc:
        Exit Sub
ProcException:
        MsgBox Err.Description & Chr(10) & "thrown from " & strModuleName & ": PromoteRecord"
        Resume ExitProc
    #End If
End Sub

' Flushes the cache by redimming it without preserve
Public Sub Flush()
    #If Not blnDeveloperMode Then
        On Error GoTo ProcException
    #End If
    '********************************************************************************
    
    If IsMissing(CacheDepth) Or CacheDepth = 0 Then CacheDepth = DefaultCacheDepth
    
    ReDim Cache(CacheDepth, 1) As String
    
    '********************************************************************************
    #If Not blnDeveloperMode Then
ExitProc:
        Exit Sub
ProcException:
        MsgBox Err.Description & Chr(10) & "thrown from " & strModuleName & ": Flush"
        Resume ExitProc
    #End If
End Sub

' Dumps the cache to a file for debugging
Public Sub DumpCacheToFile(ByVal strFileName)
    #If Not blnDeveloperMode Then
        On Error GoTo ProcException
    #End If
    '********************************************************************************
    
    Dim strData As String
    Dim intIndex As Integer
    For intIndex = 0 To CacheDepth
        If Not Cache(intIndex, 0) = vbNullString Then
            If intIndex = 0 Then
                strData = intIndex & ":" & Cache(intIndex, 0) & "-" & Cache(intIndex, 1)
            Else
                strData = strData & Chr(10) & intIndex & ":" & Cache(intIndex, 0) & "-" & Cache(intIndex, 1)
            End If
        Else
            If intIndex = 0 Then
                strData = intIndex & ": - "
            Else
                strData = strData & Chr(10) & intIndex & ": - "
            End If
        End If
    Next

    Open strFileName For Output As #1

    Print #1, strData

    Close #1
    
    '********************************************************************************
    #If Not blnDeveloperMode Then
ExitProc:
        Exit Sub
ProcException:
        MsgBox Err.Description & Chr(10) & "thrown from " & strModuleName & ": DumpCacheToFile"
        Resume ExitProc
    #End If
End Sub

' The Init procedure (VBAs version of the constructor)
Private Sub Class_Initialize()
    If CacheDepth = 0 Then CacheDepth = DefaultCacheDepth
    
    ReDim Cache(CacheDepth, 1) As String
End Sub
