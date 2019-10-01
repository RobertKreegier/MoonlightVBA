Attribute VB_Name = "UTILS"
'***************************************************************************************************
Option Explicit
#Const blnDeveloperMode = False
Private Const strModuleName As String = "UTILS"
'**** Author : Robert M Kreegier
'**** Purpose: Utility Subs and Functions
'**** Notes:    This is a collection of utility functions that shouldn't be dependant on other
'****           modules. They should all be portable and easily usable in other applications.
'***************************************************************************************************

'***************************************************************************************************
' Author  : Chris Read, heavily modified by Robert Kreegier
' Purpose : Exports a selected range (table) to an xml format using table headers.
' Notes   : Requires entry of file name.
'           Uses table headers as Element names for each row.
' Source  : http://www.professionalexcel.com/2014/04/export-excel-table-range-to-xml-using-vba/
'***************************************************************************************************
Sub ExportRangeToXML(Optional ByRef rngSelection As Range, Optional ByVal strFileName As String = vbNullString)
    #If Not blnDeveloperMode Then
        On Error GoTo ProcException
    #End If
    '***********************************************************************************************
    
    Dim strXML As String
    Dim varTable As Variant
    Dim intRow As Integer
    Dim intCol As Integer
    Dim intFileNum As Integer
    Dim strRowElementName As String
    Dim strTableElementName As String
    Dim varColumnHeaders As Variant
    
    ' Set custom names
    strTableElementName = "Table"
    strRowElementName = "Row"
    
    ' check validity of strFileName, assign strFileName a zero-length string if not valid
    If FileExists(strFileName) And ExtractExtension(strFileName) = "xml" Then
        ' check if the file already exists
        If Dir(strFileName) <> vbNullString Then
            If MsgBox("Are your sure you'd like to overwrite this file:" & Chr(10) & strFileName, vbYesNo + vbQuestion, "Overwrite?") = vbNo Then
                strFileName = vbNullString
            End If
        End If
        
    Else
        InfoBox "Invalid file. Please choose another file name to save as..."
        strFileName = vbNullString
    End If
    
    ' Open the file dialog if we don't have a filepath
    Status "Get xml file path to save to..."   ' give our status
    If strFileName = vbNullString Then
        strFileName = Application.GetSaveAsFilename(, "(*.xml),*.xml", , "Save As...")
    End If
    
    If strFileName = "False" Then GoTo ExitProc
    
    If strFileName <> vbNullString Then
        'Get table data
        If IsNothing(rngSelection) Then
            varTable = Selection.Value
            varColumnHeaders = Selection.Rows(1).Value
        Else
            varTable = rngSelection.Value
            varColumnHeaders = rngSelection.Rows(1).Value
        End If
        
        'Build xml
        strXML = "<?xml version=""1.0"" encoding=""utf-8""?>"
        strXML = strXML & "<" & strTableElementName & ">"
        For intRow = 2 To UBound(varTable, 1)
            strXML = strXML & "<" & strRowElementName & ">"
            For intCol = 1 To UBound(varTable, 2)
                strXML = strXML & "<" & varColumnHeaders(1, intCol) & ">" & _
                    varTable(intRow, intCol) & "</" & varColumnHeaders(1, intCol) & ">"
            Next
            strXML = strXML & "</" & strRowElementName & ">"
        Next
        strXML = strXML & "</" & strTableElementName & ">"
        
        'Get next file number
        intFileNum = FreeFile
        
        'Open the file, write output, then close file
        Open strFileName For Output As #intFileNum
        Print #intFileNum, strXML
        Close #intFileNum
        
    ' invalid file path
    Else
        InfoBox "Invalid filepath", True
    End If
    
    '***********************************************************************************************
ExitProc:
    Exit Sub
ProcException:
    #If Not blnDeveloperMode Then
        InfoBox Err.Description & Chr(10) & "thrown from " & strModuleName & ": ExportRangeToXML", True
        Resume ExitProc
    #End If
End Sub

'***************************************************************************************************
' Author  : Robert Kreegier
' Purpose : Extracts from a path the filename or the highest level directory after the last slash
' Example : "C:\Users\UserName\Desktop\Test.txt" returns "Test.txt"
'           "C:\Users\UserName\Desktop" returns "Desktop"
'           ...however
'           "C:\Users\UserName\Desktop\" returns vbNullString
'***************************************************************************************************
Function ExtractFileName(ByVal strFilePathAndName As String) As String
    ExtractFileName = Right(strFilePathAndName, Len(strFilePathAndName) - InStrRev(strFilePathAndName, "\"))
End Function

'***************************************************************************************************
' Author  : Robert Kreegier
' Purpose : Extracts from a filename (with path or not) the extension of the file
' Example : "C:\Users\UserName\Desktop\Test.txt" returns "txt"
'           "C:\Users\UserName\Desktop\index.html" returns "html"
'           ...however
'           "C:\Users\UserName\Desktop\" returns vbNullString
'***************************************************************************************************
Function ExtractExtension(ByVal strFilePathAndName As String) As String
    ExtractExtension = vbNullString
    
    Dim intDotPos As Integer: intDotPos = InStrRev(strFilePathAndName, ".", , vbTextCompare)
    Dim intSlashPos As Integer: intSlashPos = InStrRev(strFilePathAndName, "\", , vbTextCompare)
    
    If intDotPos > intSlashPos Then
        ExtractExtension = Right(strFilePathAndName, Len(strFilePathAndName) - intDotPos)
    End If
End Function

'***************************************************************************************************
' Author  : Robert Kreegier
' Purpose : Extracts from a filename the path of the file.
' Example : "C:\Users\UserName\Desktop\Test.txt" returns "C:\Users\UserName\Desktop\"
'           "C:\Users\UserName\Desktop\index.html" returns "C:\Users\UserName\Desktop\"
'***************************************************************************************************
Function ExtractPath(ByVal strFilePathAndName As String) As String
    ExtractPath = Left(strFilePathAndName, InStrRev(strFilePathAndName, "\"))
End Function

'***************************************************************************************************
' Author  : Robert Kreegier
' Purpose : Verifies if a file exists in the specified path
'***************************************************************************************************
Function FileExists(ByVal strFilePathAndName As String) As Boolean
    FileExists = False
    
    Dim strFileName As String: strFileName = vbNullString
    Dim strDirTest As String: strDirTest = vbNullString
    
    On Error Resume Next
    strFileName = ExtractFileName(strFilePathAndName)
    strDirTest = Dir(strFilePathAndName)
    On Error GoTo 0
    
    If Not strDirTest = vbNullString And Not strFileName = vbNullString Then
        FileExists = True
    End If
End Function

'***************************************************************************************************
' Author  : GrahamSkan 2013-02-08
' Purpose : Removes "illegal" chars from a string so it can be used in a filename
' Source  : https://www.experts-exchange.com/questions/28025657/Vba-Code-Eliminate-Illegal-Characters-from-a-filename.html
'***************************************************************************************************
Function LegalizeFileName(ByVal strFileNameIn As String) As String
    Dim i As Integer
    
    Const strIllegals = "\/|?*<>"":"
    LegalizeFileName = strFileNameIn
    For i = 1 To Len(strIllegals)
        LegalizeFileName = Replace(LegalizeFileName, Mid$(strIllegals, i, 1), "_")
    Next i
End Function

'***************************************************************************************************
' Author  : Robert Kreegier
' Purpose : Writes the specified string data to a file
'***************************************************************************************************
Sub WriteStrToFile(ByVal strFileName As String, ByVal strData As String, Optional ByVal Overwrite As Boolean = False)
    If Overwrite And FileExists(strFileName) Then
        ' delete the file before moving on
        Kill strFileName
    End If

    Open strFileName For Append As #1

    Print #1, strData

    Close #1
End Sub

'***************************************************************************************************
' Author  : Justin Kay ("Jroonk") 8/15/2014
' Purpose : VBA Macro using late binding to copy text to clipboard.
' Source  : http://stackoverflow.com/questions/14219455/excel-vba-code-to-copy-a-specific-string-to-clipboard
'***************************************************************************************************
Sub CopyText(ByVal strText As String)
    Dim MSForms_DataObject As Object
    Set MSForms_DataObject = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    MSForms_DataObject.SetText strText
    MSForms_DataObject.PutInClipboard
    Set MSForms_DataObject = Nothing
End Sub

'***************************************************************************************************
' Author : Robert Kreegier
' Purpose: Determines if a color (.Interior.Color) can be subjectively interpreted as Grey
'***************************************************************************************************
Function IsGrey(ByRef dblColor As Double) As Boolean
    ' separate dblColor into RGB components so we can analyze their relationships
    Dim C As Long: C = dblColor
    Dim R As Long: R = C Mod 256
    Dim G As Long: G = C \ 256 Mod 256
    Dim B As Long: B = C \ 65536 Mod 256
    
    ' preset our return value
    IsGrey = False
    
    ' The following values have been subjectively judged via trial and error and have no empirical basis.
    ' First we see if the values are either too dark or too light. If either R, G, or B are below or
    ' above these values, then we can safely assume that the resulting color is not going to be grey.
    If R > 65 And G > 65 And B > 65 And R < 230 And G < 230 And B < 230 Then
        ' Next, we look at the ratios between R, G, and B. If the ratio is too high, or rather if the
        ' channels are too different from each other, then we can say that the resulting color won't
        ' be grey.
        If (R / G) > 0.75 And (R / G) < 1.25 Then
            If (R / B) > 0.75 And (R / B) < 1.25 Then
                If (G / B) > 0.75 And (G / B) < 1.25 Then
                    IsGrey = True
                End If
            End If
        End If
    End If
End Function

'***************************************************************************************************
' Author : Robert Kreegier
' Purpose: Determines if a color (.Interior.Color) can be subjectively interpreted as Green
'***************************************************************************************************
Function IsGreen(ByRef dblColor As Double) As Boolean
    ' separate dblColor into RGB components so we can analyze their relationships
    Dim C As Long: C = dblColor
    Dim R As Long: R = C Mod 256
    Dim G As Long: G = C \ 256 Mod 256
    Dim B As Long: B = C \ 65536 Mod 256
    
    ' preset our return value
    IsGreen = False
    
    ' The following values have been subjectively judged via trial and error and have no empirical basis.
    ' If the green channel is lower than 65, we judge it to be too dark.
    If G > 65 Then
        ' If the ratio of red and blue to green is lower than .92 and .73, respectively, then we can say this color is green.
        If (R / G) < 0.92 And (B / G) < 0.73 Then
            IsGreen = True
        End If
    End If
End Function

'***************************************************************************************************
' Author : Johannes, Robert Kreegier
' Purpose: Merges two arrays that were transposed ranges
' Source : http://stackoverflow.com/questions/1588913/how-do-i-merge-two-arrays-in-vba
'***************************************************************************************************
Function MergeTransRanges(ByVal Arr1 As Variant, ByVal Arr2 As Variant) As Variant
    Dim returnThis() As Variant
    Dim len1 As Integer: len1 = UBound(Arr1)
    Dim len2 As Integer: len2 = UBound(Arr2)
    Dim lenRe As Integer: lenRe = len1 + len2 + 1
    
    ReDim returnThis(LBound(Arr1) To lenRe)
    
    Dim counter As Integer
    For counter = LBound(Arr1) To UBound(Arr1)
        returnThis(counter) = Arr1(counter)
    Next counter
    
    For counter = counter To lenRe - LBound(Arr1)
        returnThis(counter) = Arr2(counter - len1)
    Next counter

    MergeTransRanges = returnThis
End Function

'***************************************************************************************************
' Author : Robert M Kreegier
' Purpose: Counts the number of unique values in a range
'***************************************************************************************************
Function CountUnique(ByRef rngTarget As Range) As Long
    CountUnique = UniqueDict(rngTarget).Count
End Function

'***************************************************************************************************
' Author : Robert M Kreegier
' Purpose: Returns a dictionary object with all the unique values in a range
'***************************************************************************************************
Function UniqueDict(ByRef rngTarget As Range) As Object
    Dim vntArray() As Variant
    vntArray = rngTarget

    Dim dctUnique As Object: Set dctUnique = CreateObject("Scripting.Dictionary")
    Dim strPrevious As String: strPrevious = vbNullString
    Dim lngRowIndex As Long
    Dim lngColIndex As Long
    Dim lngKey As Long: lngKey = 0

    For lngColIndex = 1 To UBound(vntArray, 2)
        For lngRowIndex = 1 To UBound(vntArray, 1)
            If vntArray(lngRowIndex, lngColIndex) <> strPrevious And vntArray(lngRowIndex, lngColIndex) <> vbNullString Then
                If Not dctUnique.Exists(vntArray(lngRowIndex, lngColIndex)) Then
                    dctUnique.Add vntArray(lngRowIndex, lngColIndex), lngKey
                    lngKey = lngKey + 1
                    strPrevious = vntArray(lngRowIndex, lngColIndex)
                End If
            End If
        Next lngRowIndex
    Next lngColIndex

    Set UniqueDict = dctUnique
End Function

'***************************************************************************************************
' Author  : Chip Pearson and Pearson Software Consulting, LLC
' Purpose : This sorts a Dictionary object. If SortByKey is False, the
'           the sort is done based on the Items of the Dictionary, and
'           these items must be simple data types. They may not be
'           Object, Arrays, or User-Defined Types. If SortByKey is True,
'           the Dictionary is sorted by Key value, and the Items in the
'           Dictionary may be Object as well as simple variables.
'
'           If sort by key is True, all element of the Dictionary
'           must have a non-blank Key value. If Key is vbNullString
'           the procedure will terminate.
'
'           By defualt, sorting is done in Ascending order. You can
'           sort by Descending order by setting the Descending parameter
'           to True.
'
'           By default, text comparisons are done case-INSENSITIVE (e.g.,
'           "a" = "A"). To use case-SENSITIVE comparisons (e.g., "a" <> "A")
'           set CompareMode to vbBinaryCompare.
'
' Note    : This procedure requires the
'           QSortInPlace function, which is described and available for
'           download at www.cpearson.com/excel/qsort.htm
' Source  : http://www.cpearson.com/excel/CollectionsAndDictionaries.htm
'***************************************************************************************************
Public Sub SortDictionary(Dict As Object, _
                          SortByKey As Boolean, _
                          Optional Descending As Boolean = False, _
                          Optional CompareMode As VbCompareMethod = vbTextCompare)
    Dim Ndx As Long
    Dim KeyValue As String
    Dim ItemValue As Variant
    Dim Arr() As Variant
    Dim KeyArr() As String
    Dim VTypes() As VbVarType
    
    
    Dim V As Variant
    Dim SplitArr As Variant
    
    Dim TempDict As Object
    '''''''''''''''''''''''''''''
    ' Ensure Dict is not Nothing.
    '''''''''''''''''''''''''''''
    If Dict Is Nothing Then
        Exit Sub
    End If
    ''''''''''''''''''''''''''''
    ' If the number of elements
    ' in Dict is 0 or 1, no
    ' sorting is required.
    ''''''''''''''''''''''''''''
    If (Dict.Count = 0) Or (Dict.Count = 1) Then
        Exit Sub
    End If
    
    ''''''''''''''''''''''''''''
    ' Create a new TempDict.
    ''''''''''''''''''''''''''''
    Set TempDict = CreateObject("Scripting.Dictionary") 'New Scripting.Dictionary
    
    If SortByKey = True Then
        ''''''''''''''''''''''''''''''''''''''''
        ' We're sorting by key. Redim the Arr
        ' to the number of elements in the
        ' Dict object, and load that array
        ' with the key names.
        ''''''''''''''''''''''''''''''''''''''''
        ReDim Arr(0 To Dict.Count - 1)
        
        For Ndx = 0 To Dict.Count - 1
            Arr(Ndx) = Dict.Keys()(Ndx)
        Next Ndx
        
        ''''''''''''''''''''''''''''''''''''''
        ' Sort the key names.
        ''''''''''''''''''''''''''''''''''''''
        QSortInPlace InputArray:=Arr, LB:=-1, UB:=-1, Descending:=Descending, CompareMode:=CompareMode
        ''''''''''''''''''''''''''''''''''''''''''''
        ' Load TempDict. The key value come from
        ' our sorted array of keys Arr, and the
        ' Item comes from the original Dict object.
        ''''''''''''''''''''''''''''''''''''''''''''
        For Ndx = 0 To Dict.Count - 1
            KeyValue = Arr(Ndx)
            TempDict.Add Key:=KeyValue, Item:=Dict.Item(KeyValue)
        Next Ndx
        '''''''''''''''''''''''''''''''''
        ' Set the passed in Dict object
        ' to our TempDict object.
        '''''''''''''''''''''''''''''''''
        Set Dict = TempDict
        ''''''''''''''''''''''''''''''''
        ' This is the end of processing.
        ''''''''''''''''''''''''''''''''
    Else
        '''''''''''''''''''''''''''''''''''''''''''''''
        ' Here, we're sorting by items. The Items must
        ' be simple data types. They may NOT be Objects,
        ' arrays, or UserDefineTypes.
        ' First, ReDim Arr and VTypes to the number
        ' of elements in the Dict object. Arr will
        ' hold a string containing
        '   Item & vbNullChar & Key
        ' This keeps the association between the
        ' item and its key.
        '''''''''''''''''''''''''''''''''''''''''''''''
        ReDim Arr(0 To Dict.Count - 1)
        ReDim VTypes(0 To Dict.Count - 1)
    
        For Ndx = 0 To Dict.Count - 1
            If (IsObject(Dict.Items(Ndx)) = True) Or _
                (IsArray(Dict.Items(Ndx)) = True) Or _
                VarType(Dict.Items(Ndx)) = vbUserDefinedType Then
                Debug.Print "***** ITEM IN DICTIONARY WAS OBJECT OR ARRAY OR UDT"
                Exit Sub
            End If
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' Here, we create a string containing
            '       Item & vbNullChar & Key
            ' This preserves the associate between an item and its
            ' key. Store the VarType of the Item in the VTypes
            ' array. We'll use these values later to convert
            ' back to the proper data type for Item.
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                Arr(Ndx) = Dict.Items(Ndx) & vbNullChar & Dict.Keys(Ndx)
                VTypes(Ndx) = VarType(Dict.Items(Ndx))
                
        Next Ndx
        ''''''''''''''''''''''''''''''''''
        ' Sort the array that contains the
        ' items of the Dictionary along
        ' with their associated keys
        ''''''''''''''''''''''''''''''''''
        QSortInPlace InputArray:=Arr, LB:=-1, UB:=-1, Descending:=Descending, CompareMode:=vbTextCompare
        
        For Ndx = LBound(Arr) To UBound(Arr)
            '''''''''''''''''''''''''''''''''''''
            ' Loop trhogh the array of sorted
            ' Items, Split based on vbNullChar
            ' to get the Key from the element
            ' of the array Arr.
            SplitArr = Split(Arr(Ndx), vbNullChar)
            ''''''''''''''''''''''''''''''''''''''''''
            ' It may have been possible that item in
            ' the dictionary contains a vbNullChar.
            ' Therefore, use UBound to get the
            ' key value, which will necessarily
            ' be the last item of SplitArr.
            ' Then Redim Preserve SplitArr
            ' to UBound - 1 to get rid of the
            ' Key element, and use Join
            ' to reassemble to original value
            ' of the Item.
            '''''''''''''''''''''''''''''''''''''''''
            KeyValue = SplitArr(UBound(SplitArr))
            ReDim Preserve SplitArr(LBound(SplitArr) To UBound(SplitArr) - 1)
            ItemValue = Join(SplitArr, vbNullChar)
            '''''''''''''''''''''''''''''''''''''''
            ' Join will set ItemValue to a string
            ' regardless of what the original
            ' data type was. Test the VTypes(Ndx)
            ' value to convert ItemValue back to
            ' the proper data type.
            '''''''''''''''''''''''''''''''''''''''
            Select Case VTypes(Ndx)
                Case vbBoolean
                    ItemValue = CBool(ItemValue)
                Case vbByte
                    ItemValue = CByte(ItemValue)
                Case vbCurrency
                    ItemValue = CCur(ItemValue)
                Case vbDate
                    ItemValue = CDate(ItemValue)
                Case vbDecimal
                    ItemValue = CDec(ItemValue)
                Case vbDouble
                    ItemValue = CDbl(ItemValue)
                Case vbInteger
                    ItemValue = CInt(ItemValue)
                Case vbLong
                    ItemValue = CLng(ItemValue)
                Case vbSingle
                    ItemValue = CSng(ItemValue)
                Case vbString
                    ItemValue = CStr(ItemValue)
                Case Else
                    ItemValue = ItemValue
            End Select
            ''''''''''''''''''''''''''''''''''''''
            ' Finally, add the Item and Key to
            ' our TempDict dictionary.
            
            TempDict.Add Key:=KeyValue, Item:=ItemValue
        Next Ndx
    End If
    
    
    '''''''''''''''''''''''''''''''''
    ' Set the passed in Dict object
    ' to our TempDict object.
    '''''''''''''''''''''''''''''''''
    Set Dict = TempDict
End Sub

'***************************************************************************************************
' Author  : Chip Pearson at Pearson Software Consulting
' Purpose : This function sorts the array InputArray in place -- this is, the original array in the
'           calling procedure is sorted. It will work with either string data or numeric data.
'           It need not sort the entire array. You can sort only part of the array by setting the LB and
'           UB parameters to the first (LB) and last (UB) element indexes that you want to sort.
'           LB and UB are optional parameters. If omitted LB is set to the LBound of InputArray, and if
'           omitted UB is set to the UBound of the InputArray. If you want to sort the entire array,
'           omit the LB and UB parameters, or set both to -1, or set LB = LBound(InputArray) and set
'           UB to UBound(InputArray).
'
'           By default, the sort method is case INSENSTIVE (case doens't matter: "A", "b", "C", "d").
'           To make it case SENSITIVE (case matters: "A" "C" "b" "d"), set the CompareMode argument
'           to vbBinaryCompare (=0). If Compare mode is omitted or is any value other than vbBinaryCompare,
'           it is assumed to be vbTextCompare and the sorting is done case INSENSITIVE.
'
'           The function returns TRUE if the array was successfully sorted or FALSE if an error
'           occurred. If an error occurs (e.g., LB > UB), a message box indicating the error is
'           displayed. To suppress message boxes, set the NoAlerts parameter to TRUE.
'
''''''''''''''''''''''''''''''''''''''
' MODIFYING THIS CODE:
''''''''''''''''''''''''''''''''''''''
' If you modify this code and you call "Exit Procedure", you MUST decrment the RecursionLevel
' variable. E.g.,
'       If SomethingThatCausesAnExit Then
'           RecursionLevel = RecursionLevel - 1
'           Exit Function
'       End If
'''''''''''''''''''''''''''''''''''''''
'
' Note    : If you coerce InputArray to a ByVal argument, QSortInPlace will not be
'           able to reference the InputArray in the calling procedure and the array will
'           not be sorted.
'
' This function uses the following procedures. These are declared as Private procedures:
'       IsArrayAllocated
'       IsSimpleDataType
'       IsSimpleNumericType
'       QSortCompare
'       NumberOfArrayDimensions
'       ReverseArrayInPlace
' Source  : www.cpearson.com/excel/SortingArrays.aspx
'***************************************************************************************************
Public Function QSortInPlace( _
    ByRef InputArray As Variant, _
    Optional ByVal LB As Long = -1&, _
    Optional ByVal UB As Long = -1&, _
    Optional ByVal Descending As Boolean = False, _
    Optional ByVal CompareMode As VbCompareMethod = vbTextCompare, _
    Optional ByVal NoAlerts As Boolean = False) As Boolean
    Dim Temp As Variant
    Dim Buffer As Variant
    Dim CurLow As Long
    Dim CurHigh As Long
    Dim CurMidpoint As Long
    Dim Ndx As Long
    Dim pCompareMode As VbCompareMethod
    
    '''''''''''''''''''''''''
    ' Set the default result.
    '''''''''''''''''''''''''
    QSortInPlace = False
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' This variable is used to determine the level
    ' of recursion  (the function calling itself).
    ' RecursionLevel is incremented when this procedure
    ' is called, either initially by a calling procedure
    ' or recursively by itself. The variable is decremented
    ' when the procedure exits. We do the input parameter
    ' validation only when RecursionLevel is 1 (when
    ' the function is called by another function, not
    ' when it is called recursively).
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Static RecursionLevel As Long
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Keep track of the recursion level -- that is, how many
    ' times the procedure has called itself.
    ' Carry out the validation routines only when this
    ' procedure is first called. Don't run the
    ' validations on a recursive call to the
    ' procedure.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    RecursionLevel = RecursionLevel + 1
    
    If RecursionLevel = 1 Then
        ''''''''''''''''''''''''''''''''''
        ' Ensure InputArray is an array.
        ''''''''''''''''''''''''''''''''''
        If IsArray(InputArray) = False Then
            If NoAlerts = False Then
                MsgBox "The InputArray parameter is not an array."
            End If
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' InputArray is not an array. Exit with a False result.
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''
            RecursionLevel = RecursionLevel - 1
            Exit Function
        End If
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Test LB and UB. If < 0 then set to LBound and UBound
        ' of the InputArray.
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If LB < 0 Then
            LB = LBound(InputArray)
        End If
        If UB < 0 Then
            UB = UBound(InputArray)
        End If
        
        Select Case NumberOfArrayDimensions(InputArray)
            Case 0
                ''''''''''''''''''''''''''''''''''''''''''
                ' Zero dimensions indicates an unallocated
                ' dynamic array.
                ''''''''''''''''''''''''''''''''''''''''''
                If NoAlerts = False Then
                    MsgBox "The InputArray is an empty, unallocated array."
                End If
                RecursionLevel = RecursionLevel - 1
                Exit Function
            Case 1
                ''''''''''''''''''''''''''''''''''''''''''
                ' We sort ONLY single dimensional arrays.
                ''''''''''''''''''''''''''''''''''''''''''
            Case Else
                ''''''''''''''''''''''''''''''''''''''''''
                ' We sort ONLY single dimensional arrays.
                ''''''''''''''''''''''''''''''''''''''''''
                If NoAlerts = False Then
                    MsgBox "The InputArray is multi-dimensional." & _
                          "QSortInPlace works only on single-dimensional arrays."
                End If
                RecursionLevel = RecursionLevel - 1
                Exit Function
        End Select
        '''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Ensure that InputArray is an array of simple data
        ' types, not other arrays or objects. This tests
        ' the data type of only the first element of
        ' InputArray. If InputArray is an array of Variants,
        ' subsequent data types may not be simple data types
        ' (e.g., they may be objects or other arrays), and
        ' this may cause QSortInPlace to fail on the StrComp
        ' operation.
        '''''''''''''''''''''''''''''''''''''''''''''''''''
        If IsSimpleDataType(InputArray(LBound(InputArray))) = False Then
            If NoAlerts = False Then
                MsgBox "InputArray is not an array of simple data types."
                RecursionLevel = RecursionLevel - 1
                Exit Function
            End If
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' ensure that the LB parameter is valid.
        ''''''''''''''''''''''''''''''''''''''''''''''''''''
        Select Case LB
            Case Is < LBound(InputArray)
                If NoAlerts = False Then
                    MsgBox "The LB lower bound parameter is less than the LBound of the InputArray"
                End If
                RecursionLevel = RecursionLevel - 1
                Exit Function
            Case Is > UBound(InputArray)
                If NoAlerts = False Then
                    MsgBox "The LB lower bound parameter is greater than the UBound of the InputArray"
                End If
                RecursionLevel = RecursionLevel - 1
                Exit Function
            Case Is > UB
                If NoAlerts = False Then
                    MsgBox "The LB lower bound parameter is greater than the UB upper bound parameter."
                End If
                RecursionLevel = RecursionLevel - 1
                Exit Function
        End Select
    
        ''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' ensure the UB parameter is valid.
        ''''''''''''''''''''''''''''''''''''''''''''''''''''
        Select Case UB
            Case Is > UBound(InputArray)
                If NoAlerts = False Then
                    MsgBox "The UB upper bound parameter is greater than the upper bound of the InputArray."
                End If
                RecursionLevel = RecursionLevel - 1
                Exit Function
            Case Is < LBound(InputArray)
                If NoAlerts = False Then
                    MsgBox "The UB upper bound parameter is less than the lower bound of the InputArray."
                End If
                RecursionLevel = RecursionLevel - 1
                Exit Function
            Case Is < LB
                If NoAlerts = False Then
                    MsgBox "the UB upper bound parameter is less than the LB lower bound parameter."
                End If
                RecursionLevel = RecursionLevel - 1
                Exit Function
        End Select
    
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' if UB = LB, we have nothing to sort, so get out.
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If UB = LB Then
            QSortInPlace = True
            RecursionLevel = RecursionLevel - 1
            Exit Function
        End If
    
    End If ' RecursionLevel = 1
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Ensure that CompareMode is either vbBinaryCompare  or
    ' vbTextCompare. If it is neither, default to vbTextCompare.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If (CompareMode = vbBinaryCompare) Or (CompareMode = vbTextCompare) Then
        pCompareMode = CompareMode
    Else
        pCompareMode = vbTextCompare
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Begin the actual sorting process.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    CurLow = LB
    CurHigh = UB
    
    If LB = 0 Then
        CurMidpoint = ((LB + UB) \ 2) + 1
    Else
        CurMidpoint = (LB + UB) \ 2 ' note integer division (\) here
    End If
    Temp = InputArray(CurMidpoint)
    
    Do While (CurLow <= CurHigh)
        
        Do While QSortCompare(V1:=InputArray(CurLow), V2:=Temp, CompareMode:=pCompareMode) < 0
            CurLow = CurLow + 1
            If CurLow = UB Then
                Exit Do
            End If
        Loop
        
        Do While QSortCompare(V1:=Temp, V2:=InputArray(CurHigh), CompareMode:=pCompareMode) < 0
            CurHigh = CurHigh - 1
            If CurHigh = LB Then
               Exit Do
            End If
        Loop
    
        If (CurLow <= CurHigh) Then
            Buffer = InputArray(CurLow)
            InputArray(CurLow) = InputArray(CurHigh)
            InputArray(CurHigh) = Buffer
            CurLow = CurLow + 1
            CurHigh = CurHigh - 1
        End If
    Loop
    
    If LB < CurHigh Then
        QSortInPlace InputArray:=InputArray, LB:=LB, UB:=CurHigh, _
            Descending:=Descending, CompareMode:=pCompareMode, NoAlerts:=True
    End If
    
    If CurLow < UB Then
        QSortInPlace InputArray:=InputArray, LB:=CurLow, UB:=UB, _
            Descending:=Descending, CompareMode:=pCompareMode, NoAlerts:=True
    End If
    
    '''''''''''''''''''''''''''''''''''''
    ' If Descending is True, reverse the
    ' order of the array, but only if the
    ' recursion level is 1.
    '''''''''''''''''''''''''''''''''''''
    If Descending = True Then
        If RecursionLevel = 1 Then
            ReverseArrayInPlace2 InputArray, LB, UB
        End If
    End If
    
    RecursionLevel = RecursionLevel - 1
    QSortInPlace = True
End Function

'***************************************************************************************************
' Author  : Chip Pearson at Pearson Software Consulting
' Purpose : This function is used in QSortInPlace to compare two elements. If
'           V1 AND V2 are both numeric data types (integer, long, single, double)
'           they are converted to Doubles and compared. If V1 and V2 are BOTH strings
'           that contain numeric data, they are converted to Doubles and compared.
'           If either V1 or V2 is a string and does NOT contain numeric data, both
'           V1 and V2 are converted to Strings and compared with StrComp.
'
'           The result is -1 if V1 < V2,
'                          0 if V1 = V2
'                          1 if V1 > V2
'           For text comparisons, case sensitivity is controlled by CompareMode.
'           If this is vbBinaryCompare, the result is case SENSITIVE. If this
'           is omitted or any other value, the result is case INSENSITIVE.
' Source  : www.cpearson.com/excel/SortingArrays.aspx
'***************************************************************************************************
Public Function QSortCompare(V1 As Variant, V2 As Variant, _
                             Optional CompareMode As VbCompareMethod = vbTextCompare) As Long
    Dim D1 As Double
    Dim D2 As Double
    Dim S1 As String
    Dim S2 As String
    
    Dim Compare As VbCompareMethod
    ''''''''''''''''''''''''''''''''''''''''''''''''
    ' Test CompareMode. Any value other than
    ' vbBinaryCompare will default to vbTextCompare.
    ''''''''''''''''''''''''''''''''''''''''''''''''
    If CompareMode = vbBinaryCompare Or CompareMode = vbTextCompare Then
        Compare = CompareMode
    Else
        Compare = vbTextCompare
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''
    ' If either V1 or V2 is either an array or
    ' an Object, raise a error 13 - Type Mismatch.
    '''''''''''''''''''''''''''''''''''''''''''''''
    If IsArray(V1) = True Or IsArray(V2) = True Then
        Err.Raise 13
        Exit Function
    End If
    If IsObject(V1) = True Or IsObject(V2) = True Then
        Err.Raise 13
        Exit Function
    End If
    
    If IsSimpleNumericType(V1) = True Then
        If IsSimpleNumericType(V2) = True Then
            '''''''''''''''''''''''''''''''''''''
            ' If BOTH V1 and V2 are numeric data
            ' types, then convert to Doubles and
            ' do an arithmetic compare and
            ' return the result.
            '''''''''''''''''''''''''''''''''''''
            D1 = CDbl(V1)
            D2 = CDbl(V2)
            If D1 = D2 Then
                QSortCompare = 0
                Exit Function
            End If
            If D1 < D2 Then
                QSortCompare = -1
                Exit Function
            End If
            If D1 > D2 Then
                QSortCompare = 1
                Exit Function
            End If
        End If
    End If
    ''''''''''''''''''''''''''''''''''''''''''''
    ' Either V1 or V2 was not numeric data type.
    ' Test whether BOTH V1 AND V2 are numeric
    ' strings. If BOTH are numeric, convert to
    ' Doubles and do a arithmetic comparison.
    ''''''''''''''''''''''''''''''''''''''''''''
    If IsNumeric(V1) = True And IsNumeric(V2) = True Then
        D1 = CDbl(V1)
        D2 = CDbl(V2)
        If D1 = D2 Then
            QSortCompare = 0
            Exit Function
        End If
        If D1 < D2 Then
            QSortCompare = -1
            Exit Function
        End If
        If D1 > D2 Then
            QSortCompare = 1
            Exit Function
        End If
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''
    ' Either or both V1 and V2 was not numeric
    ' string. In this case, convert to Strings
    ' and use StrComp to compare.
    ''''''''''''''''''''''''''''''''''''''''''''''
    S1 = CStr(V1)
    S2 = CStr(V2)
    QSortCompare = StrComp(S1, S2, Compare)
End Function

'***************************************************************************************************
' Author  : Chip Pearson at Pearson Software Consulting
' Purpose : This function returns the number of dimensions of an array. An unallocated dynamic array
'           has 0 dimensions. This condition can also be tested with IsArrayEmpty.
' Source  : www.cpearson.com/excel/SortingArrays.aspx
'***************************************************************************************************
Public Function NumberOfArrayDimensions(Arr As Variant) As Integer
    Dim Ndx As Integer
    Dim Res As Integer
    On Error Resume Next
    ' Loop, increasing the dimension index Ndx, until an error occurs.
    ' An error will occur when Ndx exceeds the number of dimension
    ' in the array. Return Ndx - 1.
    Do
        Ndx = Ndx + 1
        Res = UBound(Arr, Ndx)
    Loop Until Err.number <> 0
    
    NumberOfArrayDimensions = Ndx - 1
End Function

'***************************************************************************************************
' Author  : Chip Pearson at Pearson Software Consulting
' Purpose : This procedure reverses the order of an array in place -- this is, the array variable
'           in the calling procedure is sorted. An error will occur if InputArray is not an array,
'           if it is an empty, unallocated array, or if the number of dimensions is not 1.
'
' Note    : Before calling the ReverseArrayInPlace procedure, consider if your needs can
'           be met by simply reading the existing array in reverse order (Step -1). If so, you can save
'           the overhead added to your application by calling this function.
'
'           The function returns TRUE if the array was successfully reversed, or FALSE if
'           an error occurred.
'
'           If an error occurred, a message box is displayed indicating the error. To suppress
'           the message box and simply return FALSE, set the NoAlerts parameter to TRUE.
' Source  : www.cpearson.com/excel/SortingArrays.aspx
'***************************************************************************************************
Public Function ReverseArrayInPlace(InputArray As Variant, _
                                    Optional NoAlerts As Boolean = False) As Boolean
    Dim Temp As Variant
    Dim Ndx As Long
    Dim Ndx2 As Long
    Dim OrigN As Long
    Dim NewN As Long
    Dim NewArr() As Variant
    
    ''''''''''''''''''''''''''''''''
    ' Set the default return value.
    ''''''''''''''''''''''''''''''''
    ReverseArrayInPlace = False
    
    '''''''''''''''''''''''''''''''''
    ' Ensure we have an array
    '''''''''''''''''''''''''''''''''
    If IsArray(InputArray) = False Then
       If NoAlerts = False Then
            MsgBox "The InputArray parameter is not an array."
        End If
        Exit Function
    End If
    
    ''''''''''''''''''''''''''''''''''''''
    ' Test the number of dimensions of the
    ' InputArray. If 0, we have an empty,
    ' unallocated array. Get out with
    ' an error message. If greater than
    ' one, we have a multi-dimensional
    ' array, which is not allowed. Only
    ' an allocated 1-dimensional array is
    ' allowed.
    ''''''''''''''''''''''''''''''''''''''
    Select Case NumberOfArrayDimensions(InputArray)
        Case 0
            '''''''''''''''''''''''''''''''''''''''''''
            ' Zero dimensions indicates an unallocated
            ' dynamic array.
            '''''''''''''''''''''''''''''''''''''''''''
            If NoAlerts = False Then
                MsgBox "The input array is an empty, unallocated array."
            End If
            Exit Function
        Case 1
            '''''''''''''''''''''''''''''''''''''''''''
            ' We can reverse ONLY a single dimensional
            ' arrray.
            '''''''''''''''''''''''''''''''''''''''''''
        Case Else
            '''''''''''''''''''''''''''''''''''''''''''
            ' We can reverse ONLY a single dimensional
            ' arrray.
            '''''''''''''''''''''''''''''''''''''''''''
            If NoAlerts = False Then
                MsgBox "The input array multi-dimensional. ReverseArrayInPlace works only " & _
                       "on single-dimensional arrays."
            End If
            Exit Function
    
    End Select
    
    '''''''''''''''''''''''''''''''''''''''''''''
    ' Ensure that we have only simple data types,
    ' not an array of objects or arrays.
    '''''''''''''''''''''''''''''''''''''''''''''
    If IsSimpleDataType(InputArray(LBound(InputArray))) = False Then
        If NoAlerts = False Then
            MsgBox "The input array contains arrays, objects, or other complex data types." & vbCrLf & _
                "ReverseArrayInPlace can reverse only arrays of simple data types."
            Exit Function
        End If
    End If
    
    ReDim NewArr(LBound(InputArray) To UBound(InputArray))
    NewN = UBound(NewArr)
    For OrigN = LBound(InputArray) To UBound(InputArray)
        NewArr(NewN) = InputArray(OrigN)
        NewN = NewN - 1
    Next OrigN
    
    For NewN = LBound(NewArr) To UBound(NewArr)
        InputArray(NewN) = NewArr(NewN)
    Next NewN
    
    ReverseArrayInPlace = True
End Function

'***************************************************************************************************
' Author  : Chip Pearson at Pearson Software Consulting
' Purpose : This reverses the order of elements in InputArray. To reverse the entire array, omit or
'           set to less than 0 the LB and UB parameters. To reverse only part of tbe array, set LB and/or
'           UB to the LBound and UBound of the sub array to be reversed.
' Source  : www.cpearson.com/excel/SortingArrays.aspx
'***************************************************************************************************
Public Function ReverseArrayInPlace2(InputArray As Variant, _
    Optional LB As Long = -1, Optional UB As Long = -1, _
    Optional NoAlerts As Boolean = False) As Boolean
    Dim N As Long
    Dim Temp As Variant
    Dim Ndx As Long
    Dim Ndx2 As Long
    Dim OrigN As Long
    Dim NewN As Long
    Dim NewArr() As Variant
    
    ''''''''''''''''''''''''''''''''
    ' Set the default return value.
    ''''''''''''''''''''''''''''''''
    ReverseArrayInPlace2 = False
    
    '''''''''''''''''''''''''''''''''
    ' Ensure we have an array
    '''''''''''''''''''''''''''''''''
    If IsArray(InputArray) = False Then
        If NoAlerts = False Then
            MsgBox "The InputArray parameter is not an array."
        End If
        Exit Function
    End If
    
    ''''''''''''''''''''''''''''''''''''''
    ' Test the number of dimensions of the
    ' InputArray. If 0, we have an empty,
    ' unallocated array. Get out with
    ' an error message. If greater than
    ' one, we have a multi-dimensional
    ' array, which is not allowed. Only
    ' an allocated 1-dimensional array is
    ' allowed.
    ''''''''''''''''''''''''''''''''''''''
    Select Case NumberOfArrayDimensions(InputArray)
        Case 0
            '''''''''''''''''''''''''''''''''''''''''''
            ' Zero dimensions indicates an unallocated
            ' dynamic array.
            '''''''''''''''''''''''''''''''''''''''''''
            If NoAlerts = False Then
                MsgBox "The input array is an empty, unallocated array."
            End If
            Exit Function
        Case 1
            '''''''''''''''''''''''''''''''''''''''''''
            ' We can reverse ONLY a single dimensional
            ' arrray.
            '''''''''''''''''''''''''''''''''''''''''''
        Case Else
            '''''''''''''''''''''''''''''''''''''''''''
            ' We can reverse ONLY a single dimensional
            ' arrray.
            '''''''''''''''''''''''''''''''''''''''''''
            If NoAlerts = False Then
                MsgBox "The input array multi-dimensional. ReverseArrayInPlace works only " & _
                       "on single-dimensional arrays."
            End If
            Exit Function
    
    End Select
    
    '''''''''''''''''''''''''''''''''''''''''''''
    ' Ensure that we have only simple data types,
    ' not an array of objects or arrays.
    '''''''''''''''''''''''''''''''''''''''''''''
    If IsSimpleDataType(InputArray(LBound(InputArray))) = False Then
        If NoAlerts = False Then
            MsgBox "The input array contains arrays, objects, or other complex data types." & vbCrLf & _
                "ReverseArrayInPlace can reverse only arrays of simple data types."
            Exit Function
        End If
    End If
    
    If LB < 0 Then
        LB = LBound(InputArray)
    End If
    If UB < 0 Then
        UB = UBound(InputArray)
    End If
    
    For N = LB To (LB + ((UB - LB - 1) \ 2))
        Temp = InputArray(N)
        InputArray(N) = InputArray(UB - (N - LB))
        InputArray(UB - (N - LB)) = Temp
    Next N
    
    ReverseArrayInPlace2 = True
End Function

'***************************************************************************************************
' Author  : Chip Pearson at Pearson Software Consulting
' Purpose : This returns TRUE if V is one of the following data types:
'               vbBoolean
'               vbByte
'               vbCurrency
'               vbDate
'               vbDecimal
'               vbDouble
'               vbInteger
'               vbLong
'               vbSingle
'               vbVariant if it contains a numeric value
'           It returns FALSE for any other data type, including any array
'           or vbEmpty.
' Source  : www.cpearson.com/excel/SortingArrays.aspx
'***************************************************************************************************
Public Function IsSimpleNumericType(V As Variant) As Boolean
    If IsSimpleDataType(V) = True Then
        Select Case VarType(V)
            Case vbBoolean, _
                    vbByte, _
                    vbCurrency, _
                    vbDate, _
                    vbDecimal, _
                    vbDouble, _
                    vbInteger, _
                    vbLong, _
                    vbSingle
                IsSimpleNumericType = True
            Case vbVariant
                If IsNumeric(V) = True Then
                    IsSimpleNumericType = True
                Else
                    IsSimpleNumericType = False
                End If
            Case Else
                IsSimpleNumericType = False
        End Select
    Else
        IsSimpleNumericType = False
    End If
End Function

'***************************************************************************************************
' Author  : Chip Pearson at Pearson Software Consulting
' Purpose : This function returns TRUE if V is one of the following
'           variable types (as returned by the VarType function:
'               vbBoolean
'               vbByte
'               vbCurrency
'               vbDate
'               vbDecimal
'               vbDouble
'               vbEmpty
'               vbError
'               vbInteger
'               vbLong
'               vbNull
'               vbSingle
'               vbString
'               vbVariant
'
'           It returns FALSE if V is any one of the following variable
'           types:
'               vbArray
'               vbDataObject
'               vbObject
'               vbUserDefinedType
'               or if it is an array of any type.
' Source  : www.cpearson.com/excel/SortingArrays.aspx
'***************************************************************************************************
Public Function IsSimpleDataType(V As Variant) As Boolean
    On Error Resume Next
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Test if V is an array. We can't just use VarType(V) = vbArray
    ' because the VarType of an array is vbArray + VarType(type
    ' of array element). E.g, the VarType of an Array of Longs is
    ' 8195 = vbArray + vbLong.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If IsArray(V) = True Then
        IsSimpleDataType = False
        Exit Function
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' We must also explicitly check whether V is an object, rather
    ' relying on VarType(V) to equal vbObject. The reason is that
    ' if V is an object and that object has a default proprety, VarType
    ' returns the data type of the default property. For example, if
    ' V is an Excel.Range object pointing to cell A1, and A1 contains
    ' 12345, VarType(V) would return vbDouble, the since Value is
    ' the default property of an Excel.Range object and the default
    ' numeric type of Value in Excel is Double. Thus, in order to
    ' prevent this type of behavior with default properties, we test
    ' IsObject(V) to see if V is an object.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If IsObject(V) = True Then
        IsSimpleDataType = False
        Exit Function
    End If
    '''''''''''''''''''''''''''''''''''''
    ' Test the value returned by VarType.
    '''''''''''''''''''''''''''''''''''''
    Select Case VarType(V)
        Case vbArray, vbDataObject, vbObject, vbUserDefinedType
            '''''''''''''''''''''''
            ' not simple data types
            '''''''''''''''''''''''
            IsSimpleDataType = False
        Case Else
            ''''''''''''''''''''''''''''''''''''
            ' otherwise it is a simple data type
            ''''''''''''''''''''''''''''''''''''
            IsSimpleDataType = True
    End Select
End Function

'***************************************************************************************************
' Author  : Chip Pearson at Pearson Software Consulting
' Purpose : Returns TRUE if the array is allocated (either a static array or a dynamic array that has been
'           sized with Redim) or FALSE if the array has not been allocated (a dynamic that has not yet
'           been sized with Redim, or a dynamic array that has been Erased).
' Source  : www.cpearson.com/excel/SortingArrays.aspx
'***************************************************************************************************
Public Function IsArrayAllocated(Arr As Variant) As Boolean
    Dim N As Long
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''
    ' If Arr is not an array, return FALSE and get out.
    '''''''''''''''''''''''''''''''''''''''''''''''''''
    If IsArray(Arr) = False Then
        IsArrayAllocated = False
        Exit Function
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Try to get the UBound of the array. If the array has not been allocated,
    ' an error will occur. Test Err.Number to see if an error occured.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    On Error Resume Next
    N = UBound(Arr, 1)
    If Err.number = 0 Then
        '''''''''''''''''''''''''''''''''''''
        ' No error. Array has been allocated.
        '''''''''''''''''''''''''''''''''''''
        IsArrayAllocated = True
    Else
        '''''''''''''''''''''''''''''''''''''
        ' Error. Unallocated array.
        '''''''''''''''''''''''''''''''''''''
        IsArrayAllocated = False
    End If
End Function

'***************************************************************************************************
' Author : Robert M Kreegier, Internet
' Purpose: Determines if stringToBeFound is in an array of strings (arr).
' Source : http://stackoverflow.com/questions/10951687/how-to-search-for-string-in-an-array
'***************************************************************************************************
Function IsInArray(stringToBeFound As String, Arr As Variant) As Boolean
    IsInArray = (UBound(Filter(Arr, stringToBeFound)) > -1)
End Function

'***************************************************************************************************
' Author : Scott Huish
' Purpose: Removes alphabetic characters from the string and returns the resulting string
' Source : http://www.mrexcel.com/forum/excel-questions/498357-strip-alpha-characters-string.html
'***************************************************************************************************
Function RemoveAlpha(ByVal strInput As String) As String
    With CreateObject("vbscript.regexp")
        .Pattern = "[A-Za-z]"
        .Global = True
        RemoveAlpha = .Replace(strInput, "")
    End With
End Function

'***************************************************************************************************
' Author : Robert M Kreegier
' Purpose: Removes duplicate and adjacent strTarget from within strOriginal and returns
'          the result.
'***************************************************************************************************
Function RemoveDuplicates(ByVal strOriginal As String, ByVal strTarget) As String
    Dim strTest As String
    While strTest <> strOriginal
        strTest = strOriginal
        strOriginal = Replace(strOriginal, strTarget & strTarget, strTarget)
    Wend
    
    RemoveDuplicates = strOriginal
End Function

'***************************************************************************************************
' Author : Robert M Kreegier
' Purpose: Returns True if vntInput is Nothing, False otherwise
'***************************************************************************************************
Function IsNothing(ByRef objInput As Object) As Boolean
    #If Not blnDeveloperMode Then
        On Error GoTo ProcException
    #End If
    '***********************************************************************************************
    
    IsNothing = True
    
    On Error Resume Next
    
    If Not objInput Is Nothing Then IsNothing = False
    
    #If Not blnDeveloperMode Then
        On Error GoTo ProcException
    #Else
        On Error GoTo 0
    #End If
    
    '***********************************************************************************************
ExitProc:
        Exit Function
ProcException:
    #If Not blnDeveloperMode Then
        InfoBox Err.Description & Chr(10) & "thrown from " & strModuleName & ": IsNothing", True
        Resume ExitProc
    #End If
End Function

'***************************************************************************************************
' Author : Robert M Kreegier
' Purpose: Returns True if vntInput is Not Nothing (or something)
'***************************************************************************************************
Function IsSomething(ByRef objInput As Object) As Boolean
    #If Not blnDeveloperMode Then
        On Error GoTo ProcException
    #End If
    '***********************************************************************************************
    
    IsSomething = False
    
    On Error Resume Next
    
    If Not objInput Is Nothing Then IsSomething = True
    
    #If Not blnDeveloperMode Then
        On Error GoTo ProcException
    #Else
        On Error GoTo 0
    #End If
    
    '***********************************************************************************************
ExitProc:
        Exit Function
ProcException:
    #If Not blnDeveloperMode Then
        InfoBox Err.Description & Chr(10) & "thrown from " & strModuleName & ": Isothing", True
        Resume ExitProc
    #End If
End Function

'***************************************************************************************************
' Author : Robert M Kreegier
' Purpose: Returns True if rngInput is a range that's been set
'***************************************************************************************************
Function IsSetRange(ByVal rngInput As Range) As Boolean
    IsSetRange = False

    On Error Resume Next

    If IsSomething(rngInput) Then
        Dim strAddress As String: strAddress = "Nothing"
        strAddress = rngInput.Address
        
        If strAddress <> "Nothing" Then IsSetRange = True
    End If

    On Error GoTo 0

End Function

'***************************************************************************************************
' Author : Robert M Kreegier
' Purpose: Returns True if there is an image by the name of strName in ActiveSheet, or
'          optionally ws
'***************************************************************************************************
Function IsImage(ByVal strName As String, Optional ByRef ws As Worksheet) As Boolean
    IsImage = False
    
    If ws Is Nothing Then Set ws = ActiveSheet
    
    On Error Resume Next
    If ws.Shapes(strName).Name <> strName Then
        IsImage = False
    Else
        IsImage = True
    End If
End Function

'***************************************************************************************************
' Author : Robert Kreegier
' Purpose: Helper function to get a single filename from the user, returned as a string
'***************************************************************************************************
Function OpenOneFileDialog(ByVal strTitle As String, ByVal strFilterName As String, ByVal strFilterString) As String
    With Application.FileDialog(msoFileDialogOpen)
        .AllowMultiSelect = False
        .Title = strTitle
        .Filters.Clear
        .Filters.Add strFilterName, strFilterString
        .InitialView = msoFileDialogViewDetails
        If .Show = -1 Then
            If .SelectedItems(1) <> vbNullString Then
                OpenOneFileDialog = .SelectedItems(1)
            End If
        End If
    End With
End Function

'***************************************************************************************************
' Author : Robert Kreegier
' Purpose: Opens a workbook from the provided filename
'***************************************************************************************************
Function OpenWB(Optional ByVal strFilePathAndName As String, Optional ByRef blnWasAlreadyOpen As Boolean = False) As Workbook
    #If Not blnDeveloperMode Then
        On Error GoTo ProcException
    #End If
    '***********************************************************************************************
    
    Set OpenWB = Nothing
    
    ' check validity of filepath, assign filepath a zero-length string if not valid
    If Not FileExists(strFilePathAndName) Then
        InfoBox "The file or path" & Chr(10) & strFilePathAndName & Chr(10) & "does not exist."
        strFilePathAndName = vbNullString
    End If
    
    ' Open the file dialog if we don't have a filepath
    If strFilePathAndName = vbNullString Then
        strFilePathAndName = OpenOneFileDialog("Select an Excel Spreadsheet to import...", "AHR Excel Files", "*.xls;*.xlsx;*.xlsm;*.xlsb")
    End If
    
    ' check validity of filepath again
    If Not strFilePathAndName = vbNullString Then
        If FileExists(strFilePathAndName) Then
            Dim wb As Workbook
            
            If IsWorkBookOpen(strFilePathAndName) Then
                blnWasAlreadyOpen = True
                Set wb = GetObject(strFilePathAndName)
            Else
                blnWasAlreadyOpen = False
                Set wb = Workbooks.Open(strFilePathAndName)
            End If
            
            If IsSomething(wb) Then
                Set OpenWB = wb
            End If
        
        Else
            InfoBox "The file or path" & Chr(10) & strFilePathAndName & Chr(10) & "does not exist."
        End If
    End If
    
    '***********************************************************************************************
ExitProc:
    Exit Function
ProcException:
    #If Not blnDeveloperMode Then
        InfoBox Err.Description & Chr(10) & "thrown from " & strModuleName & ": OpenWB", True
        Resume ExitProc
    #End If
End Function

'***************************************************************************************************
' Author : Robert M Kreegier
' Purpose: Returns True if the supplied strWorksheetName matches the name of a worksheet
'          in ThisWorkbook, and optionally wb.
'***************************************************************************************************
Function WorksheetExists(ByVal strWorksheetName As String, Optional ByRef wb As Workbook) As Boolean
    On Error Resume Next
    If IsNothing(wb) Then Set wb = ThisWorkbook
    WorksheetExists = (wb.Worksheets(strWorksheetName).Name <> vbNullString)
End Function

'***************************************************************************************************
' Author  : Unaccredited via the internet, editing by Robert M Kreegier
' Purpose : Returns true/false if a workbook is already open or not on this computer
' Source  : http://stackoverflow.com/questions/9373082/detect-whether-excel-workbook-is-already-open
'           https://support.microsoft.com/en-us/help/291295/macro-code-to-check-whether-a-file-is-already-open
'***************************************************************************************************
Function IsWorkBookOpen(ByVal strFileName As String) As Boolean
    #If Not blnDeveloperMode Then
        On Error GoTo ProcException
    #End If
    '***********************************************************************************************
    
    Dim ff As Long
    Dim ErrNo As Long

    On Error Resume Next
    ff = FreeFile()
    Open strFileName For Input Lock Read As #ff
    Close ff
    ErrNo = Err
    #If Not blnDeveloperMode Then
        On Error GoTo ProcException
    #Else
        On Error GoTo 0
    #End If

    Select Case ErrNo
        Case 0:    IsWorkBookOpen = False
        Case 70:   IsWorkBookOpen = True
        Case Else: Error ErrNo
    End Select
    
    '***********************************************************************************************
ExitProc:
    Exit Function
ProcException:
    #If Not blnDeveloperMode Then
        InfoBox Err.Description & Chr(10) & "thrown from " & strModuleName & ": IsWorkbookOpen", True
        Resume ExitProc
    #End If
End Function
'***************************************************************************************************
' Author : Robert M Kreegier
' Purpose: Determines if any worksheets are protected in ThisWorkbook, or optionally wb
'***************************************************************************************************
Function AnySheetsProtected(Optional ByRef wb As Workbook) As Boolean
    #If Not blnDeveloperMode Then
        On Error GoTo ProcException
    #End If
    '***********************************************************************************************
    
    AnySheetsProtected = False
    
    If wb Is Nothing Then Set wb = ThisWorkbook
    
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.ProtectContents = True Then
            AnySheetsProtected = True
        End If
    Next ws
    
    '***********************************************************************************************
ExitProc:
    Exit Function
ProcException:
    #If Not blnDeveloperMode Then
        InfoBox Err.Description & Chr(10) & "thrown from " & strModuleName & ": AnySheetsProtected", True
        Resume ExitProc
    #End If
End Function

'***************************************************************************************************
' Author : Robert M Kreegier
' Purpose: Determines if all the sheets are protected in ThisWorkbook, or optionally wb
'***************************************************************************************************
Function AllSheetsProtected(Optional ByRef wb As Workbook) As Boolean
    #If Not blnDeveloperMode Then
        On Error GoTo ProcException
    #End If
    '***********************************************************************************************
    
    If wb Is Nothing Then Set wb = ThisWorkbook
    
    AllSheetsProtected = True
    
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.CodeName <> "Variables_Global" Then
            If ws.ProtectContents = False Then
                AllSheetsProtected = False
                Exit For
            End If
        End If
    Next ws
    
    '***********************************************************************************************
ExitProc:
        Exit Function
ProcException:
    #If Not blnDeveloperMode Then
        InfoBox Err.Description & Chr(10) & "thrown from " & strModuleName & ": AllSheetsProtected", True
        Resume ExitProc
    #End If
End Function

'***************************************************************************************************
' Author : Robert M Kreegier, Mark Borgerding
' Purpose: Convert an integer to its string representation in a given base
' Source : http://stackoverflow.com/questions/2267362/convert-integer-to-a-string-in-a-given-numeric-base-in-python/28666223#28666223
'***************************************************************************************************
Function IntToBase(ByVal number, ByVal base, Optional ByVal alphabet = "0123456789abcdefghijklmnopqrstuvwxyz")
    Dim strResult As String: strResult = vbNullString
    Dim idx As Long
    
    If base < 2 Or base > Len(alphabet) Then
        If base = 64 Then ' assume base64 rather than raise error
            alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
        Else
            Status "base out of range"
        End If
    End If
    
    If number <= 0 Then
        If number = 0 Then
            IntToBase = Mid(alphabet, 1, 1)
            Exit Function
        Else
            IntToBase = "-" & IntToBase(-number, base, alphabet)
            Exit Function
        End If
        
    ' else number is non-negative real
    Else
        While number > 0
            idx = number Mod base
            number = Int(number / base)
            strResult = Mid(alphabet, idx + 1, 1) & strResult
        Wend
    End If
    
    IntToBase = strResult
End Function

'***************************************************************************************************
' Author : Robert M Kreegier, robartsd
' Purpose: Converts a row and column number to an A1 address
'***************************************************************************************************
Function Addy(ByVal lngRow As Long, ByVal lngCol As Long) As String
    #If Not blnDeveloperMode Then
        On Error GoTo ProcException
    #End If
    '***********************************************************************************************
    
    Addy = Cells(lngRow, lngCol).Address
    
    '***********************************************************************************************
ExitProc:
        Exit Function
ProcException:
    #If Not blnDeveloperMode Then
        InfoBox Err.Description & Chr(10) & "thrown from " & strModuleName & ": Addy", True
        Resume ExitProc
    #End If
End Function

'***************************************************************************************************
' Author : Robert M Kreegier, robartsd
' Purpose: Converts a column number (lngCol) to it's corresponding alphabetic letter
' Source : http://stackoverflow.com/questions/12796973/function-to-convert-column-number-to-letter
'***************************************************************************************************
Function CLet(ByVal lngCol As Long) As String
    #If Not blnDeveloperMode Then
        On Error GoTo ProcException
    #End If
    '***********************************************************************************************
    
    Dim N As Long
    Dim C As Byte
    Dim s As String

    If lngCol > 0 And lngCol <= Columns.Count Then
        N = lngCol
        
        Do
            C = ((N - 1) Mod 26)
            s = Chr(C + 65) & s
            N = (N - C) \ 26
        Loop While N > 0
        
        CLet = s
    Else
        InfoBox "CLet: The column number must be" & Chr(10) & "0 < lngCol < Columns.count", True
        CLet = vbNullString
    End If
    
    '***********************************************************************************************
ExitProc:
        Exit Function
ProcException:
    #If Not blnDeveloperMode Then
        InfoBox Err.Description & Chr(10) & "thrown from " & strModuleName & ": CLet", True
        Resume ExitProc
    #End If
End Function

'***************************************************************************************************
' Author : Robert M Kreegier, His Nibbs
' Purpose: Converts a column letter (strCol) to it's corresponding number
' Source : http://www.vbforums.com/showthread.php?680083-VBA-How-to-convert-column-alphabet-to-number
'***************************************************************************************************
Function CNum(ByVal strCol As String) As Long
    #If Not blnDeveloperMode Then
        On Error GoTo ProcException
    #End If
    '***********************************************************************************************
    
    Dim lngResult As Long: lngResult = 0
    Dim lngIndex As Long
    
    For lngIndex = 1 To Len(strCol)
        lngResult = lngResult * 26 + (Asc(UCase(Mid(strCol, lngIndex, 1))) - 64)
    Next
    
    CNum = lngResult
    
    '***********************************************************************************************
ExitProc:
        Exit Function
ProcException:
    #If Not blnDeveloperMode Then
        InfoBox Err.Description & Chr(10) & "thrown from " & strModuleName & ": CNum", True
        Resume ExitProc
    #End If
End Function

'***************************************************************************************************
' Author : Robert M Kreegier
' Purpose: Converts a hex string (RGBColor) to a system color for use in forms
'***************************************************************************************************
Function HexToSystemColor(ByVal RGBColor As String)
    HexToSystemColor = "&H00" & Hex(RGBColor)
End Function

'***************************************************************************************************
' Author : Robert M Kreegier
' Purpose: GetCodeName takes a component name (strName) and returns it's corresponding
'          code name, if available
'***************************************************************************************************
Function GetCodeName(ByVal strName As String, ByRef wb As Workbook)
    #If Not blnDeveloperMode Then
        On Error GoTo ProcException
    #End If
    '***********************************************************************************************
    
    Dim wbcomp
    Dim sheet_name
    
    If strName = "ThisWorkbook" Then
        GetCodeName = "ThisWorkbook"
        Exit Function
    End If
    
    ' loop through all the components and look for the one that matches strName
    For Each wbcomp In wb.VBProject.VBComponents
        If wbcomp.Type = 100 Then
            sheet_name = wbcomp.Properties("Name").Value
        Else
            sheet_name = wbcomp.Name
        End If
        
        If sheet_name = strName Then
            GetCodeName = wbcomp.Name
            Exit Function
        End If
    Next
    
    GetCodeName = vbNullString
    
    '***********************************************************************************************
ExitProc:
        Exit Function
ProcException:
    #If Not blnDeveloperMode Then
        InfoBox Err.Description & Chr(10) & "thrown from " & strModuleName & ": GetCodeName", True
        Resume ExitProc
    #End If
End Function

'***************************************************************************************************
' Author : Robert M Kreegier
' Purpose: Inserts a string (strInsert) into strOriginal at the specified intIndex
'***************************************************************************************************
Function StrIntoStr(ByVal strOriginal As String, ByVal strInsert As String, ByVal intIndex As Integer) As String
    StrIntoStr = Mid(strOriginal, 1, intIndex) & strInsert & Mid(strOriginal, intIndex + 1, Len(strOriginal) - intIndex)
End Function

'***************************************************************************************************
' Author : Robert M Kreegier
' Purpose: Returns a string that's within strString, defined by the start (lngX) and end (lngY)
'          position.
'***************************************************************************************************
Function MidXY(ByVal strString As String, ByVal lngX As Long, ByVal lngY As Long) As String
    MidXY = Mid(strString, lngX, (lngY - lngX + 1))
End Function

'***************************************************************************************************
' Author : Robert M Kreegier
' Purpose: Retreives the last row of the worksheet ws
'***************************************************************************************************
Function GetLastRow(ByRef ws As Worksheet) As Long
    'GetLastRow = ws.Range("A1048576").End(xlUp).Row ' doesn't work when last row height is zero
    GetLastRow = ws.Cells.Find("*", SearchOrder:=xlByRows, LookIn:=xlFormulas, searchdirection:=xlPrevious).Row
End Function

'***************************************************************************************************
' Author : David Lee, adaptation by Robert M Kreegier
' Purpose: Checks if a string is fully alphabetic
' Source : https://techniclee.wordpress.com/2010/07/21/isletter-function-for-vba/
'***************************************************************************************************
Function IsLetter(strValue As String) As Boolean
    Dim intPos As Integer
    For intPos = 1 To Len(strValue)
        Select Case Asc(Mid(strValue, intPos, 1))
            Case 65 To 90, 97 To 122
                IsLetter = True
            Case Else
                IsLetter = False
                Exit For
        End Select
    Next
End Function

'***************************************************************************************************
' Author : Robert M Kreegier
' Purpose: Retreives the last column of the worksheet ws
'***************************************************************************************************
Function GetLastCol(Optional ByRef ws As Worksheet) As Long
    'GetLastCol = ws.Range(wsheet.PageSetup.PrintArea).Columns.Count                                   ' no good if we're using this to actually set the print area
    ''GetLastCol = ws.Cells(2, wsheet.Columns.Count).End(xlToLeft).Column                              ' close, but still not good if the last cell is merged. will return the column of the left most merged cell
    '''GetLastCol = ws.Cells.Find("*", searchorder:=xlByColumns, LookIn:=xlFormulas, searchdirection:=xlPrevious).Column   ' okay, but will only return the right most column with stuff in it
    
    ' this returns the right most column regardless of stuff in the cells or merges
    If IsNothing(ws) Then Set ws = ActiveSheet
    
    Dim lastcol As Long: lastcol = ws.Cells.Find("*", SearchOrder:=xlByColumns, LookIn:=xlFormulas, searchdirection:=xlPrevious).Column ' ws.Cells(2, ws.Columns.count).End(xlToLeft).Column
    
    Dim mergecells As Range
    Set mergecells = ws.Cells(2, lastcol).MergeArea

    GetLastCol = mergecells.Columns(mergecells.Columns.Count).Column
End Function

'***************************************************************************************************
' Author : RichardSchollar
' Purpose: Checks to see if the given string is a valid range address
' Source : http://www.ozgrid.com/forum/showthread.php?t=95645
'***************************************************************************************************
Function IsAddress(strAddress As String) As Boolean
    On Error Resume Next
    Dim R As Range: Set R = Range(strAddress)
    If IsSomething(R) Then IsAddress = True
End Function

'***************************************************************************************
' Author : Charles H. Pearson
' Purpose: A Union operation that accepts parameters that are Nothing.
' Source : www.cpearson.com/Excel/BetterUnion.aspx
'***************************************************************************************
Function Union2(ParamArray Ranges() As Variant) As Range
    Dim N As Long
    Dim RR As Range
    For N = LBound(Ranges) To UBound(Ranges)
        If IsObject(Ranges(N)) Then
            If Not Ranges(N) Is Nothing Then
                If TypeOf Ranges(N) Is Excel.Range Then
                    If Not RR Is Nothing Then
                        Set RR = Application.Union(RR, Ranges(N))
                    Else
                        Set RR = Ranges(N)
                    End If
                End If
            End If
        End If
    Next N
    Set Union2 = RR
End Function

'***************************************************************************************
' Author : Charles H. Pearson
' Purpose: This provides Union functionality without duplicating cells when ranges
'          overlap. Requires the Union2 function.
' Source : www.cpearson.com/Excel/BetterUnion.aspx
'***************************************************************************************
Function ProperUnion(ParamArray Ranges() As Variant) As Range
    Dim ResR As Range
    Dim N As Long
    Dim R As Range
    
    If Not Ranges(LBound(Ranges)) Is Nothing Then
        Set ResR = Ranges(LBound(Ranges))
    End If
    For N = LBound(Ranges) + 1 To UBound(Ranges)
        If Not Ranges(N) Is Nothing Then
            For Each R In Ranges(N).Cells
                If Application.Intersect(ResR, R) Is Nothing Then
                    Set ResR = Union2(ResR, R)
                End If
            Next R
        End If
    Next N
    Set ProperUnion = ResR
End Function

'***************************************************************************************************
' Author : Robert M Kreegier
' Purpose: Exports ActiveSheet as a PDF and optionally opens it up for display after saving
'***************************************************************************************************
Function ExportActiveSheetAsPDF(Optional ByVal OpenAfterPublish As Boolean = True) As String
    #If Not blnDeveloperMode Then
        On Error GoTo ProcException
    #End If
    '***********************************************************************************************
    
    Dim datestr As String: datestr = GetCycleDateString
    Dim facility As String: facility = GetVar("V_FAC_NAME")
    Dim Area As String: Area = GetVar("V_FAC_AREA")
    Dim FileName As String: FileName = datestr & " " & facility & "-" & Area & "-AHR" & ".pdf"
    Dim fullfilepath As String: fullfilepath = Left(ThisWorkbook.FullName, InStrRev(ThisWorkbook.FullName, "\")) & FileName
    Dim saveasfilename As String: saveasfilename = Application.GetSaveAsFilename(fullfilepath, "Adobe PDF Files (*.pdf), *.pdf")
    
    'Dim saveasfilename As String: saveasfilename = Application.GetSaveAsFilename(FileFilter:="Adobe PDF Files (*.pdf), *.pdf")
    
    If saveasfilename <> vbNullString And saveasfilename <> "False" Then
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
                                        FileName:=saveasfilename, _
                                        Quality:=xlQualityStandard, _
                                        IncludeDocProperties:=True, _
                                        IgnorePrintAreas:=False, _
                                        OpenAfterPublish:=OpenAfterPublish
    End If
    
    ExportActiveSheetAsPDF = saveasfilename
        
    '***********************************************************************************************
ExitProc:
        Exit Function
ProcException:
    #If Not blnDeveloperMode Then
        InfoBox Err.Description & Chr(10) & "thrown from " & strModuleName & ": ExportActiveSheetAsPDF", True
        Resume ExitProc
    #End If
End Function

'***************************************************************************************************
' Author : Microsoft
' Purpose: Determines if a file by the name of strFileName is currently open by the system
' Source : https://support.microsoft.com/en-us/kb/291295
'***************************************************************************************************
Function IsFileOpen(ByVal strFileName As String)
    Dim filenum As Integer, errnum As Integer

    On Error Resume Next   ' Turn error checking off.
    filenum = FreeFile()   ' Get a free file number.
    ' Attempt to open the file and lock it.
    Open strFileName For Input Lock Read As #filenum
    Close filenum          ' Close the file.
    errnum = Err           ' Save the error number that occurred.
    On Error GoTo 0        ' Turn error checking back on.

    ' Check to see which error occurred.
    Select Case errnum

        ' No error occurred.
        ' File is NOT already open by another user.
        Case 0
         IsFileOpen = False

        ' Error number for "Permission Denied."
        ' File is already opened by another user.
        Case 70
            IsFileOpen = True

        ' Another error occurred.
        Case Else
            Error errnum
    End Select
End Function

'***************************************************************************************************
' Author : Robert M Kreegier
' Purpose: Takes a timestamp of the format "1/6/2017 12:54:13 AM" and returns it as a Date object
'***************************************************************************************************
Function TimeStampToDate(ByVal strDateTime As String) As Date
    ' must be in the format 1/6/2017 1:54:13 AM
    If Len(strDateTime) >= 19 Then
        Dim strDate As String: strDate = Left(strDateTime, Len(strDateTime) - 11)
        Dim strTime As String: strTime = Right(strDateTime, 11)
        
        If IsDate(strDateTime) Then
            TimeStampToDate = DateValue(strDate) + TimeValue(strTime)
        End If
    End If
End Function

'***************************************************************************************************
' Author : Robert M Kreegier
' Purpose: Returns a timestamp in the format of "1/6/2017 12:54:13 AM"
'***************************************************************************************************
Function GetTimeStamp() As Date
    GetTimeStamp = DateValue(Date) + TimeValue(Time)
End Function

'***************************************************************************************************
' Author : Robert M Kreegier
' Purpose: Determines if the provided date (dtmDate) is within a leap year and returns True/False
' Notes  : This does not accept just the year. It must be a full date: "1/1/2016" rather than "2016"
'***************************************************************************************************
Function IsLeapYear(ByVal dtmDate As Date) As Boolean
    ' This works in a tricky way by taking advantage of VBA's IsDate. We'll presume that since Microsoft is a multi-billion
    ' dollar company, they have their ducks in a row and know how to properly check if Feb 29 exists in any particular year.
    ' We'll assume this is a more robust option than what's written below because of the multiple rules regarding leap years
    ' and days and seconds and centuries and whatever else astronomers decide to come up with. Instead of reinventing the
    ' wheel, we'll piggy-back off of Microsoft's work.
    IsLeapYear = IsDate("2/29/" & Year(dtmDate))
    
    ' this also works...
    ' source: http://excelribbon.tips.net/T009978_Determining_If_a_Year_is_a_Leap_Year.html
'    Dim YearNo As Long: YearNo = Year(dtmDate)
'
'    If YearNo Mod 100 = 0 Then
'       IsLeapYear = ((YearNo \ 100) Mod 4 = 0)
'    Else
'       IsLeapYear = (YearNo Mod 4 = 0)
'    End If
End Function

'***************************************************************************************************
' Author  : Allen Wyatt
' Purpose : Selects all visible worksheets
' Source  : http://excel.tips.net/T003058_Selecting_All_Visible_Worksheets_in_a_Macro.html
'***************************************************************************************************
Sub SelectAllSheets()
    Dim mySheet As Worksheet
    For Each mySheet In ThisWorkbook.Sheets
        With mySheet
            If .Visible = True Then .Select Replace:=False
        End With
    Next mySheet
End Sub

'***************************************************************************************************
' Author  : Robert Kreegier
' Purpose : Exports the active worksheet to a new workbook, without any macro code, rendering it
'           "dead"
'***************************************************************************************************
Function ExportActiveWSasDeadWB(Optional ByVal strSaveAsPath As String = vbNullString) As Workbook
    #If Not blnDeveloperMode Then
        On Error GoTo ProcException
    #End If
    '***********************************************************************************************
    
    Dim strDate As String: strDate = GetCycleDateString
    Dim strFacName As String: strFacName = GetVar("V_FAC_NAME")
    Dim strFacArea As String: strFacArea = GetVar("V_FAC_AREA")
    Dim strFileName As String: strFileName = strDate & " " & strFacName & "-" & strFacArea & "-AHR" & ".xlsx"
    Dim strFullPath As String: strFullPath = Left(ThisWorkbook.FullName, InStrRev(ThisWorkbook.FullName, "\")) & strFileName
    Dim wsActiveWS As Worksheet: Set wsActiveWS = ActiveSheet
    
    Dim blnCancel As Boolean: blnCancel = False
    While strSaveAsPath = vbNullString And blnCancel = False
        strSaveAsPath = Application.GetSaveAsFilename(strFullPath, "Excel Workbook (*.xlsx), *.xlsx")
        
        ' check if this file already exists
        If strSaveAsPath <> vbNullString And Dir(strSaveAsPath) <> vbNullString Then
            Select Case MsgBox("It appears this file already exists. Would you like to overwrite it?", vbYesNoCancel + vbExclamation + vbDefaultButton3, "Overwrite?")
                Case vbNo
                    strSaveAsPath = vbNullString
                    
                Case vbCancel
                    strSaveAsPath = vbNullString
                    blnCancel = True
            End Select
        End If
    Wend
    
    If strSaveAsPath <> vbNullString And strSaveAsPath <> "False" Then
        Dim NewBook As Workbook: Set NewBook = Workbooks.Add
        
        ' copy/paste the worksheet to the new workbook
        wsActiveWS.Copy Before:=NewBook.Sheets(1)
        
        ' copy and paste the values to get rid of any formulas
        NewBook.Sheets(1).Cells.Copy
        NewBook.Sheets(1).Range("A1").PasteSpecial Paste:=xlPasteValues

        ' Get the code/object name of the new sheet...
        Dim strObjectName As String
        strObjectName = NewBook.Sheets(1).CodeName
        
        ' Remove all lines from its code module...
        With NewBook.VBProject.VBComponents(strObjectName).CodeModule
            .DeleteLines 1, .CountOfLines
        End With
        
        ' The whole sheet ends up selected, so let's select cell A1 to clear that minor annoyance
        NewBook.Sheets(1).Cells(1, 1).Select
        
        ' Removed the range names
        On Error Resume Next
        Dim nmName As Name
        For Each nmName In NewBook.Names
            nmName.Delete
        Next
        #If Not blnDeveloperMode Then
            On Error GoTo ProcException
        #Else
            On Error GoTo 0
        #End If
        
        ' Remove the link in the Area Name image and set its text
        NewBook.Sheets(1).Shapes("Area Name").OLEFormat.Object.Formula = vbNullString
        
        ' edit the attributes of the new workbook
        With NewBook
            .Title = "SKF Asset Health Report - " & strFacName & " - " & strFacArea
            .SaveAs FileName:=strSaveAsPath
            '.Close
        End With
        
        Set ExportActiveWSasDeadWB = NewBook
    End If
        
    '***********************************************************************************************
ExitProc:
        Exit Function
ProcException:
    #If Not blnDeveloperMode Then
        InfoBox Err.Description & Chr(10) & "thrown from " & strModuleName & ": ExportActiveWSasDeadWB", True
        Resume ExitProc
    #End If
End Function

'***************************************************************************************************
' Author  : Robert Kreegier
' Purpose : SwapInEscapeChars parses a text string and replaces special chars with their associated
'           escape char sequence.
'***************************************************************************************************
Function SwapInEscapeChars(ByVal strOriginalText As String) As String
    #If Not blnDeveloperMode Then
        On Error GoTo ProcException
    #End If
    '************************************************************
    
    SwapInEscapeChars = vbNullString
    
    strOriginalText = Replace(strOriginalText, Chr(13), Chr(10))
    strOriginalText = RemoveDuplicates(strOriginalText, Chr(10))
    strOriginalText = Replace(strOriginalText, Chr(10), "\n")
    
    SwapInEscapeChars = strOriginalText

    '************************************************************
ExitProc:
    Exit Function
ProcException:
    #If Not blnDeveloperMode Then
        InfoBox Err.Description & Chr(10) & "thrown from " & strModuleName & ": SwapInEscapeChars", True
        Resume ExitProc
    #End If
End Function

'***************************************************************************************************
' Author  : Robert Kreegier
' Purpose : SwapOutEscapeChars parses a text string and replaces special chars with their associated
'           escape char sequence.
'***************************************************************************************************
Function SwapOutEscapeChars(ByVal strOriginalText As String) As String
    #If Not blnDeveloperMode Then
        On Error GoTo ProcException
    #End If
    '************************************************************
    
    SwapOutEscapeChars = vbNullString
    
    strOriginalText = Replace(strOriginalText, "\n", Chr(10))
    
    SwapOutEscapeChars = strOriginalText

    '************************************************************
ExitProc:
    Exit Function
ProcException:
    #If Not blnDeveloperMode Then
        InfoBox Err.Description & Chr(10) & "thrown from " & strModuleName & ": SwapOutEscapeChars", True
        Resume ExitProc
    #End If
End Function

'***************************************************************************************************
' Author : Daniel Pineault, CARDA Consultants Inc., Modified slightly by Robert M Kreegier to just deal with the Interior object
' Purpose: Copy a Cell/Range's gradient fill properties from one Cell/Range to another
' Source : http://www.cardaconsultants.com
' Params :
'   FromInterior  = Range that contains the Background properties to be copied
'   ToInterior    = Range you wish to copy the Background properties to
' Example:
'   CopyBkGrnd Range("B5"), Range("A1")
'   CopyBkGrnd Sheet1.Range("B5"), Sheet2.Range("A1:B2")
'***************************************************************************************************
Sub CopyBkGrnd(ByRef FromInterior As Interior, ByRef ToInterior As Interior)
    #If Not blnDeveloperMode Then
        On Error GoTo ProcException
    #End If
    '***********************************************************************************************
    
    On Error Resume Next
    
    'ToInterior.Pattern = FromInterior.Pattern
    If FromInterior.Gradient.ColorStops.Count = 0 Then
        'Solid Fill Color Properties
        ToInterior.PatternColorIndex = FromInterior.PatternColorIndex
        ToInterior.TintAndShade = FromInterior.TintAndShade
        ToInterior.PatternTintAndShade = FromInterior.PatternTintAndShade
        ToInterior.Color = FromInterior.Color
    Else
        'Gradient Fill Properties
        ToInterior.Gradient.RectangleLeft = FromInterior.Gradient.RectangleLeft
        ToInterior.Gradient.RectangleRight = FromInterior.Gradient.RectangleRight
        ToInterior.Gradient.RectangleTop = FromInterior.Gradient.RectangleTop
        ToInterior.Gradient.RectangleBottom = FromInterior.Gradient.RectangleBottom
        ToInterior.Gradient.Degree = FromInterior.Gradient.Degree
        ToInterior.Gradient.ColorStops.Clear
        Dim i
        For i = 1 To FromInterior.Gradient.ColorStops.Count
            With ToInterior.Gradient.ColorStops.Add(i - 1)
                .ThemeColor = FromInterior.Gradient.ColorStops(i).ThemeColor
                .TintAndShade = FromInterior.Gradient.ColorStops(i).TintAndShade
                .Color = FromInterior.Gradient.ColorStops(i).Color
            End With
        Next i
    End If
    
    '***********************************************************************************************
ExitProc:
    Exit Sub
ProcException:
    #If Not blnDeveloperMode Then
        InfoBox Err.Description & Chr(10) & "thrown from " & strModuleName & ": CopyBkGrnd", True
        Resume ExitProc
    #End If
End Sub
