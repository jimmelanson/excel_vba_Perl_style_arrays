VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'   Created by Jim Melanson, 2017.05.07
'   jmelanson1965@gmail.com
'
' REQUIRES MSExcel 2010 or higher
'

Private p_Text As String
Private p_Delimeter As String
Private p_CountElements As Long
Private p_ForceValue As Boolean
Private p_ForceUnique As Boolean
Private p_SortAsNumbers As Boolean
Private p_SortAsText As Boolean
Private p_GrepIgnoreCase As Boolean
Private p_SliceIgnoreCase As Boolean
Private p_RemoveIgnoreCase As Boolean

Private Sub Class_Initialize()
    p_CountElements = 0
    p_Delimeter = ","
    p_Text = ""
    p_SortAsNumbers = False
    p_SortAsText = True
    p_ForceValue = False
    p_ForceUnique = False
    p_GrepIgnoreCase = False
    p_SliceIgnoreCase = False
End Sub

Private Sub Class_Terminate()
    p_CountElements = 0
    p_Delimeter = ","
    p_Text = ""
End Sub

Private Sub Reset()
    p_CountElements = 0
    p_Text = ""
End Sub

Property Get Delimeter() As String
    Delimeter = p_Delimeter
End Property
Property Let Delimeter(ByVal strValue As String)
    If strValue <> "" Then
        p_Delimeter = strValue
    End If
End Property

Property Get ForceValue() As Boolean
    ForceValues = p_ForceValue
End Property
Property Let ForceValue(ByVal boolValue As Boolean)
    If boolValue = True Then
        p_ForceValue = True
    Else
        p_ForceValue = False
    End If
End Property

Property Get ForceUnique() As Boolean
    ForceUnique = p_ForceUnique
End Property
Property Let ForceUnique(ByVal boolValue As Boolean)
    If boolValue = True Then
        p_ForceUnique = True
    Else
        p_ForceUnique = False
    End If
End Property

Property Get GrepIgnoreCase() As Boolean
    GrepIgnoreCase = p_GrepIgnoreCase
End Property
Property Let GrepIgnoreCase(ByVal boolValue As Boolean)
    If boolValue = True Then
        p_GrepIgnoreCase = True
    Else
        p_GrepIgnoreCase = False
    End If
End Property

Property Get SliceIgnoreCase() As Boolean
    SliceIgnoreCase = p_SliceIgnoreCase
End Property
Property Let SliceIgnoreCase(ByVal boolValue As Boolean)
    If boolValue = True Then
        p_SliceIgnoreCase = True
    Else
        p_SliceIgnoreCase = False
    End If
End Property

Property Get RemoveIgnoreCase() As Boolean
    RemoveIgnoreCase = p_RemoveIgnoreCase
End Property
Property Let RemoveIgnoreCase(ByVal boolValue As Boolean)
    If boolValue = True Then
        p_RemoveIgnoreCase = True
    Else
        p_RemoveIgnoreCase = False
    End If
End Property


'########################################
'#
'#  PUBLIC METHODS
'#
'########################################

'########################################
'#
'#  Counting
'#
'########################################

Public Function LastIndex() As Long
    LastIndex = (CountElements - 1)
End Function

Public Function CountElements() As Long
    'Fix some common errors
    p_CountElements = 0
    CountElements = 0
    If p_Text <> "" Then
        If Mid(p_Text, 1, 1) = p_Delimeter Then
            'Delimeter at front of string list
            p_Text = Mid(p_Text, 2, Len(p_Text))
        End If
        If Mid(p_Text, Len(p_Text), 1) = p_Delimeter Then
            'delimeter at end of string list
            p_Text = Mid(p_Text, 1, Len(p_Text) - 1)
        End If
    End If
    If Len(p_Text) > 0 And InStr(p_Text, p_Delimeter) = 0 Then
        'There is a stringlist but there are no delimeters; this means there is ONE element in the array.
        p_CountElements = 1
        CountElements = 1
    Else
        'As we iterate through the string list, we are reading characters then a delimeter, more characters
        'and then another delimeter, etc. This means that every time you encounter a delimeter, the characters
        'to the left of it are an element that gets counted.
        '
        'Because of our common fix at the start of this procedure, a string list will never end with a delimeter.
        'This means that however many delimeters there are, the element count will always be one higher. This is
        'because after the last delimeter, there will still be characters to the right of that last delimeter.
        'As they do not have a delimeter after them, they never get examined. This means that no matter how many
        'delimeters there are, there is always one more elements.
        Dim iChrs As Long
        For iChrs = 1 To Len(p_Text)
            If Mid(p_Text, iChrs, 1) = p_Delimeter Then
                CountElements = CountElements + 1
            End If
        Next iChrs
        If CountElements > 0 Then
            CountElements = CountElements + 1
        End If
        p_CountElements = CountElements
    End If
End Function

'########################################
'#
'#  Populate Object Array w/List
'#
'########################################

Private Function SetArrayFromString_CompareCount(ByVal content As String, ByVal compare As String, ByVal longNotAt As Long) As Long
    SetArrayFromString_CompareCount = 0
    'This is the easy way to see if the item is within the p_Text
    'If this function returns anything higher than 0, it is a duplicate
    If Right(content, Len(compare) + 1) = p_Delimeter & compare Then
        'Check last element
        If longNotAt <> (Len(content) - Len(compare)) + 1 Then
            SetArrayFromString_CompareCount = 1
        End If
    ElseIf Left(content, Len(compare) + 1) = compare & p_Delimeter Then
        'Check first element
        If longNotAt <> 1 Then
            SetArrayFromString_CompareCount = 1
        End If
    Else
        'Check all the nth elements
        Dim iChars As Long
        For iChars = (Len(content) - Len(compare)) To 1 Step -1
            'Looking in: blue,yellow,red,green,purple
            'Looking for:   ,red,
            'Red is at: Mid(string, 13, 3)
            ',red, is at: Mid(string, 12, 5)
            If Mid(content, iChars, Len(compare) + 2) = p_Delimeter & compare & p_Delimeter Then
                If iChars + 1 <> longNotAt Then
                    SetArrayFromString_CompareCount = SetArrayFromString_CompareCount + 1
                End If
            End If
        Next iChars
    End If
End Function

Public Sub SetArrayFromString(ByVal content As String)
    Reset
    p_CountElements = 0
    If content <> "" Then
        p_Text = content
        'Check to see if we are forcing non-empty values/fix delimeter mistakes
        If p_ForceValue = True Then
            'We are forcing values. The input is in the format:
            'contentDELIMETERcontentDELIMETERcontentDELIMETER
            'So we need to look for:
            'contentDELIMETERDELIMETERcontentDELIMETERcontentDELIMETER
            Dim regexThis As Object
            Set regexThis = New RegExp
            regexThis.Global = True
            regexThis.Pattern = p_Delimeter & p_Delimeter & "+"
            p_Text = regexThis.Replace(p_Text, p_Delimeter)
            'Remove any leading delimeter
            If Mid(p_Text, 1, 1) = p_Delimeter Then
                p_Text = Mid(p_Text, 2, Len(p_Text))
            End If
            'Remove any trailing delimeter
            If Mid(p_Text, Len(p_Text), 1) = p_Delimeter Then
                p_Text = Mid(p_Text, 1, Len(p_Text) - 1)
            End If
        End If

        If InStr(p_Text, p_Delimeter) = 0 Then
            'This is a single item list so no more checks needed.
            p_CountElements = 1
        Else
            'Principal iterator
            Dim iChars As Long
            If p_ForceUnique = True Then
                'We may be modifying the list as we move along if necessary
                'This will be the new list we restore to the p_Text property
                Dim strNewText As String
                'Value holder
                Dim iLastDelimeter As Long
                iLastDelimeter = 0
                'Execute the loop
                For iChars = Len(p_Text) To 1 Step -1
                    'check to see we are not matching the first element if this is not the first element
                    'check to see we are not matching the last element if this is not the last element
                    'check to see we are not matching an nth element, if this is not the same nth element
                    If iChars = 1 Then
                        'We have gone through the whole list and are now
                        'at the very first element, so let's check it.
                        If SetArrayFromString_CompareCount(p_Text, Mid(p_Text, 1, iLastDelimeter - 1), 1) = 0 Then
                            'Not a duplicate so add it in AND remove it from the original
                            If strNewText <> "" Then
                                strNewText = p_Delimeter & strNewText
                            End If
                            strNewText = Mid(p_Text, 1, iLastDelimeter - 1) & strNewText
                        End If
                    Else
                        If Mid(p_Text, iChars, 1) = p_Delimeter Then
                            'Not the first item in the list so is nth item OR last item
                            If iLastDelimeter = 0 Then
                                'If this is zero, it is the last item in the list, the first item
                                'that we are checking.
                                If SetArrayFromString_CompareCount(p_Text, Mid(p_Text, iChars + 1, Len(p_Text)), iChars + 1) = 0 Then
                                    'Not a duplicate so add it in AND remove it from the original
                                    If strNewText <> "" Then
                                        strNewText = p_Delimeter & strNewText
                                    End If
                                    strNewText = Mid(p_Text, iChars + 1, Len(p_Text)) & strNewText
                                End If
                            Else
                                'nth element
                                '
                                'We have already had an array elelment checked because iLastDelimeter > 0
                                'So now we are looking at a delimeter, so we have to evaluate from
                                'this delimeter up to the last delimeter. However, those values would
                                'include this delimeter and the last delimeter. So that means we have
                                'to take the substring from this delimeter PLUS 1 character, to the
                                'last delimeter MINUS one character. Remember, we are working the string
                                'from back to front.
                                'So is the string was - blue,yellow,red,gold
                                '...and the current character is (in a 1-based index): 5
                                '...then the last delimeter would be (in a 1-based index): 12
                                '...so from 5 to 12 would be: ,yellow,
                                '...so we do this character + 1 (6) to the last delimeter minus 1 (11)
                                '...which gives us: yellow
                                'The snipped will be: Mid(p_Text, iChars + 1, (iLastDelimeter - iChars) - 1)
                                'To calculate the legth of the snippet we subtract the current position of
                                'the search from the last delimeter position.
                                'iLastDelimeter = 12
                                'iLastDelimeter - iChars = 12 - 5
                                'Length of substringis 12 - 5 = 7
                                'We also need to subtract 1 more character because the start of that subtraction
                                'is a delimeter and not he actual word. So, the length of the element is:
                                '(Last Delimeter - Current Delimeter) - one character offset
                                '(12 - 5) - 1 = 6
                                'Yellow has 6 characters
                                '
                                'SetArrayFromString_CompareCount(p_Text, Mid(p_Text, iChars + 1, (iLastDelimeter - iChars) - 1), iChars + 1)
                                '                                LIST  ,                Snippet                                , START POS
                                If SetArrayFromString_CompareCount(p_Text, Mid(p_Text, iChars + 1, (iLastDelimeter - iChars) - 1), iChars + 1) = 0 Then
                                    'Not a duplicate so add it in AND remove it from the original
                                    If strNewText <> "" Then
                                        strNewText = p_Delimeter & strNewText
                                    End If
                                    strNewText = Mid(p_Text, iChars + 1, (iLastDelimeter - iChars) - 1) & strNewText
                                End If
                            End If
                            iLastDelimeter = iChars
                        End If
                    End If
                Next iChars
                p_Text = strNewText
            End If
            'There is a delimeter so there is more than one element
            'Count the elements
            For iChars = 1 To Len(p_Text)
                If Mid(p_Text, iChars, 1) = p_Delimeter Then
                    p_CountElements = p_CountElements + 1
                End If
            Next iChars
            'Add one more because the last element does not end with a delimeter
            p_CountElements = p_CountElements + 1
        End If
    End If
End Sub

Public Sub SetArray(list() As Variant)
    Reset
    Dim i As Long
    Dim iCompare As Long
    Dim countCompare As Long
    For i = LBound(list) To UBound(list)
        If list(i) <> "" Then
            countCompare = 0
            'If the list item has content, we don't have to worry about force value.
            '
            'Now we have to look at if we are focing unique
            If p_ForceUnique = True Then
                'We have to count how many times this item appears in the list
                For iCompare = LBound(list) To UBound(list)
                    If list(iCompare) = list(i) Then
                        countCompare = countCompare + 1
                    End If
                Next iCompare
                If countCompare = 1 Then
                    If p_Text <> "" Then
                        p_Text = p_Text & p_Delimeter
                    End If
                    p_Text = p_Text & list(i)
                End If
            Else
                'The item has content and ForceUnique is false, so we can just add it.
                If p_Text <> "" Then
                    p_Text = p_Text & p_Delimeter
                End If
                p_Text = p_Text & list(i)
            
            End If
        Else
            'This item does not have any content. If ForceValue is true, we skip it, otherwise we add a blank element.
            If p_ForceValue <> True Then
                If p_Text <> "" Then
                    p_Text = p_Text & p_Delimeter
                End If
                p_Text = p_Text & list(i)
            End If
        End If
    Next i
    Call CountElements
End Sub

'########################################
'#
'#  Add to and remove from list
'#
'########################################

Public Function Element(ByVal index As Long) As String
    If index > CountElements Then
        Element = -1
    Else
        Dim thislist() As String
        thislist = Split(p_Text, p_Delimeter)
        Element = thislist(index)
    End If
End Function

Public Function Shift() As String
    If p_CountElements = 1 Then
        Shift = p_Text
        p_CountElements = 0
        p_Text = ""
    ElseIf p_CountElements > 1 Then
        Dim intFirstDel As Long
        intFirstDel = FirstDelimeter
        Shift = Left(p_Text, intFirstDel - 1)
        p_Text = Right(p_Text, Len(p_Text) - intFirstDel)
        Call CountElements
    End If
End Function

Public Function Pop() As String
    If p_CountElements = 1 Then
        Pop = p_Text
        p_CountElements = 0
        p_Text = ""
    ElseIf p_CountElements > 1 Then
        Dim intLastDel As Long
        intLastDel = LastDelimeter
        Pop = Mid(p_Text, intLastDel + 1, Len(p_Text))
        p_Text = Mid(p_Text, 1, intLastDel - 1)
        Call CountElements
    End If
End Function

Public Sub UnShift(ByVal content As String)
    If p_ForceValue = True And content = "" Then
        Exit Sub
    End If
    If p_ForceUnique = True And Exists(content) = True Then
        Exit Sub
    End If
    If p_Text <> "" Then
        p_Text = content & p_Delimeter & p_Text
    Else
        p_Text = content
    End If
    Call CountElements
End Sub

Public Sub Push(ByVal content As String)
    If p_ForceValue = True And content = "" Then
        Exit Sub
    End If
    If p_ForceUnique = True And Exists(content) = True Then
        Exit Sub
    End If
    If p_Text <> "" Then
        p_Text = p_Text & p_Delimeter & content
    Else
        p_Text = content
    End If
    Call CountElements
End Sub

'########################################
'#
'#  Remove()
'#
'########################################
'To use RegEx in Excel, you need to activate the reference:
'In the VBA editor, go to Tools > References > Microsoft VBScript Regular Expressions 5.5 (or higher) and check the box next to it, and then save the workbook.

Public Sub Remove(ByVal inputvalue As Variant)
    If inputvalue <> "" Then
        Dim i As Long
        If IsDigits(inputvalue) = True Then
            'treat as index
            Dim index As Long
            index = CLng(inputvalue)
            If index > -1 Then
                Dim strModifiedList As String
                For i = 0 To LastIndex
                    If i <> index Then
                        If strModifiedList <> "" Then
                            strModifiedList = strModifiedList & p_Delimeter
                        End If
                        strModifiedList = strModifiedList & Element(i)
                    End If
                Next i
                p_Text = strModifiedList
                Call CountElements
            End If
        Else
            'treat as regex
            Dim strNewData As String
            strNewData = ""
            Dim regexThis As New RegExp
            regexThis.Global = True
            regexThis.IgnoreCase = IIf(p_RemoveIgnoreCase = True, True, False)
            regexThis.Pattern = CStr(inputvalue)
            Dim arrData() As String
            arrData = Split(p_Text, ",")
            For i = 0 To UBound(arrData)
                If regexThis.Test(arrData(i)) = False Then
                    If strNewData <> "" Then
                        strNewData = strNewData & p_Delimeter
                    End If
                    strNewData = strNewData & arrData(i)
                End If
            Next i
            Set regexThis = Nothing
            p_Text = strNewData
            Call CountElements
        End If
    End If
End Sub

'########################################
'#
'#  Edit()
'#
'########################################

Public Sub Edit(ByVal index As Long, ByVal content As String)
    If index > -1 And content <> "" Then
        Dim i As Long
        Dim strModifiedList As String
        For i = 0 To LastIndex
            If i = index Then
                If strModifiedList <> "" Then
                    strModifiedList = strModifiedList & p_Delimeter
                End If
                strModifiedList = strModifiedList & content
            Else
                If strModifiedList <> "" Then
                    strModifiedList = strModifiedList & p_Delimeter
                End If
                strModifiedList = strModifiedList & Element(i)
            End If
        Next i
        p_Text = strModifiedList
        Call CountElements
    End If
End Sub

'########################################
'#
'#  Grep()
'#
'########################################
'
'To use RegEx in Excel, you need to activate the reference:
'In the VBA editor, go to Tools > References > Microsoft VBScript Regular Expressions 5.5 (or higher) and check the box next to it, and then save the workbook.

Public Function Grep(ByVal strPattern As String, Optional ByVal strReplaceWith As String) As String
    Grep = ""
    'Grep first match  objArray.Grep("\w")
    'Grep first replacements objArray.Grep("\w", "123")

    'This is where people start smirking because the class to replace arrays, uses arrays. Please keep
    'in mind that this class is to make arrays easier, not replace them totally. There are some things
    'that are much faster using Excels native arrays. This class just makes it easier. Okay, putting
    'the soap box away.
    '
    'We need to take the existing data held as a long string and store it in an array to loop through
    'as this is faster that iterating through the length of the array.
    Dim arrData As Variant
    ReDim arrData(p_CountElements - 1)
    arrData = Split(p_Text, p_Delimeter)
    '
    'Get the last index and assign it to a variable. This is much faster in a loop context than
    'recalculating it with every iteration.
    Dim longLastIndex As Long
    longLastIndex = UBound(arrData)
    '
    'Declare an array to hold each item from the original array as it is processed
    Dim strOriginalData As String
    'The Grep() requires a pattern to match.
    If strPattern <> "" Then
        'Declare variables
        Dim iElement As Long

        'Declare and prepare REGEX object
        Dim regexThis As New RegExp
        regexThis.Pattern = strPattern
        regexThis.Global = True
        regexThis.IgnoreCase = IIf(p_GrepIgnoreCase = True, True, False)
        'Loop through all elements in the array
        For iElement = 0 To longLastIndex
            'Is this array element a match for the pattern? YES
            If regexThis.Test(arrData(iElement)) = True Then
                'we process every regex match.
                If strReplaceWith <> "" Then
                    'Are we making a replacement? YES
                    'We need to add the grepped element back to the original data array
                    If Grep <> "" Then
                        Grep = Grep & p_Delimeter
                    End If
                    Grep = Grep & strReplaceWith
                    If strOriginalData <> "" Then
                        strOriginalData = strOriginalData & p_Delimeter
                    End If
                    strOriginalData = strOriginalData & arrData(iElement)
                Else
                    'We need to add the grepped element back to the original data array
                    If Grep <> "" Then
                        Grep = Grep & p_Delimeter
                    End If
                    Grep = Grep & arrData(iElement)
                    If strOriginalData <> "" Then
                        strOriginalData = strOriginalData & p_Delimeter
                    End If
                    strOriginalData = strOriginalData & arrData(iElement)
                End If
            Else
                'Is this array element a match for the pattern? NO
                'The only thing we have to do is make sure this element stays in the array.
                If strOriginalData <> "" Then
                    strOriginalData = strOriginalData & p_Delimeter
                End If
                strOriginalData = strOriginalData & arrData(iElement)
            End If
        Next iElement
        Set regexThis = Nothing
        p_Text = strOriginalData
        Call CountElements
    End If
End Function

'########################################
'#
'#  Slice()
'#
'########################################
'To use RegEx in Excel, you need to activate the reference:
'In the VBA editor, go to Tools > References > Microsoft VBScript Regular Expressions 5.5 (or higher) and check the box next to it, and then save the workbook.

Public Function Slice(ByVal inputvalue As Variant, Optional ByVal count As Long) As String
    'Removes one or more items from the array and returns them as a string list
    Slice = ""
    If IsDigits(CStr(inputvalue)) = True Then
        Dim index As Long
        index = CLng(inputvalue)
        If count = 0 Then
            count = (LastIndex - index) + 1
        End If
        If index < 0 Or count < 0 Then
            Slice = "-1"
            Exit Function
        End If
        If index > LastIndex Or (index + count) > LastIndex Then
            Slice = "-1"
            Exit Function
        End If
        Dim iElement As Long
        Dim strNewText As String
        For iElement = 0 To LastIndex
            If iElement < index Or iElement > ((index + count) - 1) Then
                If strNewText <> "" Then
                    strNewText = strNewText & p_Delimeter
                End If
                strNewText = strNewText & Element(iElement)
            Else
                If Slice <> "" Then
                    Slice = Slice & p_Delimeter
                End If
                Slice = Slice & Element(iElement)
            End If
        Next iElement
        p_Text = strNewText
        Call CountElements
    Else
        'Is a regex
        Dim strNewData As String
        strNewData = ""
        Dim regexThis As New RegExp
        regexThis.Global = True
        regexThis.IgnoreCase = IIf(p_SliceIgnoreCase = True, True, False)
        regexThis.Pattern = CStr(inputvalue)
        Dim arrData() As String
        arrData = Split(p_Text, ",")
        Dim i As Long
        For i = 0 To UBound(arrData)
            If regexThis.Test(arrData(i)) = True Then
                If Slice <> "" Then
                    Slice = Slice & p_Delimeter
                End If
                Slice = Slice & arrData(i)
            Else
                If strNewData <> "" Then
                    strNewData = strNewData & p_Delimeter
                End If
                strNewData = strNewData & arrData(i)
            End If
        Next i
        Set regexThis = Nothing
        p_Text = strNewData
        Call CountElements
    End If
End Function

'########################################
'#
'#  Splice()
'#
'########################################

Public Function Splice(ByVal index As Long, ByVal content As String) As Boolean
    'Adds to arrays, may remove
    'Index is the starting point, content is what gets inserted
    'Returns true if successful
    Splice = False
    If content = "" Or p_Text = "" Then
        Exit Function
    End If
    If index < 0 Or index > LastIndex Then
        Exit Function
    End If
    If p_ForceUnique = True Then
        Dim iUnique As Long
        Dim strUniqueText As String
        Dim arrUnique() As String
        arrUnique = Split(content, p_Delimeter)
        For iUnique = 0 To UBound(arrUnique)
            If Exists(arrUnique(iUnique)) = False Then
                If strUniqueText <> "" Then
                    strUniqueText = strUniqueText & p_Delimeter
                End If
                strUniqueText = strUniqueText & arrUnique(iUnique)
            End If
        Next iUnique
        content = strUniqueText
    End If
    If p_ForceValue = True Then
        Dim regexThis As Object
        Set regexThis = New RegExp
        regexThis.Global = True
        regexThis.Pattern = p_Delimeter & p_Delimeter & "+"
        content = regexThis.Replace(content, p_Delimeter)
        'Remove leading delimeter
        If Mid(content, 1, 1) = p_Delimeter Then
            content = Mid(content, 2, Len(content))
        End If
        'Remove trailing delimeter
        If Mid(content, Len(content), 1) = p_Delimeter Then
            content = Mid(content, 1, Len(content) - 1)
        End If
    End If
    If index = 0 Then
        'insert at start
        p_Text = content & p_Delimeter & p_Text
        Splice = True
        Call CountElements
    Else
        'index is > 0
        Dim iElement As Long
        Dim strNewText As String
        Dim arrElements() As String
        arrElements = Split(p_Text, p_Delimeter)
        For iElement = 0 To UBound(arrElements)
            If strNewText <> "" Then
                strNewText = strNewText & p_Delimeter
            End If
            If iElement = index Then
                strNewText = strNewText & content & p_Delimeter
            End If
            strNewText = strNewText & arrElements(iElement)
        Next iElement
        If p_Text <> strNewText Then
            Splice = True
            p_Text = strNewText
            Call CountElements
        End If
    End If
End Function

'########################################
'#
'#  Join()
'#
'########################################

Public Sub Join(ByRef objOtherArray As Object)
    Dim strNewData As String
    strNewData = objOtherArray.raw
    If p_ForceUnique = True Then
        Dim iUnique As Long
        Dim strUniqueText As String
        Dim arrUnique() As String
        arrUnique = Split(strNewData, p_Delimeter)
        For iUnique = 0 To UBound(arrUnique)
            If Exists(arrUnique(iUnique)) = False Then
                If strUniqueText <> "" Then
                    strUniqueText = strUniqueText & p_Delimeter
                End If
                strUniqueText = strUniqueText & arrUnique(iUnique)
            End If
        Next iUnique
        strNewData = strUniqueText
    End If
    If p_ForceValue = True Then
        Dim regexThis As Object
        Set regexThis = New RegExp
        regexThis.Global = True
        regexThis.Pattern = p_Delimeter & p_Delimeter & "+"
        strNewData = regexThis.Replace(strNewData, p_Delimeter)
        'Remove leading delimeter
        If Mid(strNewData, 1, 1) = p_Delimeter Then
            strNewData = Mid(strNewData, 2, Len(strNewData))
        End If
        'Remove trailing delimeter
        If Mid(strNewData, Len(strNewData), 1) = p_Delimeter Then
            strNewData = Mid(strNewData, 1, Len(strNewData) - 1)
        End If
    End If
    If Len(strNewData) > 0 Then
        p_Text = p_Text & p_Delimeter & strNewData
    End If
End Sub

'########################################
'#
'#  Exists()
'#
'########################################

Public Function Exists(ByVal content As String) As Boolean
    Exists = False
    If content <> "" And p_Text <> "" Then
        If InStr(p_Text, p_Delimeter) > 0 Then
            'There is more than one element in the array
            'Declare variables
            Dim iChars As Long
            Dim iLastDelimeter As Long
            iLastDelimeter = 0
            'Loop through length of string list
            For iChars = 1 To Len(p_Text)
                If iChars = Len(p_Text) Then
                    'If this is the last character, it will not be a delimeter but we have to test it.
                    If Mid(p_Text, iLastDelimeter + 1, Len(p_Text)) = content Then
                        Exists = True
                        Exit Function
                    End If
                ElseIf Mid(p_Text, iChars, 1) = p_Delimeter Then
                    If iLastDelimeter > 0 Then
                        If Mid(p_Text, iLastDelimeter + 1, (iChars - iLastDelimeter) - 1) = content Then
                            Exists = True
                            Exit Function
                        End If
                    Else
                        'if iLastDelimeter is not set, we are looking at the first element
                        If Mid(p_Text, 1, iChars - 1) = content Then
                            Exists = True
                            Exit Function
                        End If
                    End If
                    iLastDelimeter = iChars
                End If
            Next iChars
        Else
            'There is only one element in the array.
            If p_Text = content Then
                Exists = True
            End If
        End If
    End If
End Function

'########################################
'#
'#  IndexOf()
'#
'########################################

Public Function IndexOf(ByVal content As String) As Long
    'Default value to -1, which means that the target was not found
    IndexOf = -1
    'Make sure there is a target to look for, and content in which to look for the target
    If content <> "" And p_Text <> "" Then
        'Declare & default variables
        Dim iElement As Long
        Dim iLastDelimeter As Long
        iLastDelimeter = 0
        Dim iChars As Long
        'Loop through all elements of the array in the string form
        'We are going to look for delimeters as the space between delimeters
        'is array elements. We then match the element against the target to
        'see if it is a match. If so, that (delimeter iteration - 1) is the index position
        'in the concept of the array
        For iChars = 1 To Len(p_Text)
            If iChars = Len(p_Text) Then
                'If chars equals the length of the string data we are looking in, then
                'we are at the last element. As the string data does NOT end with a delimeter,
                'we do stuff here to pretend we found the last delimeter.
                'REMEMBER: The index is (element - 1) so that is why we increment AFTER we match.
                If Mid(p_Text, iLastDelimeter + 1, Len(p_Text)) = content Then
                    'We look at iLastDelimeter + 1 because that is the first letter of this
                    'array elelment, where iLastDelimete is a delimiter and not part of the element.
                    IndexOf = iElement
                End If
                iElement = iElement + 1
            ElseIf Mid(p_Text, iChars, 1) = p_Delimeter Then
                If iLastDelimeter > 0 Then
                    'We need to look at element residing between (last delimeter position + 1)
                    'and (this delimeter position - 1) to find the actual array element.
                    'REMEMBER: The index is (element - 1) so that is why we increment AFTER we match.
                    If Mid(p_Text, iLastDelimeter + 1, iChars - (iLastDelimeter + 1)) = content Then
                        IndexOf = iElement
                    End If
                Else
                    'If the delimeter iterator is zero, then this is the first delimeter we have found.
                    'That means that the characters to the left are part of the first element.
                    'REMEMBER: The index is (element - 1) so that is why we increment AFTER we match.
                    iElement = 0
                    If Mid(p_Text, 1, iChars - 1) = content Then
                        IndexOf = iElement
                    End If
                End If
                iElement = iElement + 1
                iLastDelimeter = iChars
            End If
        Next iChars
    End If
End Function

'########################################
'#
'#  Sorting
'#
'########################################

Property Get SortAsNumbers() As Boolean
    SortAsNumbers = p_SortAsNumbers
End Property
Property Let SortAsNumbers(ByVal boolValue As Boolean)
    If boolValue = True Then
        p_SortAsNumbers = True
        p_SortAsText = False
    Else
        p_SortAsNumbers = False
        p_SortAsText = True
    End If
End Property

Property Get SortAsText() As Boolean
    SortAsText = p_SortAsText
End Property
Property Let SortAsText(ByVal boolValue As Boolean)
    If boolValue = True Then
        p_SortAsNumbers = False
        p_SortAsText = True
    Else
        p_SortAsNumbers = True
        p_SortAsText = False
    End If
End Property

Public Sub Reverse()
    If p_Text <> "" Then
        Dim strReversed As String
        'Declare & default variables
        Dim iChars As Long
        For iChars = Len(p_Text) To 1 Step -1
            If iChars = 1 Then
                If strReversed <> "" Then
                    strReversed = strReversed & p_Delimeter
                End If
                strReversed = strReversed & p_Text
            ElseIf Mid(p_Text, iChars, 1) = p_Delimeter Then
                'If we encounter a delimeter, we add one and take all the characters to the end as an element.
                'Then we chop off the delimeter. Then we continue.
                If strReversed <> "" Then
                    strReversed = strReversed & p_Delimeter
                End If
                strReversed = strReversed & Mid(p_Text, iChars + 1, Len(p_Text))
                p_Text = Mid(p_Text, 1, iChars - 1)
            End If
        Next iChars
        p_Text = strReversed
    End If
End Sub

Public Sub SortAscending()
    If p_Text <> "" Then
        Dim thislist() As String
        ReDim thislist(CountElements)
        thislist = Split(p_Text, p_Delimeter)
        Dim First As Integer, Last As Long
        Dim i As Long, j As Long
        Dim temp As String
        If p_SortAsText = True Then
            First = LBound(thislist)
            Last = UBound(thislist)
            For i = First To Last - 1
                For j = i + 1 To Last
                    If thislist(i) > thislist(j) Then
                        temp = thislist(j)
                        thislist(j) = thislist(i)
                        thislist(i) = temp
                    End If
                Next j
            Next i
        Else
            First = LBound(thislist)
            Last = UBound(thislist)
            For i = First To Last - 1
                For j = i + 1 To Last
                    If CLng(thislist(i)) > CLng(thislist(j)) Then
                        temp = thislist(j)
                        thislist(j) = thislist(i)
                        thislist(i) = temp
                    End If
                Next j
            Next i
        End If

        First = LBound(thislist)
        Last = UBound(thislist)
        p_Text = ""
        For i = First To Last
            If p_Text <> "" Then
                p_Text = p_Text & p_Delimeter
            End If

            p_Text = p_Text & thislist(i)
        Next i
        Call CountElements
    End If
End Sub

Public Sub SortTextAscending(list() As Variant)
    'Taken from power programming with VBA
     Dim First As Integer, Last As Long
     Dim i As Long, j As Long
     Dim temp As String
     First = LBound(list)
     Last = UBound(list)
     For i = First To Last - 1
         For j = i + 1 To Last
             If list(i) > list(j) Then
                 temp = list(j)
                 list(j) = list(i)
                 list(i) = temp
             End If
         Next j
     Next i
End Sub

Public Sub SortTextDescending(list() As Variant)
    'Taken from power programming with VBA
     Dim First As Integer, Last As Long
     Dim i As Long, j As Long
     Dim temp As String
     First = LBound(list)
     Last = UBound(list)
     For i = First To Last - 1
         For j = i + 1 To Last
             If list(i) < list(j) Then
                 temp = list(j)
                 list(j) = list(i)
                 list(i) = temp
             End If
         Next j
     Next i
End Sub

Public Sub SortNumberAscending(list() As Variant)
    'Taken from power programming with VBA
     Dim First As Integer, Last As Long
     Dim i As Long, j As Long
     Dim temp As String
     First = LBound(list)
     Last = UBound(list)
     For i = First To Last - 1
         For j = i + 1 To Last
             If CLng(list(i)) < CLng(list(j)) Then
                 temp = list(j)
                 list(j) = list(i)
                 list(i) = temp
             End If
         Next j
     Next i
End Sub

Public Sub SortNumberDescending(list() As Variant)
    'Taken from power programming with VBA
     Dim First As Integer, Last As Long
     Dim i As Long, j As Long
     Dim temp As String
     First = LBound(list)
     Last = UBound(list)
     For i = First To Last - 1
         For j = i + 1 To Last
             If CLng(list(i)) > CLng(list(j)) Then
                 temp = list(j)
                 list(j) = list(i)
                 list(i) = temp
             End If
         Next j
     Next i
End Sub


'########################################
'#
'#  PRIVATE METHODS
'#
'########################################

Private Function FirstDelimeter() As Long
    FirstDelimeter = -1
    If Len(p_Text) > 0 Then
        Dim i As Long
        For i = 1 To Len(p_Text)
            If Mid(p_Text, i, 1) = p_Delimeter Then
                FirstDelimeter = i
                Exit For
            End If
        Next i
    End If
End Function

Private Function LastDelimeter() As Long
    LastDelimeter = -1
    If Len(p_Text) > 0 Then
        Dim i As Long
        For i = Len(p_Text) To 1 Step -1
            If Mid(p_Text, i, 1) = p_Delimeter Then
                LastDelimeter = i
                Exit For
            End If
        Next i
    End If
End Function

Private Function IsDigits(ByVal strTarget As String) As Boolean
    IsDigits = False
    Dim regexThis As New RegExp
    regexThis.Pattern = "^-{0,1}[0-9]+$"
    IsDigits = regexThis.Test(strTarget)
End Function

'########################################
'#
'#  DEBUGGING & TESTING
'#
'########################################

Public Function raw() As String
    raw = p_Text
End Function



