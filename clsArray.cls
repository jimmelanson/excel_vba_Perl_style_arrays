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
' REQUIRES MSExcel 2016 or higher
'
'Trying to do anything comples with VBA arrays is like banging your head on a wall
'but without that satisfying feeling when you stop. Because of this, I created this
'class to do the nifty stuff I can do with arrays in Perl (mostly).
'
'This object takes a string of values and treats them as an array. If you using numbers
'that are more complicated than a double, then this object probably won't work for you.
'
'
'Declaring the object (early binding always recommended)
'    Dim objArray As New clsArray
'
'
'The default delimeter for this object is the comma. You can specify a different delimeter
'but you must do so prior to passing the string of data to the object.
'
'    objArray.Delimeter = ":"
'
'
'You can force the object to only accept valid values. That is, you can force the object
'to reject empty values.
'
'    objArray.ForceValue = True
'
'
'You can force the object to only accept unique values.
'
'   objArray.ForceUnique = True
'
'
'#CONSTRUCTING THE OBJECT LIST
'If you do happen to have your data in an array, it needs to be a Variant array. You can
'pass it to the object like this:
'
'    objArray.SetArray list:=arrMyData
'
'
'I usually grab data from a worksheet and concatenate it into a string to pass to the object:
'
'    Dim strMyStuff As String
'    Dim y As Long
'    For y = 0 To 95
'        If strMyStuff <> "" Then
'            strMyStuff = strMyStuff & ":"
'        End If
'        strMyStuff = strMyStuff & Cells(y, 1).Value
'    Next y
'    objArray.SetArrayFromString content:=strMyStuff
'
'
'Now that your data is loaded into the object, we can do all sorts of nifty things:
'
'#ADD TO EXISTING LIST:
'To add one more element to the end of the list, use "Push"
'
'    objArray.Push content:="DooDad"
'
'If you want to add an item to the front of the list, use "UnShift"
'    objArray.UnShift content:="WhizBang"
'
'
'#REMOVE FROM EXISTING LIST
'If you want to remove an item from the end of the list, use "Pop".
'The value removed is returned as a String type
'
'    Dim strItem As String
'    strItem = objArray.Pop
'
'
'If you want to remove the first item from the list, use "Shift".
'The value removed is returned as a String type
'
'    Dim strItem As String
'    strItem = objArray.Shift
'
'
'#REMOVE TEXT MATCH OR BY INDEX
'You can remove an element from the list by matching by
'indicating the index. The removed value is not returned.
'
'Remove by Index (0-based)
'   objArray.Remove index:=n
'
'Remove by text, case-sensitive
'    oArray.Remove index:=oArray.IndexOf("whizbang")
'
'NOTE: if you do need to get the value of the element at the
'index before removing the element at the index, handle it
'this way:
'   dim strValue as string
'   strValue = objArray.Element(n)
'   ojbArray.Remove(n)
'
'
'#READ FROM EXISTING LIST WITHOUT REMOVING
'If you want the contents of an element, you can access the list elements
'directly. Remember that this ojbect acts like a 0-based list. Read it like this:
'
'    Dim strItem As String
'    strItem = objArray.Element(9) '10th item in the list
'
'You can find out how many items are in the list with:
'
'    Dim intCountItems As Integer
'    intCountItems = objArray.CountElements
'
'There is a special counter used for Do While loops:
'
'    Dim intCount As Integer
'    intCount = objArray.CountRemaining
'    Do While objArray.CountRemaining > 0
'       strValue = objArray.Pop
'    Loop
'
'
'Get the 0-based index of an element in the list, case-sensitive
'
'   Dim intIndex as Integer
'   intIndex = oArray.IndexOf content:="whizbang"
'
'
'#SORTING THE LIST
'If you want to sort the list (ascending):
'
'    objArray.SortAscending
'
'
'If you want to sort the list (descending), then first sort it
'ascending and then reverse sort the list:
'
'    objArray.SortAscending
'    objArray.Reverse
'
'
'#DEBUGGING
'If you want to see what values the object is holding while you are
'developing your procedure, use the raw() method.
'
'   Debug.Print oArray.raw
'
'
'#DON'T FORGET TO ALWAYS DESTROY THE OBJECT AT THE END OF THE PROCEDURE
'
'    Set objArray = Nothing

Private p_Text As String
Private p_Delimeter As String
Private p_CountRemaining As Integer
Private p_ForceValue As Boolean
Private p_ForceUnique As Boolean
Private p_SortAsNumbers As Boolean
Private p_SortAsText As Boolean

Private Sub Class_Initialize()
    p_CountRemaining = 0
    p_Delimeter = ","
    p_Text = ""
    p_SortAsNumbers = False
    p_SortAsText = True
    p_ForceValue = False
    p_ForceUnique = False
End Sub

Private Sub Class_Terminate()
    p_CountRemaining = 0
    p_Delimeter = ""
    p_Text = ""
End Sub

Private Sub Reset()
    p_CountRemaining = 0
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

Public Function LastIndex() As Integer
    LastIndex = (CountElements - 1)
End Function

Public Function CountRemaining() As Integer
    CountRemaining = CountElements
End Function

Public Function CountElements() As Integer
    If p_Text = p_Delimeter Then
        p_Text = ""
        p_CountRemaining = 0
        CountElements = 0
    Else
        Dim arrChrs() As String
        arrChrs = Split(p_Text, p_Delimeter)
        Dim iChrs As Integer
        iChrs = UBound(arrChrs, 1)
        If iChrs = -1 Then
            If Len(p_Text) > 0 Then
                CountElements = 1
            Else
                CountElements = 0
            End If
        Else
            CountElements = iChrs + 1
        End If
    End If
End Function

'########################################
'#
'#  Populate Object Array w/List
'#
'########################################

Public Sub SetArrayFromString(ByVal content As String)
    Reset
    If content <> "" Then
        p_Text = content
        p_CountRemaining = CountElements
        'Check to see if we are forcing anything
        If p_ForceValue = True Or p_ForceUnique = True Then
            'We will be modifying the list as we move along if necessary
            'This will be the new list we restore to the p_Text property
            Dim strNewText As String
            'Declare variable
            Dim strTest As String
            Dim boolIgnore As Boolean
            'Execute the loop
            Do While CountRemaining > 0
                boolIgnore = False
                'Grab the first element remaining in the list
                strTest = Shift
                'Check to see if we are forcing both content and uniqueness
                If p_ForceValue = True Then
                    'We are forcing values
                    If strTest = "" Then
                        boolIgnore = True
                    End If
                End If
                If p_ForceUnique = True Then
                    'We are forcing uniqueness
                    If exists(strTest) > 0 Then
                        boolIgnore = True
                    End If
                End If
                If boolIgnore = False Then
                    If strNewText <> "" Then
                        strNewText = strNewText & p_Delimeter
                    End If
                    strNewText = strNewText & strTest
                End If
            Loop
            p_Text = strNewText
            p_CountRemaining = CountRemaining
        End If
    End If
End Sub

Public Sub SetArray(list() As Variant)
    Reset
    Dim i As Integer
    For i = LBound(list) To UBound(list)
        If (list(i) = "" And p_ForceValue = False) Or list(i) <> "" Then
            If p_ForceUnique = True And exists(list(i)) > 0 Then
                'We actually want to do NOTHING if this condition is true
            Else
                If p_Text <> "" Then
                    p_Text = p_Text & p_Delimeter
                End If
                p_Text = p_Text & list(i)
            End If
        End If
    Next i
    p_CountRemaining = CountElements
End Sub

'########################################
'#
'#  Add to and remove from list
'#
'########################################

Public Function Element(ByVal index As Integer) As String
    If index > CountElements Then
        Element = -1
    Else
        Dim thislist() As String
        thislist = Split(p_Text, p_Delimeter)
        Element = thislist(index)
    End If
End Function

Public Function Shift() As String
    p_CountRemaining = CountElements
    If p_CountRemaining = 1 Then
        Shift = p_Text
        p_CountRemaining = 0
        p_Text = ""
    ElseIf p_CountRemaining > 1 Then
        Dim intFirstDel As Integer
        intFirstDel = FirstDelimeter
        Shift = Left(p_Text, intFirstDel - 1)
        p_Text = Right(p_Text, Len(p_Text) - intFirstDel)
        p_CountRemaining = CountElements
    End If
End Function

Public Function Pop() As String
    p_CountRemaining = CountElements
    If p_CountRemaining = 1 Then
        Pop = p_Text
        p_CountRemaining = 0
        p_Text = ""
    ElseIf p_CountRemaining > 1 Then
        Dim intLastDel As Integer
        intLastDel = LastDelimeter
        Pop = Mid(p_Text, intLastDel + 1, Len(p_Text))
        p_Text = Mid(p_Text, 1, intLastDel - 1)
        p_CountRemaining = CountElements
    End If
End Function

Public Sub UnShift(ByVal content As String)
    If p_ForceValue = True And content = "" Then
        Exit Sub
    End If
    If p_ForceUnique = True And exists(content) > 0 Then
        Exit Sub
    End If
    If p_Text <> "" Then
        p_Text = content & p_Delimeter & p_Text
    Else
        p_Text = content
    End If
    p_CountRemaining = CountElements
End Sub

Public Sub Push(ByVal content As String)
    If p_ForceValue = True And content = "" Then
        Exit Sub
    End If
    If p_ForceUnique = True And exists(content) > 0 Then
        Exit Sub
    End If
    If p_Text <> "" Then
        p_Text = p_Text & p_Delimeter & content
    Else
        p_Text = content
    End If
    p_CountRemaining = CountElements
End Sub


Public Sub Remove(ByVal index As Integer)
    If index > -1 Then
        'We will be modifying the list as we move along if necessary
        'This will be the new list we restore to the p_Text property
        Dim strNewText As String
        'Declare variable
        Dim strTest As String
        'Declare index counter
        Dim intIndex As Integer
        intIndex = -1
        'Execute the loop
        Do While CountRemaining > 0
            'Grab the first element remaining in the list
            strTest = Shift
            'Upcount the index counter
            intIndex = intIndex + 1
            If intIndex <> index Then
                If strNewText <> "" Then
                    strNewText = strNewText & p_Delimeter
                End If
                strNewText = strNewText & strTest
            End If
        Loop
        p_Text = strNewText
        p_CountRemaining = CountRemaining
    End If
End Sub

Public Sub Splice(ByVal content As String, ByVal index As Integer)
    If p_ForceValue = True And content = "" Then
        Exit Sub
    End If
    If p_ForceUnique = True And exists(content) > 0 Then
        Exit Sub
    End If
    p_CountRemaining = CountElements
    If index >= 0 And index <= p_CountRemaining Then
        If p_Text <> "" Then
            Dim strLeft As String
            Dim intIterator As Integer
            intIterator = 0
            Do Until intIterator = index
                If strLeft <> "" Then
                    strLeft = strLeft & p_Delimeter
                End If
                strLeft = strLeft & Shift
                intIterator = intIterator + 1
            Loop
            strLeft = strLeft & p_Delimeter & content & p_Delimeter & p_Text
            p_Text = strLeft
            p_CountRemaining = CountRemaining
        End If
    End If
End Sub

Public Function exists(ByVal content As String) As Integer
    exists = 0
    If content <> "" And p_Text <> "" Then
        Dim strOriginal As String
        strOriginal = p_Text
        Do While CountRemaining > 0
            If content = Pop Then
                exists = exists + 1
            End If
        Loop
        p_Text = strOriginal
        p_CountRemaining = CountRemaining
    End If
End Function

Public Function IndexOf(ByVal content As String) As Integer
    p_CountRemaining = CountElements
    Dim intReturnIndex As Integer
    intReturnIndex = -1
    If content <> "" Then
        If p_Text <> "" Then
            Dim strOriginal As String
            strOriginal = p_Text
            Dim intIterator As Integer
            intIterator = -1
            Do While CountRemaining > 0
                intIterator = intIterator + 1
                If content = Shift Then
                    intReturnIndex = intIterator
                    Exit Do
                End If
            Loop
            IndexOf = intReturnIndex
            p_Text = strOriginal
            p_CountRemaining = CountRemaining
        End If
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
        Do While CountRemaining > 0
            If strReversed <> "" Then
                strReversed = strReversed & p_Delimeter
            End If
            strReversed = strReversed & Pop
        Loop
        p_Text = strReversed
        p_CountRemaining = CountRemaining
    End If
End Sub

Public Sub SortAscending()
    If p_Text <> "" Then
        Dim thislist() As String
        ReDim thislist(CountRemaining)
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
        p_Text = Join(thislist, p_Delimeter)
        p_CountRemaining = CountRemaining
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

Private Function FirstDelimeter() As Integer
    FirstDelimeter = -1
    If Len(p_Text) > 0 Then
        Dim i As Integer
        For i = 1 To Len(p_Text)
            If Mid(p_Text, i, 1) = p_Delimeter Then
                FirstDelimeter = i
                Exit For
            End If
        Next i
    End If
End Function

Private Function LastDelimeter() As Integer
    LastDelimeter = -1
    If Len(p_Text) > 0 Then
        Dim i As Integer
        For i = Len(p_Text) To 1 Step -1
            If Mid(p_Text, i, 1) = p_Delimeter Then
                LastDelimeter = i
                Exit For
            End If
        Next i
    End If
End Function


'########################################
'#
'#  DEBUGGING & TESTING
'#
'########################################

Public Function raw() As String
    raw = p_Text
End Function



