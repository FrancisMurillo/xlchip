Attribute VB_Name = "VaseAssert"
'# This global variable determines if an assert method failed or passed
Private gAssertion As Boolean
Private gFirstFailed As String
Private gFirstFailedMessage As String

Public Property Get TestResult() As Boolean
    TestResult = gAssertion
End Property
Public Property Get FirstFailedTestMethod() As String
    FirstFailedTestMethod = gFirstFailed
End Property
Public Property Get FirstFailedTestMessage() As String
    FirstFailedTestMessage = gFirstFailedMessage
End Property

'=======================
'--- Assertion Tools ---
'=======================

'# Sets the Assertion globals for use
Public Sub InitAssert()
    gAssertion = True
    gFirstFailed = ""
    gFirstFailedMessage = ""
End Sub


'# Base Assert Method
Public Sub Assert_(Cond As Boolean, _
        Optional Message As String = "", _
        Optional AssertName As String = "Assert") ' Name to avoid Debug.Assert conflict or confusion
    gAssertion = gAssertion And Cond ' Update assertion variable
    If Not Cond And gFirstFailed = "" Then ' Log the first fail condition for logging
        gFirstFailed = AssertName
    End If
End Sub

'# Assert if condition is true
Public Sub AssertTrue(Cond As Boolean, Optional Message As String = "")
    Assert_ Cond, Message:=Message, AssertName:="AssertTrue"
End Sub

'# Assert condition is false
Public Sub AssertFalse(Cond As Boolean, Optional Message As String = "")
    Assert_ Not Cond, Message:=Message, AssertName:="AssertFalse"
End Sub

'# Assert two variables are equal
Public Sub AssertEqual(LeftVal As Variant, RightVal As Variant, Optional Message As String = "")
    Assert_ Equal_(LeftVal, RightVal), Message:=Message, AssertName:="AssertEqual"
End Sub

'# Assert left variable is like the right variable
Public Sub AssertLike(LeftVal As Variant, RightVal As Variant, Optional Message As String = "")
    Assert_ Like_(LeftVal, RightVal), Message:=Message, AssertName:="AssertLike"
End Sub

'# Assert greater than
Public Sub AssertGreaterThan(LeftVal As Variant, RightVal As Variant, Optional Message As String = "")
    Assert_ GreaterThan_(LeftVal, RightVal), Message:=Message, AssertName:="AssertGreaterThan"
End Sub

'# Assert greater than or equal
Public Sub AssertGreaterThanOrEqual(LeftVal As Variant, RightVal As Variant, Optional Message As String = "")
    Assert_ GreaterThanOrEqual_(LeftVal, RightVal), Message:=Message, AssertName:="AssertGreaterThanOrEqual"
End Sub

'# Assert less than
Public Sub AssertLess(LeftVal As Variant, RightVal As Variant, Optional Message As String = "")
    Assert_ LessThan_(LeftVal, RightVal), Message:=Message, AssertName:="AssertLess"
End Sub

'# Assert less than or equal
Public Sub AssertLessThanOrEqual(LeftVal As Variant, RightVal As Variant, Optional Message As String = "")
    Assert_ LessThanOrEqual_(LeftVal, RightVal), Message:=Message, AssertName:="AssertLessThanOrEqual"
End Sub

Private Function Equal_(LeftVal As Variant, RightVal As Variant) As Boolean
    Dim PreClear As Boolean
    PreClear = (Err.Number = 0) ' Save the default error state
On Error Resume Next
    Equal_ = False
    Equal_ = (LeftVal = RightVal) ' Mutates the error state if an error occurs here
    If PreClear Then Err.Clear ' If there was an previous error, do not clear it
End Function

Private Function Like_(LeftVal As Variant, RightVal As Variant) As Boolean
    Dim PreClear As Boolean
    PreClear = (Err.Number = 0) ' Save the default error state
On Error Resume Next
    Like_ = False
    Like_ = (LeftVal Like RightVal) ' Mutates the error state if an error occurs here
    If PreClear Then Err.Clear ' If there was an previous error, do not clear it
End Function

Private Function LessThan_(LeftVal As Variant, RightVal As Variant) As Boolean
    Dim PreClear As Boolean
    PreClear = (Err.Number = 0) ' Save the default error state
On Error Resume Next
    LessThan_ = False
    LessThan_ = (LeftVal < RightVal) ' Mutates the error state if an error occurs here
    If PreClear Then Err.Clear ' If there was an previous error, do not clear it
End Function

Private Function LessThanOrEqual_(LeftVal As Variant, RightVal As Variant) As Boolean
    Dim PreClear As Boolean
    PreClear = (Err.Number = 0) ' Save the default error state
On Error Resume Next
    LessThanOrEqual_ = False
    LessThanOrEqual_ = (LeftVal <= RightVal) ' Mutates the error state if an error occurs here
    If PreClear Then Err.Clear ' If there was an previous error, do not clear it
End Function

Private Function GreaterThan_(LeftVal As Variant, RightVal As Variant) As Boolean
    Dim PreClear As Boolean
    PreClear = (Err.Number = 0) ' Save the default error state
On Error Resume Next
    GreaterThan_ = False
    GreaterThan_ = (LeftVal > RightVal) ' Mutates the error state if an error occurs here
    If PreClear Then Err.Clear ' If there was an previous error, do not clear it
End Function

Private Function GreaterThanOrEqual_(LeftVal As Variant, RightVal As Variant) As Boolean
    Dim PreClear As Boolean
    PreClear = (Err.Number = 0) ' Save the default error state
On Error Resume Next
    GreaterThanOrEqual_ = False
    GreaterThanOrEqual_ = (LeftVal >= RightVal) ' Mutates the error state if an error occurs here
    If PreClear Then Err.Clear ' If there was an previous error, do not clear it
End Function


'# Assert not equal
Public Sub AssertNotEqual(LeftVal As Variant, RightVal As Variant, Optional Message As String = "")
    Assert_ Not Equal_(LeftVal, RightVal), Message:=Message, AssertName:="AssertNotEqual"
End Sub

'# Assert something is inside an array
Public Sub AssertInArray(Elem As Variant, Arr As Variant, Optional Message As String = "")
    Assert_ VaseLib.InArray(Elem, Arr), Message:=Message, AssertName:="AssertInArray"
End Sub

'# Assert array is of the correct size
Public Sub AssertArraySize(Size As Long, Arr As Variant, Optional Message As String = "")
    Assert_ Equal_(Size, UBound(Arr) + 1), Message:=Message, AssertName:="AssertArraySize"
End Sub

'# Assert array is of the correct size
Public Sub AssertEmptyArray(Arr As Variant, Optional Message As String = "")
    Assert_ Equal_(-1, UBound(Arr)), Message:=Message, AssertName:="AssertEmptyArray"
End Sub

'# Assert array elements are equal
Public Sub AssertEqualArrays(LeftArr As Variant, RightArr As Variant, Optional Message As String = "")
    Dim Tuple As Variant, ArrSize As Long
    ArrSize = UBound(LeftArr) + 1
    
    AssertArraySize UBound(LeftArr) + 1, RightArr, Message:=Message
    For Each Tuple In VaseLib.Zip(LeftArr, RightArr)
        AssertEqual Tuple(0), Tuple(1), Message:=Message
    Next
End Sub


