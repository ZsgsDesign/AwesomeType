VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CStringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mvarStringArray() As String
Private mvarArrayItems As Long
 
Public Sub Append(ByVal newStr As String)
    ReDim Preserve mvarStringArray(mvarArrayItems) As String
    mvarStringArray(mvarArrayItems) = newStr
    mvarArrayItems = mvarArrayItems + 1
End Sub
  
Public Property Get toString() As String
    If mvarArrayItems > 0 Then toString = Join(mvarStringArray, "")
End Property
 
Public Sub Reset()
    mvarArrayItems = 0
    Erase mvarStringArray
End Sub

Private Sub Class_Initialize()
    If mvarArrayItems > 0 Then Reset
End Sub


Private Sub Class_Terminate()
  Reset
End Sub


