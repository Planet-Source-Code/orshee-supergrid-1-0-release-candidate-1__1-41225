Attribute VB_Name = "Globals"
Option Explicit

Public Enum enuCursorType
    ctOpenForwardOnly = 0
    ctOpenKeyset = 1
    ctOpenDynamic = 2
    ctdOpenStatic = 3
End Enum

Public Enum enuCursorLocation
    clUseServer = 2
    clUseClient = 3
End Enum

Public Enum enuLockType
    ltdLockReadOnly = 1
    ltLockPessimistic = 2
    ltLockOptimistic = 3
    ltLockBatchOptimistic = 4
End Enum

Public Const vbButtonHighlight      As Long = vb3DHighlight
Public Const vbButtonLightShadow    As Long = vb3DShadow

Public Function CheckNull(Value As String) As String
    If IsNull(Value) = True Then
        CheckNull = ""
    Else
        CheckNull = CStr(Value)
    End If
End Function
