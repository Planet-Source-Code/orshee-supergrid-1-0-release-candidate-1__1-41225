VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SuperCell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum enuCellType
    ctStandard = 0
    ctFixed = 1
End Enum

Public Enum enuCellAlignment
    caLeft = &H0
    caCenter = &H1
    caRight = &H2
End Enum

Private mvarColID           As Long
Private mvarRowID           As Long
Private mvarText            As String
Private mvarData            As Object
Private mvarForeColor       As Long
Private mvarBackColor       As Long
Private mvarAlignment       As Long
Private mvarFontBold        As Boolean
Private mvarFontItalic      As Boolean
Private mvarHasFocus        As Boolean
Private mvarSelected        As Boolean
Private mobjPicture         As StdPicture

Public Property Let ColID(Value As Long)
    mvarColID = Value
End Property

Public Property Get ColID() As Long
    ColID = mvarColID
End Property

Public Property Let RowID(Value As Long)
    mvarRowID = Value
End Property

Public Property Get RowID() As Long
    RowID = mvarRowID
End Property

Public Property Let Text(Value As String)
    mvarText = Value
End Property

Public Property Get Text() As String
    Text = mvarText
End Property

Public Property Set Data(Value As Object)
    Set mvarData = Value
End Property

Public Property Get Data() As Object
    Set Data = mvarData
End Property

Public Property Let ForeColor(Value As Long)
    mvarForeColor = Value
End Property

Public Property Get ForeColor() As Long
    ForeColor = mvarForeColor
End Property

Public Property Let BackColor(Value As Long)
    mvarBackColor = Value
End Property

Public Property Get BackColor() As Long
    BackColor = mvarBackColor
End Property

Public Property Let Alignment(Value As enuCellAlignment)
    mvarAlignment = Value
End Property

Public Property Get Alignment() As enuCellAlignment
    Alignment = mvarAlignment
End Property

Public Property Let FontBold(Value As Boolean)
    mvarFontBold = Value
End Property

Public Property Get FontBold() As Boolean
    FontBold = mvarFontBold
End Property

Public Property Let FontItalic(Value As Boolean)
    mvarFontItalic = Value
End Property

Public Property Get FontItalic() As Boolean
    FontItalic = mvarFontItalic
End Property

Public Property Let HasFocus(Value As Boolean)
    mvarHasFocus = Value
End Property

Public Property Get HasFocus() As Boolean
    HasFocus = mvarHasFocus
End Property

Public Property Let Selected(Value As Boolean)
    mvarSelected = Value
End Property

Public Property Get Selected() As Boolean
    Selected = mvarSelected
End Property

Public Property Set Picture(Value As StdPicture)
    Set mobjPicture = Value
End Property

Public Property Get Picture() As StdPicture
    Set Picture = mobjPicture
End Property


Private Sub Class_Initialize()
    mvarText = ""
    Set mvarData = Nothing
    
    mvarForeColor = -1
    mvarBackColor = -1
    mvarFontBold = False
    mvarFontItalic = False
    
    mvarHasFocus = False
    mvarSelected = False
End Sub

Private Sub Class_Terminate()
    Set mvarData = Nothing
End Sub
