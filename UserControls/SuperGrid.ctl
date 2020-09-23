VERSION 5.00
Begin VB.UserControl SuperGrid 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "SuperGrid.ctx":0000
   Begin VB.TextBox txtTextEdit 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   180
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "SuperGrid.ctx":0312
      Top             =   3120
      Visible         =   0   'False
      Width           =   1665
   End
End
Attribute VB_Name = "SuperGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**************************************************************************
'* SuperGrid 1.0 (Release Candidate 1)
'* Copyright (C) 2002, ORSHEE
'* Beta testers :
'*    michael doering
'*    hrvoje komljenovic
'**************************************************************************

Option Explicit
Option Base 1

'Implement the subclassing interface
'It gives us two events 'Before' and 'After' window message is recieved
'Here it is used mainly because of scrollbars handling messages
Implements iSuperSubClasser

'***[Types]*********************************************************************
'This type in combination with SuperCell class is used for main Cell Collection
Private Type stcRow
    Cols() As New SuperCell
End Type

'***[Enumerations]***************************************************************
Public Enum enuBorderStyle
    bsNone = 0
    bsFixedSingle = 1
End Enum

Public Enum enuScrollBars
    sbNone = 0
    sbHorizontal = 1
    sbVertical = 2
    sbBoth = 3
End Enum

Public Enum enuSelectionMode
    smCell = 0
    smRow = 1
End Enum

Public Enum enuCellPictureSize
    psFitToRowHeight = 0
    ps16x16 = 16
    ps32x32 = 32
    ps48x48 = 48
End Enum

'***[Constants]*********************************************************************
Const HSCROLLBAR As Long = 0
Const VSCROLLBAR As Long = 1

'***[Default Constants]***************************************************************
Private Const mvar_def_BorderStyle          As Long = bsFixedSingle

Private Const mvar_def_ScrollBars           As Long = sbNone
Private Const mvar_def_SmallChangeH         As Long = 1
Private Const mvar_def_SmallChangeV         As Long = 1
Private Const mvar_def_LargeChangeH         As Long = 1
Private Const mvar_def_LargeChangeV         As Long = 10
Private Const mvar_def_ScrollTrack          As Boolean = False

Private Const mvar_def_ColsFixed            As Long = 1
Private Const mvar_def_RowsFixed            As Long = 1
Private Const mvar_def_Cols                 As Long = 5
Private Const mvar_def_Rows                 As Long = 150
Private Const mvar_def_ColWidth             As Long = 70 'Default Column Width
Private Const mvar_def_RowHeight            As Long = 16
Private Const mvar_def_ExtendLastCol        As Boolean = False

Private Const mvar_def_GridColor            As Long = vbButtonFace
Private Const mvar_def_GridColorFixed       As Long = vbButtonShadow

Private Const mvar_def_BackColor            As Long = vbWindowBackground
Private Const mvar_def_BackColorAlternate   As Long = vbWindowBackground
Private Const mvar_def_BackColorFixed       As Long = vbButtonFace
Private Const mvar_def_WindowBackColor      As Long = vbApplicationWorkspace
Private Const mvar_def_BackColorSel         As Long = vbHighlight

Private Const mvar_def_ForeColor            As Long = vbWindowText
Private Const mvar_def_ForeColorFixed       As Long = vbButtonText
Private Const mvar_def_ForeColorSel         As Long = vbHighlightText

Private Const mvar_def_SheetBorder          As Long = vbWindowFrame

Private Const mvar_def_CellTextWrap         As Boolean = False
Private Const mvar_def_CellPictureSize      As Long = psFitToRowHeight
Private Const mvar_def_CellMaskColor        As Long = vbButtonFace
Private Const mvar_def_Editable             As Boolean = False

Private Const mvar_def_SelectionMode        As Long = smCell

'***[Property Variables]***************************************************************
Private mvarBorderStyle             As enuBorderStyle

Private mvarScrollBars              As enuScrollBars
Private mvarCurrentPosH             As Long
Private mvarCurrentPosV             As Long
Private mvarCurrentPosHOld          As Long
Private mvarCurrentPosVOld          As Long
Private mvarScrollInfo              As SCROLLINFO
Private mvarSmallChangeH            As Long
Private mvarSmallChangeV            As Long
Private mvarLargeChangeH            As Long
Private mvarLargeChangeV            As Long
Private mvarScrollTrack             As Boolean

Private mvarColsFixed               As Long
Private mvarRowsFixed               As Long
Private mvarCols                    As Long
Private mvarRows                    As Long
Private mvarRowHeight               As Long
Private mvarExtendLastCol           As Boolean

Private mvarGridColor               As OLE_COLOR
Private mvarGridColorFixed          As OLE_COLOR

Private mvarBackColor               As OLE_COLOR
Private mvarBackColorAlternate      As OLE_COLOR
Private mvarBackColorFixed          As OLE_COLOR
Private mvarWindowBackColor         As OLE_COLOR
Private mvarBackColorSel            As OLE_COLOR

Private mvarForeColor               As OLE_COLOR
Private mvarForeColorFixed          As OLE_COLOR
Private mvarForeColorSel            As OLE_COLOR

Private mvarSheetBorder             As OLE_COLOR

Private mvarCellTextWrap            As Boolean
Private mvarCellPictureSize         As enuCellPictureSize
Private mvarCellMaskColor           As OLE_COLOR
Private mvarEditable                As Boolean

Private mvarSelectionMode           As enuSelectionMode

'***[Shared Variables]***************************************************************
Private mobjSubClasser              As SuperSubClasser

Private mvarTotalFixedWidth         As Long
Private mvarTotalFixedHeight        As Long

Private mobjRows()                  As stcRow 'Main Cell Collection
Private mvarColWidth()              As Long   'Column widths
Private mvarColWidthOptimal()       As Long   'Optimal Column widths determined in 'RenderControl' sub

Private mobjDatasource              'As IDataSource
    
Private mvarFocusRECT               As RECT
Private mvarFocusCellX              As Long
Private mvarFocusCellY              As Long

Private mvarMouseX                  As Single
Private mvarMouseY                  As Single

Private mvarEditingText             As Boolean
Private mvarResizingColumn          As Boolean
    
'***[Events]*********************************************************************
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event Move()
Event Resize()
Event ScrollH(Value As Long)
Event ScrollV(Value As Long)
Event EndScroll()
Event CellMouseMove(ColID As Long, RowID As Long)
Event CellMouseClick(ColID As Long, RowID As Long)
Event CellBeforeEdit(ColID As Long, RowID As Long)
Event CellAfterEdit(ColID As Long, RowID As Long)

'***[Life control]***************************************************************
Private Sub UserControl_Initialize()
    Set mobjSubClasser = New SuperSubClasser
    
    'Add only messages needed for processing elemental events
    'above which others are made
    With mobjSubClasser
        .AddMsg (WM_MOVE)
        .AddMsg (WM_SIZE)
        .AddMsg (WM_HSCROLL)
        .AddMsg (WM_VSCROLL)
        .Subclass UserControl.hWnd, Me, False
    End With
    
    mvarRowHeight = mvar_def_RowHeight
End Sub

Private Sub PrepareMatrix()
    Dim i As Long
    
    ReDim Preserve mobjRows(mvarRows)
    For i = 1 To mvarRows
        ReDim Preserve mobjRows(i).Cols(mvarCols)
    Next i

    ReDim Preserve mvarColWidth(mvarCols)
    ReDim Preserve mvarColWidthOptimal(mvarCols)
    For i = 1 To mvarCols
        If mvarColWidth(i) = 0 Then mvarColWidth(i) = mvar_def_ColWidth
        mvarColWidthOptimal(i) = 5
    Next i
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            If mvarFocusCellY < mvarRows Then
                SetSelection False, False
                mobjRows(mvarFocusCellY).Cols(mvarFocusCellX).HasFocus = False
                mvarFocusCellY = mvarFocusCellY + 1
                mobjRows(mvarFocusCellY).Cols(mvarFocusCellX).HasFocus = True
                If mvarFocusCellY - mvarCurrentPosV > MaxVisibleRows Then
                    SetSelection True, False
                    DoEvents
                    CurrentPosV = CurrentPosV + 1
                Else
                    SetSelection True, True
                End If
            End If
        Case vbKeyUp
            If mvarFocusCellY > mvarRowsFixed + 1 Then
                SetSelection False, False
                mobjRows(mvarFocusCellY).Cols(mvarFocusCellX).HasFocus = False
                mvarFocusCellY = mvarFocusCellY - 1
                mobjRows(mvarFocusCellY).Cols(mvarFocusCellX).HasFocus = True
                If mvarFocusCellY - mvarCurrentPosV = mvarRowsFixed Then
                    SetSelection True, False
                    DoEvents
                    CurrentPosV = CurrentPosV - 1
                Else
                    SetSelection True, True
                End If
            End If
        Case vbKeyRight
            If mvarFocusCellX < mvarCols Then
                SetSelection False, False
                mobjRows(mvarFocusCellY).Cols(mvarFocusCellX).HasFocus = False
                mvarFocusCellX = mvarFocusCellX + 1
                mobjRows(mvarFocusCellY).Cols(mvarFocusCellX).HasFocus = True
                If mvarFocusCellX - mvarCurrentPosH > MaxVisibleCols Then
                    SetSelection True, False
                    DoEvents
                    CurrentPosH = CurrentPosH + 1
                Else
                    SetSelection True, True
                End If
            End If
        Case vbKeyLeft
            If mvarFocusCellX > mvarColsFixed + 1 Then
                SetSelection False, False
                mobjRows(mvarFocusCellY).Cols(mvarFocusCellX).HasFocus = False
                mvarFocusCellX = mvarFocusCellX - 1
                mobjRows(mvarFocusCellY).Cols(mvarFocusCellX).HasFocus = True
                If mvarFocusCellX - mvarCurrentPosH = mvarColsFixed Then
                    SetSelection True, False
                    DoEvents
                    CurrentPosH = CurrentPosH - 1
                Else
                    SetSelection True, True
                End If
            End If
    End Select
        
End Sub

Private Sub SetSelection(Selected As Boolean, RenderAll As Boolean)
    Dim i As Long
    
    Select Case mvarSelectionMode
        Case smCell
        Case smRow
            For i = 1 To mvarCols
                mobjRows(mvarFocusCellY).Cols(i).Selected = Selected
                If mobjRows(mvarFocusCellY).Cols(i).HasFocus = True Then
                    mobjRows(mvarFocusCellY).Cols(i).Selected = False
                End If
            Next i
    End Select
    If RenderAll = True Then
        DoEvents
        RenderControl
    End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim myColID As Long
    Dim myRowID As Long
    
    RaiseEvent MouseDown(Button, Shift, X, Y)
    
    myColID = ColFromMouse(X)
    myRowID = RowFromMouse(Y)
    
    
    UserControl.SetFocus
    txtTextEdit.Visible = False
    If mvarEditingText = True Then
        mobjRows(mvarFocusCellY).Cols(mvarFocusCellX).Text = txtTextEdit.Text
        txtTextEdit.Visible = False
        mobjRows(mvarFocusCellY).Cols(mvarFocusCellX).HasFocus = True
        DrawCell mvarFocusCellX, mvarFocusCellY
        mvarEditingText = False
        'Raise an event after editing cell
        RaiseEvent CellAfterEdit(myColID, myRowID)
    End If


    If Button = vbLeftButton Then
        'Check standard Cells for editing
        If myColID <= mvarCols And myRowID <= mvarRows Then
            'First unset focus to old cell
            mobjRows(mvarFocusCellY).Cols(mvarFocusCellX).HasFocus = False
            SetSelection False, False
            DrawCell mvarFocusCellX, mvarFocusCellY
            mvarFocusCellX = myColID
            mvarFocusCellY = myRowID
            'Set focus to newly selected cell
            mobjRows(mvarFocusCellY).Cols(mvarFocusCellX).HasFocus = True
            SetSelection True, True
            'DrawCell mvarFocusCellX, mvarFocusCellY
            RaiseEvent CellMouseClick(myColID, myRowID)
        End If
        'Check fixed cells for resizing
        If myRowID <= mvarRowsFixed Then
            If X = WidthOfCols(myColID) Then
                mvarColWidth(myColID) = X - WidthOfCols(myColID - 1)
                Debug.Print mvarColWidth(myColID), X - WidthOfCols(myColID - 1)
                RenderControl
            End If
        End If
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim myColID As Long
    Dim myRowID As Long
    
    mvarMouseX = X
    mvarMouseY = Y
    RaiseEvent MouseMove(Button, Shift, X, Y)
    myColID = ColFromMouse(X)
    myRowID = RowFromMouse(Y)
    RaiseEvent CellMouseMove(myColID, myRowID)
    If myRowID <= mvarRowsFixed Then
        If X = WidthOfCols(myColID) Then
            Screen.MousePointer = vbSizeWE
        Else
            Screen.MousePointer = vbDefault
        End If
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_DblClick()
    Dim myColID As Long
    Dim myRowID As Long
    
    'Exit sub if Editing is disabled
    If mvarEditable = False Then Exit Sub
    
    myColID = ColFromMouse(mvarMouseX)
    myRowID = RowFromMouse(mvarMouseY)
    'Check if this cell is fixed
    If myColID - mvarCurrentPosH <= mvarColsFixed Or myRowID - mvarCurrentPosV <= mvarRowsFixed Then Exit Sub
    
    
    'Enable cell content editing
    If myColID <= mvarCols And myRowID <= mvarRows Then
        
        'Raise an event Before editing cell
        RaiseEvent CellBeforeEdit(myColID, myRowID)
            
        'First unset focus to old cell
        mobjRows(mvarFocusCellY).Cols(mvarFocusCellX).HasFocus = False
        DrawCell mvarFocusCellX, mvarFocusCellY
        mvarFocusCellX = myColID
        mvarFocusCellY = myRowID
    
        'Show the text box in front of cell to allow editing
        mvarEditingText = True
        txtTextEdit.Left = WidthOfCols(myColID - 1)
        txtTextEdit.Top = (myRowID - mvarCurrentPosV - 1) * mvarRowHeight
        If myColID <> mvarCols Then
            txtTextEdit.Width = mvarColWidth(myColID) - 1
        Else
            If ExtendLastCol = True Then
                txtTextEdit.Width = ScaleWidth - WidthOfCols(myColID - 1) - 1
            Else
                txtTextEdit.Width = mvarColWidth(myColID) - 1
            End If
        End If
        txtTextEdit.Height = mvarRowHeight - 1
        txtTextEdit.Text = mobjRows(myRowID).Cols(myColID).Text
        Select Case mobjRows(myRowID).Cols(myColID).Alignment
            Case caLeft
                txtTextEdit.Alignment = 0
            Case caCenter
                txtTextEdit.Alignment = 2
            Case caRight
                txtTextEdit.Alignment = 1
        End Select
        txtTextEdit.Visible = True
        txtTextEdit.SetFocus
    End If

End Sub

Private Sub txtTextEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyTab Or vbKeyReturn
            mobjRows(mvarFocusCellY).Cols(mvarFocusCellX).Text = txtTextEdit.Text
            txtTextEdit.Visible = False
            mobjRows(mvarFocusCellY).Cols(mvarFocusCellX).HasFocus = True
            'DrawCell mvarFocusCellX, mvarFocusCellY
            RenderControl
            mvarEditingText = False
            'Raise an event after editing cell
            RaiseEvent CellAfterEdit(mvarFocusCellX, mvarFocusCellY)

        Case vbKeyEscape
            mobjRows(mvarFocusCellY).Cols(mvarFocusCellX).HasFocus = True
            DrawCell mvarFocusCellX, mvarFocusCellY
            mvarEditingText = False
            'Raise an event after editing cell
            RaiseEvent CellAfterEdit(mvarFocusCellX, mvarFocusCellY)
    
    End Select
End Sub

Private Sub txtTextEdit_LostFocus()
    txtTextEdit.Visible = False
End Sub


Private Sub UserControl_GotFocus()
    mobjRows(mvarFocusCellY).Cols(mvarFocusCellX).HasFocus = True
    DrawCell mvarFocusCellX, mvarFocusCellY
End Sub

Private Sub UserControl_LostFocus()
    Dim myForeColor As OLE_COLOR
    Dim myBackColor As OLE_COLOR
    If mvarFocusCellX > 0 And mvarFocusCellY > 0 Then
        myForeColor = mobjRows(mvarFocusCellY).Cols(mvarFocusCellX).ForeColor
        myBackColor = mobjRows(mvarFocusCellY).Cols(mvarFocusCellX).BackColor

        mobjRows(mvarFocusCellY).Cols(mvarFocusCellX).ForeColor = mvarForeColorSel
        mobjRows(mvarFocusCellY).Cols(mvarFocusCellX).BackColor = mvarBackColorSel
        mobjRows(mvarFocusCellY).Cols(mvarFocusCellX).HasFocus = False
        DrawCell mvarFocusCellX, mvarFocusCellY
        mobjRows(mvarFocusCellY).Cols(mvarFocusCellX).ForeColor = myForeColor
        mobjRows(mvarFocusCellY).Cols(mvarFocusCellX).BackColor = myBackColor
    Else
        mobjRows(mvarFocusCellY).Cols(mvarFocusCellX).HasFocus = False
        DrawCell mvarFocusCellX, mvarFocusCellY
    End If
End Sub


Private Sub UserControl_Terminate()
    mobjSubClasser.UnSubclass
    Set mobjSubClasser = Nothing
    
    Erase mobjRows
End Sub

Private Sub SetScrollBars()
    Dim myhWnd As Long
    myhWnd = UserControl.hWnd

   
    'Set the scrollbars depending on the user selection but
    'also keep track of horizontal/vertical size of the grid
    'so don't display scrollbar if width/height of all cells
    'doesn't exceed controls width/height
    Select Case mvarScrollBars
        Case sbNone
            ShowScrollBar myhWnd, HSCROLLBAR, API_FALSE
            ShowScrollBar myhWnd, VSCROLLBAR, API_FALSE
        Case sbHorizontal
            If MaxVisibleCols < mvarCols Then
                ShowScrollBar myhWnd, HSCROLLBAR, API_TRUE
            Else
                ShowScrollBar myhWnd, HSCROLLBAR, API_FALSE
            End If
            ShowScrollBar myhWnd, VSCROLLBAR, API_FALSE
        Case sbVertical
            ShowScrollBar myhWnd, HSCROLLBAR, API_FALSE
            If MaxVisibleRows < mvarRows Then
                ShowScrollBar myhWnd, VSCROLLBAR, API_TRUE
            Else
                ShowScrollBar myhWnd, VSCROLLBAR, API_FALSE
            End If
        Case sbBoth
            If MaxVisibleCols - 1 < mvarCols Then
                ShowScrollBar myhWnd, HSCROLLBAR, API_TRUE
            Else
                ShowScrollBar myhWnd, HSCROLLBAR, API_FALSE
            End If
            If MaxVisibleRows < mvarRows Then
                ShowScrollBar myhWnd, VSCROLLBAR, API_TRUE
            Else
                ShowScrollBar myhWnd, VSCROLLBAR, API_FALSE
            End If
    End Select
    
    'Set Vertical scroller range
    GetScrollInfo hWnd, SB_VERT, mvarScrollInfo
    With mvarScrollInfo
        .fMask = SIF_RANGE Or SIF_PAGE
        .nMin = 0
        .nMax = mvarRows + mvarLargeChangeV - MaxVisibleRows - 1
        .nPage = mvarLargeChangeV
    End With
    SetScrollInfo hWnd, SB_VERT, mvarScrollInfo, True
    
    'Set Horizontal scroller range
    GetScrollInfo hWnd, SB_HORZ, mvarScrollInfo
    With mvarScrollInfo
        .fMask = SIF_RANGE
        .nMin = 0
        .nMax = mvarCols - mvarColsFixed - 1
        .nPage = mvarLargeChangeH
    End With
    SetScrollInfo hWnd, SB_HORZ, mvarScrollInfo, True
 
End Sub

Private Sub RenderControl()
    Dim myCellType As enuCellType
    Dim myX As Long
    Dim myY As Long
    Dim myWidth As Long
    Dim myAlternate As Boolean
    Dim i As Long
    Dim j As Long
    
    'Set the MaskColor for Icon rendering
    UserControl.MaskColor = mvarCellMaskColor
    
    If mvarRows = 0 Or mvarCols = 0 Then Exit Sub
    
    'LockWindowUpdate UserControl.hWnd
    Cls
    
    'Determine Width and height of Fixed cells area
    mvarTotalFixedWidth = mvarColsFixed * mvar_def_ColWidth
    mvarTotalFixedHeight = mvarRowsFixed * mvarRowHeight

    'Draw Cells
    For j = 1 To MaxVisibleRows + 1
        For i = 1 To mvarCols
            If i <= MaxVisibleCols Then
                If j <= mvarRowsFixed Then
                    If i <= mvarColsFixed Then
                        DrawCell i, j
                    Else
                        DrawCell i + mvarCurrentPosH, j
                    End If
                Else
                    If i <= mvarColsFixed Then
                        If j + mvarCurrentPosV <= mvarRows Then
                            DrawCell i, j + mvarCurrentPosV
                        End If
                    Else
                        If j + mvarCurrentPosV <= mvarRows Then
                            DrawCell i + mvarCurrentPosH, j + mvarCurrentPosV
                        End If
                    End If
                End If
            Else
                Exit For
            End If
        Next i
    Next j
    
    'Draw SheetBorder
    'Bottom line
    If mvarExtendLastCol = False Then
        Line (0, (mvarRows - mvarCurrentPosV) * mvarRowHeight)-(WidthOfCols(mvarCols, True) + 1, (mvarRows - mvarCurrentPosV) * mvarRowHeight), mvarSheetBorder
    End If
    'Right line
    If mvarExtendLastCol = False Then
        If WidthOfCols(mvarCols, True) < ScaleWidth Then
            Line (WidthOfCols(mvarCols, True), 0)-(WidthOfCols(mvarCols, True), (mvarRows - mvarCurrentPosV) * mvarRowHeight), mvarSheetBorder
        End If
    End If

    'LockWindowUpdate 0
End Sub

Private Sub DrawCell(ColID As Long, RowID As Long)
    Dim myCell As SuperCell
    Dim myCellType As enuCellType
    Dim myAlternate As Boolean
    Dim myForeColor As OLE_COLOR
    Dim myBackColor As OLE_COLOR
    Dim myWidthOfCols As Long
    Dim myRect As RECT
    Dim myhDC As Long
    
    If ColID = 0 Or RowID = 0 Or ColID > mvarCols Or RowID > mvarRows Then Exit Sub
    
    'Set the cell to specified cell in matrix
    Set myCell = mobjRows(RowID).Cols(ColID)
    
    myhDC = UserControl.hDC
    
    'Determine cell type
    If ColID <= mvarColsFixed Or RowID <= mvarRowsFixed Then
        myCellType = ctFixed
    ElseIf ColID <= mvarColsFixed Then
        myCellType = ctFixed
    ElseIf RowID <= mvarRowsFixed Then
        myCellType = ctFixed
    Else
        myCellType = ctStandard
    End If
    
    'Every second row should be painted with alternate back color
    'But color is overridden if user has manualy set the BackColor
    'for specific cell
    'BackCcolor for cells of ctFixed type are excluded from this
    'because user can't change BackColor for individual fixed cell
    'their BackColor is defined in BackColorFixed property
    'Render Cell textual content
    If myCellType = ctStandard Then
        If myCell.Selected = False Then
            If myCell.ForeColor = -1 Then
                myForeColor = mvarForeColor
            Else
                myForeColor = myCell.ForeColor
            End If
            If myCell.BackColor = -1 Then
                'Determine if Cell belongs to odd row
                'and set appropriate BackColor
                If Abs((RowID - mvarRowsFixed) Mod 2) = 0 Then
                    myAlternate = True
                    myBackColor = mvarBackColorAlternate
                Else
                    myAlternate = False
                    myBackColor = mvarBackColor
                End If
            Else
                myBackColor = myCell.BackColor
            End If
        Else
            myForeColor = mvarForeColorSel
            myBackColor = mvarBackColorSel
        End If
    Else
        myForeColor = mvarForeColorFixed
        myBackColor = mvarBackColorFixed
    End If
    
    'Determine Cells Coordinates and width
    'This is a bit tricky part because we have to have in mind fixed cols/rows
    'and current Scrollbar value
    'First Determine width of all cols before this
    myWidthOfCols = WidthOfCols(ColID - 1)
    'Define coordinates for this cell
    If myCellType = ctStandard Then
        myRect.Left = myWidthOfCols
        myRect.Top = mvarTotalFixedHeight + (RowID - mvarRowsFixed - (mvarCurrentPosV + 1)) * mvarRowHeight
    Else
        If RowID <= mvarRowsFixed Then
            myRect.Left = myWidthOfCols
            myRect.Top = (RowID - 1) * mvarRowHeight
        Else
            myRect.Left = myWidthOfCols
            myRect.Top = (RowID - (mvarCurrentPosV + 1)) * mvarRowHeight
        End If
    End If
    'Set Cell height
    myRect.Bottom = myRect.Top + mvarRowHeight
    'If this is last column and property ExtendLastCol is true
    'We'll extend the width of the cell to the end of visible area
    If ColID = mvarCols And mvarExtendLastCol = True Then
        myRect.Right = myRect.Left + (ScaleWidth - WidthOfCols(ColID - 1))
    Else
        myRect.Right = myRect.Left + mvarColWidth(ColID)
    End If
    
    'Render Cell layout
    Select Case myCellType
    Case ctStandard
        'Back rectangle
        Line (myRect.Left, myRect.Top)-(myRect.Right - 2, myRect.Bottom - 2), myBackColor, BF
        'Bottom line
        Line (myRect.Left, myRect.Bottom - 1)-(myRect.Right, myRect.Bottom - 1), mvarGridColor
        'Right line
        Line (myRect.Right - 1, myRect.Top)-(myRect.Right - 1, myRect.Bottom - 1), mvarGridColor
    Case ctFixed
        'Back rectangle
        Line (myRect.Left + 1, myRect.Top + 1)-(myRect.Right - 2, myRect.Bottom - 2), myBackColor, BF
        'Top line
        Line (myRect.Left, myRect.Top)-(myRect.Right - 1, myRect.Top), vbButtonHighlight
        'Left line
        Line (myRect.Left, myRect.Top)-(myRect.Left, myRect.Bottom - 1), vbButtonHighlight
        'Bottom line
        Line (myRect.Left, myRect.Bottom - 1)-(myRect.Right - 1, myRect.Bottom - 1), vbButtonShadow
        'Right line
        Line (myRect.Right - 1, myRect.Top)-(myRect.Right - 1, myRect.Bottom - 1), vbButtonShadow
    End Select
    
    
    'Set FontBold
    If myCell.FontBold = False Then
        UserControl.FontBold = False
    Else
        UserControl.FontBold = True
    End If
        
    'Set FontBold
    If myCell.FontItalic = False Then
        UserControl.FontItalic = False
    Else
        UserControl.FontItalic = True
    End If
    
    myRect.Left = myRect.Left + 2
    myRect.Right = myRect.Right - 3
    myRect.Top = myRect.Top + 1
    myRect.Bottom = myRect.Bottom - 1
    
    'Draw the text (and picture) on our cell, including the alignment
    UserControl.ForeColor = myForeColor
    If Not myCell.Picture Is Nothing Then
        UserControl.PaintPicture myCell.Picture, myRect.Left - 1, myRect.Top, mvarRowHeight - 3, mvarRowHeight - 3
        myRect.Left = myRect.Left + (mvarRowHeight - 3)
        If mvarCellTextWrap = False Then
            DrawText Me.hDC, myCell.Text, Len(myCell.Text), myRect, myCell.Alignment
        Else
            DrawText Me.hDC, myCell.Text, Len(myCell.Text), myRect, myCell.Alignment Or DT_WORDBREAK
        End If
    Else
        If mvarCellTextWrap = False Then
            DrawText Me.hDC, myCell.Text, Len(myCell.Text), myRect, myCell.Alignment
        Else
            DrawText Me.hDC, myCell.Text, Len(myCell.Text), myRect, myCell.Alignment Or DT_WORDBREAK
        End If
    End If
    
    'Draw Focus rectangle
    If myCell.HasFocus = True And myCellType = ctStandard Then
        DrawFocusRectangle
    End If
   
    Set myCell = Nothing
End Sub
    
Private Sub DrawFocusRectangle()
    UserControl.ForeColor = vbBlack
    With mvarFocusRECT
        .Left = WidthOfVisibleCols(mvarFocusCellX - 1)
        If mvarFocusCellX <> mvarCols Then
            .Right = WidthOfVisibleCols(mvarFocusCellX)
        Else
            If ExtendLastCol = True Then
                .Right = WidthOfVisibleCols(mvarFocusCellX)
            Else
                .Right = WidthOfVisibleCols(mvarFocusCellX - 1) + mvarColWidth(mvarFocusCellX) - 1
            End If
        End If
        .Top = (mvarFocusCellY - mvarCurrentPosV - 1) * mvarRowHeight
        .Bottom = (mvarFocusCellY - mvarCurrentPosV) * mvarRowHeight - 1
    End With
    If mvarFocusRECT.Left <> mvarFocusRECT.Right Then
        DrawFocusRect hDC, mvarFocusRECT
    End If
End Sub

Private Function MaxVisibleRows() As Long
    MaxVisibleRows = ScaleHeight \ mvarRowHeight
End Function

Private Function MaxVisibleCols() As Long
    Dim myWidth As Long
    Dim i As Long
    
    myWidth = 0
    For i = 1 To mvarCols
        If myWidth + mvarColWidth(i) < ScaleWidth Then
            myWidth = myWidth + mvarColWidth(i)
        Else
            Exit For
        End If
        If i = mvarColsFixed Then i = i + mvarCurrentPosH
    Next i
    MaxVisibleCols = i
End Function

Private Function WidthOfCols(LastColID As Long, Optional Total As Boolean = False) As Long
    Dim myWidth As Long
    Dim i As Long
    
    myWidth = 0
    If LastColID > mvarCols Then LastColID = mvarCols
    For i = 1 To LastColID
        If Total = True Then
            myWidth = myWidth + mvarColWidth(i)
        ElseIf Total = False And myWidth + mvarColWidth(i) < ScaleWidth Then
            myWidth = myWidth + mvarColWidth(i)
        Else
            Exit For
        End If
        If i = mvarColsFixed Then i = i + mvarCurrentPosH
    Next i
    WidthOfCols = myWidth
End Function

Private Function WidthOfVisibleCols(LastColID As Long) As Long
    Dim myWidth As Long
    Dim i As Long
    
    myWidth = 0
    If LastColID > mvarCols Then LastColID = mvarCols
    For i = 1 To LastColID
        If myWidth + mvarColWidth(i) < ScaleWidth Then
            myWidth = myWidth + mvarColWidth(i)
        Else
            myWidth = myWidth + mvarColWidth(i)
            Exit For
        End If
        If i = mvarColsFixed Then i = i + mvarCurrentPosH
    Next i
    WidthOfVisibleCols = myWidth
End Function

Private Function ColFromMouse(X As Single) As Long
    Dim myWidth As Long
    Dim i As Long
    
    myWidth = 0
    For i = 1 To mvarCols
        If myWidth + mvarColWidth(i) < X Then
            myWidth = myWidth + mvarColWidth(i)
        Else
            Exit For
        End If
                
        If i = mvarColsFixed Then i = i + mvarCurrentPosH
    Next i
    'If i > mvarCols Then i = mvarCols
    ColFromMouse = i
End Function

Private Function RowFromMouse(Y As Single)
    Dim myY As Long
    myY = (Y \ mvarRowHeight) + 1
    If myY <= mvarRowsFixed Then
        RowFromMouse = myY
    Else
        RowFromMouse = myY + mvarCurrentPosV
    End If
End Function

Public Property Let BorderStyle(Value As enuBorderStyle)
    mvarBorderStyle = Value
    PropertyChanged ("BorderStyle")
    UserControl.BorderStyle = mvarBorderStyle
End Property

Public Property Get BorderStyle() As enuBorderStyle
    BorderStyle = mvarBorderStyle
End Property

Public Property Let ScrollBars(Value As enuScrollBars)
    mvarScrollBars = Value
    SetScrollBars
End Property

Public Property Get ScrollBars() As enuScrollBars
    ScrollBars = mvarScrollBars
End Property

Public Property Let CurrentPosH(Value As Long)
    'First check which scrollbar is visible and is it posssble to set the position
    If mvarScrollBars = sbNone Or mvarScrollBars = sbVertical Then
        Exit Property
    End If
    
    If Value > mvarCols Then
        mvarCurrentPosH = mvarCols
    Else
        mvarCurrentPosH = Value
    End If
 
    If Value < 0 Then mvarCurrentPosH = 0
    
    With mvarScrollInfo
        .fMask = SIF_POS
        .nPos = Value
    End With
    
    SetScrollInfo hWnd, SB_HORZ, mvarScrollInfo, True

    If mvarScrollTrack = True Then
        RenderControl
    End If
    
    RaiseEvent ScrollH(mvarCurrentPosH)
End Property

Public Property Get CurrentPosH() As Long
    CurrentPosH = mvarCurrentPosH
End Property

Public Property Let CurrentPosV(Value As Long)
    'First check which scrollbar is visible and is it posssble to set the position
    If mvarScrollBars = sbNone Or mvarScrollBars = sbHorizontal Then
        Exit Property
    End If
    
    If Value > mvarRows Then
        mvarCurrentPosV = mvarRows
    Else
        mvarCurrentPosV = Value
    End If

    If Value < 0 Then mvarCurrentPosV = 0

    With mvarScrollInfo
        .fMask = SIF_POS
        .nPos = Value
    End With
    
    SetScrollInfo hWnd, SB_VERT, mvarScrollInfo, True
    
    If mvarScrollTrack = True Then
        RenderControl
    End If
    
    RaiseEvent ScrollV(mvarCurrentPosV)
End Property

Public Property Get CurrentPosV() As Long
    CurrentPosV = mvarCurrentPosV
End Property

Public Property Let SmallChangeH(Value As Long)
    Value = Abs(Value)
    If Value > mvarLargeChangeH Then
        mvarSmallChangeH = mvarLargeChangeH
    Else
        mvarSmallChangeH = Value
    End If
End Property

Public Property Get SmallChangeH() As Long
    SmallChangeH = mvarSmallChangeH
End Property

Public Property Let SmallChangeV(Value As Long)
    Value = Abs(Value)
    If Value > mvarLargeChangeV Then
        mvarSmallChangeV = mvarLargeChangeV
    Else
        mvarSmallChangeV = Value
    End If
End Property

Public Property Get SmallChangeV() As Long
    SmallChangeV = mvarSmallChangeV
End Property

Public Property Let LargeChangeH(Value As Long)
    Value = Abs(Value)
    If Value > mvarCols Then
        mvarLargeChangeH = mvarCols
    Else
        mvarLargeChangeH = Value
    End If

    SetScrollBars
End Property

Public Property Get LargeChangeH() As Long
    LargeChangeH = mvarLargeChangeH
End Property

Public Property Let LargeChangeV(Value As Long)
    Value = Abs(Value)
    If Value > mvarRows Then
        mvarLargeChangeV = mvarRows
    Else
        mvarLargeChangeV = Value
    End If

    SetScrollBars
End Property

Public Property Get LargeChangeV() As Long
    LargeChangeV = mvarLargeChangeV
End Property


Public Property Let ScrollTrack(Value As Boolean)
    mvarScrollTrack = Value
End Property

Public Property Get ScrollTrack() As Boolean
    ScrollTrack = mvarScrollTrack
End Property

Public Property Let Cols(Value As Long)
    mvarCols = Value
    PropertyChanged ("Cols")
    PrepareMatrix
    SetScrollBars
    RenderControl
End Property

Public Property Get Cols() As Long
    Cols = mvarCols
End Property

Public Property Let Rows(Value As Long)
    mvarRows = Value
    PropertyChanged ("Rows")
    PrepareMatrix
    SetScrollBars
    RenderControl
End Property

Public Property Get Rows() As Long
    Rows = mvarRows
End Property

Public Property Let ColsFixed(Value As Long)
    If Value < 0 Then
        Value = 0
    End If
    If Value > mvarCols Then
        Value = mvarCols
    End If
    mvarColsFixed = Value
    PropertyChanged ("ColsFixed")
    SetScrollBars
    RenderControl
End Property

Public Property Get ColsFixed() As Long
    ColsFixed = mvarColsFixed
End Property

Public Property Let RowsFixed(Value As Long)
    If Value < 0 Then
        Value = 0
    End If
    If Value > mvarRows Then
        Value = mvarRows
    End If
    mvarRowsFixed = Value
    PropertyChanged ("RowsFixed")
    SetScrollBars
    RenderControl
End Property

Public Property Get RowsFixed() As Long
    RowsFixed = mvarRowsFixed
End Property

Public Property Let ExtendLastCol(Value As Boolean)
    mvarExtendLastCol = Value
    PropertyChanged ("ExtendLastCol")
    RenderControl
End Property

Public Property Get ExtendLastCol() As Boolean
    ExtendLastCol = mvarExtendLastCol
End Property

Public Property Let ColWidth(Index As Long, Value As Long)
    mvarColWidth(Index) = Value
    RenderControl
End Property

Public Property Get ColWidth(Index As Long) As Long
    ColWidth = mvarColWidth(Index)
End Property

Public Property Let RowHeight(Value As Long)
    mvarRowHeight = Value
    RenderControl
End Property

Public Property Get RowHeight() As Long
    RowHeight = mvarRowHeight
End Property

Public Property Let GridColor(Value As OLE_COLOR)
    mvarGridColor = Value
    PropertyChanged ("GridColor")
    RenderControl
End Property

Public Property Get GridColor() As OLE_COLOR
    GridColor = mvarGridColor
End Property

Public Property Let GridColorFixed(Value As OLE_COLOR)
    mvarGridColorFixed = Value
    PropertyChanged ("GridColorFixed")
    RenderControl
End Property

Public Property Get GridColorFixed() As OLE_COLOR
    GridColorFixed = mvarGridColorFixed
End Property

Public Property Let BackColor(Value As OLE_COLOR)
    mvarBackColor = Value
    PropertyChanged ("BackColor")
    RenderControl
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = mvarBackColor
End Property

Public Property Let BackColorAlternate(Value As OLE_COLOR)
    mvarBackColorAlternate = Value
    PropertyChanged ("BackColorAlternate")
    RenderControl
End Property

Public Property Get BackColorAlternate() As OLE_COLOR
    BackColorAlternate = mvarBackColorAlternate
End Property

Public Property Let WindowBackColor(Value As OLE_COLOR)
    mvarWindowBackColor = Value
    PropertyChanged ("WindowBackColor")
    UserControl.BackColor = mvarWindowBackColor
    RenderControl
End Property

Public Property Get WindowBackColor() As OLE_COLOR
    WindowBackColor = mvarWindowBackColor
End Property

Public Property Let BackColorFixed(Value As OLE_COLOR)
    mvarBackColorFixed = Value
    PropertyChanged ("BackColorFixed")
    RenderControl
End Property

Public Property Get BackColorFixed() As OLE_COLOR
    BackColorFixed = mvarBackColorFixed
End Property

Public Property Let BackColorSel(Value As OLE_COLOR)
    mvarBackColorSel = Value
    PropertyChanged ("BackColorSel")
    RenderControl
End Property

Public Property Get BackColorSel() As OLE_COLOR
    BackColorSel = mvarBackColorSel
End Property

Public Property Let ForeColor(Value As OLE_COLOR)
    mvarForeColor = Value
    PropertyChanged ("ForeColor")
    RenderControl
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = mvarForeColor
End Property

Public Property Let ForeColorFixed(Value As OLE_COLOR)
    mvarForeColorFixed = Value
    PropertyChanged ("ForeColorFixed")
    RenderControl
End Property

Public Property Get ForeColorFixed() As OLE_COLOR
    ForeColorFixed = mvarForeColorFixed
End Property

Public Property Let ForeColorSel(Value As OLE_COLOR)
    mvarForeColorSel = Value
    PropertyChanged ("ForeColorSel")
    RenderControl
End Property

Public Property Get ForeColorSel() As OLE_COLOR
    ForeColorSel = mvarForeColorSel
End Property

Public Property Set Font(Value As Font)
    Set UserControl.Font = Value
    Set txtTextEdit.Font = Value
    PropertyChanged ("Font")
    SetRowHeight
    PropertyChanged ("RowHeight")
    SetScrollBars
    RenderControl
End Property

Public Property Get Font() As Font
    Set Font = UserControl.Font
End Property

Private Sub SetRowHeight()
    'We'll add 3 more pixels to TextHeight because of potential 3D borders and
    'proper vertical centering
    Select Case mvarCellPictureSize
        Case psFitToRowHeight
            mvarRowHeight = UserControl.TextHeight("Hy") + 3
        Case ps16x16
            mvarRowHeight = ps16x16 + 3
        Case ps32x32
            mvarRowHeight = ps32x32 + 3
        Case ps48x48
            mvarRowHeight = ps48x48 + 3
    End Select
    PropertyChanged "RowHeight"
End Sub

Public Property Let SheetBorder(Value As OLE_COLOR)
    mvarSheetBorder = Value
    PropertyChanged ("SheetBorder")
    RenderControl
End Property

Public Property Get SheetBorder() As OLE_COLOR
    SheetBorder = mvarSheetBorder
End Property

Public Property Let CellTextWrap(Value As Boolean)
    mvarCellTextWrap = Value
    PropertyChanged ("CellTextWrap")
    RenderControl
End Property

Public Property Get CellTextWrap() As Boolean
    CellTextWrap = mvarCellTextWrap
End Property

Public Property Let Editable(Value As Boolean)
    mvarEditable = Value
    PropertyChanged ("Editable")
    RenderControl
End Property

Public Property Get Editable() As Boolean
    Editable = mvarEditable
End Property

Public Property Let SelectionMode(Value As enuSelectionMode)
    mvarSelectionMode = Value
    PropertyChanged ("SelectionMode")
    RenderControl
End Property

Public Property Get SelectionMode() As enuSelectionMode
    SelectionMode = mvarSelectionMode
End Property

Public Property Let CellPictureSize(Value As enuCellPictureSize)
    mvarCellPictureSize = Value
    PropertyChanged ("CellPictureSize")
    SetRowHeight
    RenderControl
End Property

Public Property Get CellPictureSize() As enuCellPictureSize
    CellPictureSize = mvarCellPictureSize
End Property

Public Property Let CellMaskColor(Value As OLE_COLOR)
    mvarCellMaskColor = Value
    PropertyChanged ("CellMaskColor")
    RenderControl
End Property

Public Property Get CellMaskColor() As OLE_COLOR
    CellMaskColor = mvarCellMaskColor
End Property

Public Property Let Cells(ColID As Long, RowID As Long, Value As SuperCell)
    Set mobjRows(RowID).Cols(ColID) = Value
End Property

Public Property Get Cells(ColID As Long, RowID As Long) As SuperCell
    Set Cells = mobjRows(RowID).Cols(ColID)
End Property

Public Property Get hDC() As Long
    hDC = UserControl.hDC
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Set DataSource(ADO_Recordset As Object)
    Dim i As Long, j As Long
    
    Set mobjDatasource = ADO_Recordset

    With mobjDatasource
        Rows = .RecordCount
        Cols = .Fields.Count
        .MoveFirst
        j = 0
        Do While Not .EOF
            For i = 0 To .Fields.Count - 1
                If Not IsNull(.Fields(i).Value) Then
                    mobjRows(j + 1).Cols(i + 1).Text = CStr(.Fields(i).Value)
                End If
            Next i
            .MoveNext
            j = j + 1
        Loop
    End With
    CurrentPosV = 0
    CurrentPosH = 0
    RenderControl
End Property

'***[Methods]****************************************************************
Public Sub About()
Attribute About.VB_UserMemId = -552
    frmAbout.Show vbModal
End Sub

Public Sub Autosize()
    Dim myTextWidth As Long
    Dim i As Long
    Dim j As Long
    'RenderControl
    For j = 1 To mvarRows
        For i = 1 To mvarCols
            UserControl.FontBold = mobjRows(j).Cols(i).FontBold
            UserControl.FontItalic = mobjRows(j).Cols(i).FontItalic
            If mobjRows(j).Cols(i).Picture Is Nothing Then
                myTextWidth = TextWidth(mobjRows(j).Cols(i).Text)
            Else
                myTextWidth = TextWidth(mobjRows(j).Cols(i).Text) + mvarRowHeight
            End If
            If mvarColWidthOptimal(i) < myTextWidth + 5 Then
                mvarColWidthOptimal(i) = myTextWidth + 5
            End If
            mvarColWidth(i) = mvarColWidthOptimal(i)
        Next i
    Next j
    SetScrollBars
    RenderControl
End Sub

Public Sub Refresh()
    SetScrollBars
    RenderControl
End Sub


Private Sub iSuperSubClasser_Before(lHandled As Long, lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    'DoNothing
End Sub

Private Sub iSuperSubClasser_After(lReturn As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
   'Exit Sub
    Select Case uMsg
        Case WM_MOVE
            RaiseEvent Move
            lReturn = HTLEFT Or HTTOP
        Case WM_SIZE
            SetScrollBars
            CurrentPosH = 0
            'RenderControl 'is done with setting the current pos
            RaiseEvent Resize
            lReturn = HTSIZE
        Case WM_HSCROLL
            UserControl.SetFocus
            mvarEditingText = False
            txtTextEdit.Visible = False
            DoEvents
            With Me
                Select Case LOWORD(wParam)
                    Case SB_BOTTOM
                        .CurrentPosH = mvarCols - MaxVisibleCols
                    Case SB_TOP
                        .CurrentPosH = 0
                    Case SB_LINEDOWN
                        If mvarCurrentPosH < mvarCols - mvarColsFixed - 1 Then
                            .CurrentPosH = .CurrentPosH + mvarSmallChangeH
                        Else
                            .CurrentPosH = mvarCols - mvarColsFixed - 1
                        End If
                    Case SB_LINEUP
                        If mvarCurrentPosH - mvarSmallChangeH > 0 Then
                            .CurrentPosH = .CurrentPosH - mvarSmallChangeH
                        Else
                            .CurrentPosH = 0
                        End If
                    Case SB_PAGEDOWN
                        If mvarCurrentPosH + mvarLargeChangeH < mvarCols - mvarColsFixed - 1 Then
                            .CurrentPosH = .CurrentPosH + mvarLargeChangeH
                        Else
                            .CurrentPosH = mvarCols - mvarColsFixed - 1
                        End If
                    Case SB_PAGEUP
                        If mvarCurrentPosH > mvarLargeChangeH Then
                            .CurrentPosH = .CurrentPosH - mvarLargeChangeH
                        Else
                            .CurrentPosH = 0
                        End If
                    Case SB_THUMBPOSITION, SB_THUMBTRACK
                        .CurrentPosH = HIWORD(wParam)
                   Case SB_ENDSCROLL
                        'Each time any scroll method ends this message is passed
                        'We use it for EndScroll event that might be usefull
                        If mvarScrollTrack = False Then
                            RenderControl
                        End If
                        RaiseEvent EndScroll
                End Select
            End With
            
            lReturn = HTHSCROLL
        Case WM_VSCROLL
            UserControl.SetFocus
            mvarEditingText = False
            txtTextEdit.Visible = False
            DoEvents
            With Me
                Select Case LOWORD(wParam)
                    Case SB_BOTTOM
                        .CurrentPosV = mvarRows - MaxVisibleRows
                    Case SB_TOP
                        .CurrentPosV = 0
                    Case SB_LINEDOWN
                        If mvarCurrentPosV < mvarRows - MaxVisibleRows Then
                            .CurrentPosV = .CurrentPosV + mvarSmallChangeV
                        Else
                            CurrentPosV = mvarRows - MaxVisibleRows
                        End If
                    Case SB_LINEUP
                        If mvarCurrentPosV > 0 Then
                            .CurrentPosV = .CurrentPosV - mvarSmallChangeV
                        Else
                            CurrentPosV = 0
                        End If
                    Case SB_PAGEDOWN
                        If mvarCurrentPosV + mvarLargeChangeV < mvarRows - MaxVisibleRows Then
                            .CurrentPosV = .CurrentPosV + mvarLargeChangeV
                        Else
                            CurrentPosV = mvarRows - MaxVisibleRows
                        End If
                    Case SB_PAGEUP
                        If mvarCurrentPosV > mvarLargeChangeV Then
                            .CurrentPosV = .CurrentPosV - mvarLargeChangeV
                        Else
                            CurrentPosV = 0
                        End If
                    Case SB_THUMBPOSITION, SB_THUMBTRACK
                        .CurrentPosV = HIWORD(wParam)
                    Case SB_ENDSCROLL
                        'Each time any scroll method ends this message is passed
                        'We use it for EndScroll event that might be usefull
                        If mvarScrollTrack = False Then
                            RenderControl
                        End If
                        RaiseEvent EndScroll
                End Select
            End With
            
            lReturn = HTVSCROLL
    End Select
End Sub

'***[Property Bag]***********************************************************
Private Sub UserControl_InitProperties()
    mvarBorderStyle = mvar_def_BorderStyle
    
    mvarScrollBars = mvar_def_ScrollBars
    mvarSmallChangeH = mvar_def_SmallChangeH
    mvarSmallChangeV = mvar_def_SmallChangeV
    mvarLargeChangeH = mvar_def_LargeChangeH
    mvarLargeChangeV = mvar_def_LargeChangeV
    mvarScrollTrack = mvar_def_ScrollTrack
    
    mvarColsFixed = mvar_def_ColsFixed
    mvarRowsFixed = mvar_def_RowsFixed
    mvarCols = mvar_def_Cols
    mvarRows = mvar_def_Rows
    mvarRowHeight = mvar_def_RowHeight
    mvarExtendLastCol = mvar_def_ExtendLastCol
   
    mvarGridColor = mvar_def_GridColor
    mvarGridColorFixed = mvar_def_GridColorFixed
    
    mvarBackColor = mvar_def_BackColor
    mvarBackColorAlternate = mvar_def_BackColorAlternate
    mvarWindowBackColor = mvar_def_WindowBackColor
    mvarBackColorFixed = mvar_def_BackColorFixed
    mvarBackColorSel = mvar_def_BackColorSel

    mvarForeColor = mvar_def_ForeColor
    mvarForeColorFixed = mvar_def_ForeColorFixed
    mvarForeColorSel = mvar_def_ForeColorSel
    
    mvarSheetBorder = mvar_def_SheetBorder

    mvarCellTextWrap = mvar_def_CellTextWrap
    mvarCellPictureSize = mvar_def_CellPictureSize
    mvarCellMaskColor = mvar_def_CellMaskColor
    mvarEditable = mvar_def_Editable
    
    mvarSelectionMode = mvar_def_SelectionMode

    Set UserControl.Font = Ambient.Font

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mvarBorderStyle = PropBag.ReadProperty("BorderStyle", mvar_def_BorderStyle)
    
    mvarScrollBars = PropBag.ReadProperty("ScrollBars", mvar_def_ScrollBars)
    mvarSmallChangeH = PropBag.ReadProperty("SmallChangeH", mvar_def_SmallChangeH)
    mvarSmallChangeV = PropBag.ReadProperty("SmallChangeV", mvar_def_SmallChangeV)
    mvarLargeChangeH = PropBag.ReadProperty("LargeChangeH", mvar_def_LargeChangeH)
    mvarLargeChangeV = PropBag.ReadProperty("LargeChangeV", mvar_def_LargeChangeV)
    mvarScrollTrack = PropBag.ReadProperty("ScrollTrack", mvar_def_ScrollTrack)
    
    mvarColsFixed = PropBag.ReadProperty("ColsFixed", mvar_def_ColsFixed)
    mvarRowsFixed = PropBag.ReadProperty("RowsFixed", mvar_def_RowsFixed)
    mvarCols = PropBag.ReadProperty("Cols", mvar_def_Cols)
    mvarRows = PropBag.ReadProperty("Rows", mvar_def_Rows)
    mvarRowHeight = PropBag.ReadProperty("RowHeight", mvar_def_RowHeight)
    mvarExtendLastCol = PropBag.ReadProperty("ExtendLastCol", mvar_def_ExtendLastCol)
    
    mvarGridColor = PropBag.ReadProperty("GridColor", mvar_def_GridColor)
    mvarGridColorFixed = PropBag.ReadProperty("GridColorFixed", mvar_def_GridColorFixed)
    
    mvarBackColor = PropBag.ReadProperty("BackColor", mvar_def_BackColor)
    mvarBackColorAlternate = PropBag.ReadProperty("BackColorAlternate", mvar_def_BackColorAlternate)
    mvarWindowBackColor = PropBag.ReadProperty("WindowBackColor", mvar_def_WindowBackColor)
    UserControl.BackColor = mvarWindowBackColor
    mvarBackColorFixed = PropBag.ReadProperty("BackColorFixed", mvar_def_BackColorFixed)
    mvarBackColorSel = PropBag.ReadProperty("BackColorSel", mvar_def_BackColorSel)

    mvarForeColor = PropBag.ReadProperty("ForeColor", mvar_def_ForeColor)
    mvarForeColorFixed = PropBag.ReadProperty("ForeColorFixed", mvar_def_ForeColorFixed)
    mvarForeColorSel = PropBag.ReadProperty("ForeColorSel", mvar_def_ForeColorSel)

    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set txtTextEdit.Font = PropBag.ReadProperty("Font", Ambient.Font)

    mvarSheetBorder = PropBag.ReadProperty("SheetBorder", mvar_def_SheetBorder)
    
    mvarCellTextWrap = PropBag.ReadProperty("CellTextWrap", mvar_def_CellTextWrap)
    mvarCellPictureSize = PropBag.ReadProperty("CellPictureSize", mvar_def_CellPictureSize)
    mvarCellMaskColor = PropBag.ReadProperty("CellMaskColor", mvar_def_CellMaskColor)
    mvarEditable = PropBag.ReadProperty("Editable", mvar_def_Editable)
    
    mvarSelectionMode = PropBag.ReadProperty("SelectionMode", mvar_def_SelectionMode)
    
    'Set First Cell to be Cell with focus
    mvarFocusCellX = mvarColsFixed + 1
    mvarFocusCellY = mvarRowsFixed + 1
    
    PrepareMatrix
    
    SetScrollBars
    RenderControl
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "BorderStyle", mvarBorderStyle, mvar_def_BorderStyle
    
    PropBag.WriteProperty "ScrollBars", mvarScrollBars, mvar_def_ScrollBars
    PropBag.WriteProperty "SmallChangeH", mvarSmallChangeH, mvar_def_SmallChangeH
    PropBag.WriteProperty "SmallChangeV", mvarSmallChangeV, mvar_def_SmallChangeV
    PropBag.WriteProperty "LargeChangeH", mvarLargeChangeH, mvar_def_LargeChangeH
    PropBag.WriteProperty "LargeChangeV", mvarLargeChangeV, mvar_def_LargeChangeV
    PropBag.WriteProperty "ScrollTrack", mvarScrollTrack, mvar_def_ScrollTrack
    
    PropBag.WriteProperty "ColsFixed", mvarColsFixed, mvar_def_ColsFixed
    PropBag.WriteProperty "RowsFixed", mvarRowsFixed, mvar_def_RowsFixed
    PropBag.WriteProperty "Cols", mvarCols, mvar_def_Cols
    PropBag.WriteProperty "Rows", mvarRows, mvar_def_Rows
    PropBag.WriteProperty "RowHeight", mvarRowHeight, mvar_def_RowHeight
    PropBag.WriteProperty "ExtendLastCol", mvarExtendLastCol, mvar_def_ExtendLastCol
    
    PropBag.WriteProperty "GridColor", mvarGridColor, mvar_def_GridColor
    PropBag.WriteProperty "GridColorFixed", mvarGridColorFixed, mvar_def_GridColorFixed
    
    PropBag.WriteProperty "BackColor", mvarBackColor, mvar_def_BackColor
    PropBag.WriteProperty "BackColorAlternate", mvarBackColorAlternate, mvar_def_BackColorAlternate
    PropBag.WriteProperty "WindowBackColor", mvarWindowBackColor, mvar_def_WindowBackColor
    PropBag.WriteProperty "BackColorFixed", mvarBackColorFixed, mvar_def_BackColorFixed
    PropBag.WriteProperty "BackColorSel", mvarBackColorSel, mvar_def_BackColorSel
    
    PropBag.WriteProperty "ForeColor", mvarForeColor, mvar_def_ForeColor
    PropBag.WriteProperty "ForeColorFixed", mvarForeColorFixed, mvar_def_ForeColorFixed
    PropBag.WriteProperty "ForeColorSel", mvarForeColorSel, mvar_def_ForeColorSel
    
    PropBag.WriteProperty "Font", UserControl.Font, Ambient.Font
    
    PropBag.WriteProperty "SheetBorder", mvarSheetBorder, mvar_def_SheetBorder
    
    PropBag.WriteProperty "CellTextWrap", mvarCellTextWrap, mvar_def_CellTextWrap
    PropBag.WriteProperty "CellPictureSize", mvarCellPictureSize, mvar_def_CellPictureSize
    PropBag.WriteProperty "CellMaskColor", mvarCellMaskColor, mvar_def_CellMaskColor
    PropBag.WriteProperty "Editable", mvarEditable, mvar_def_Editable
    
    PropBag.WriteProperty "SelectionMode", mvarSelectionMode, mvar_def_SelectionMode
    
End Sub

