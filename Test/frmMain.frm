VERSION 5.00
Object = "{0508871B-4AF1-45CA-8C66-F80FBDC5E620}#9.0#0"; "acSuperGrid.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmMain 
   Caption         =   "Super Grid 1.0 - Test"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7065
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7215
   ScaleWidth      =   7065
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picToolbar 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   0
      ScaleHeight     =   450
      ScaleWidth      =   7065
      TabIndex        =   1
      Top             =   6765
      Width           =   7065
      Begin VB.CommandButton Command1 
         Caption         =   "Populate test data"
         Height          =   345
         Left            =   1380
         TabIndex        =   3
         Top             =   30
         Width           =   1845
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Populate from database"
         Height          =   345
         Left            =   3240
         TabIndex        =   2
         Top             =   30
         Width           =   1845
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9120
      Top             =   7680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":058A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0E7B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":176C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":205D
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":294E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin acSuperGrid.SuperGrid SuperGrid1 
      Align           =   1  'Align Top
      Height          =   6705
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7065
      _ExtentX        =   12462
      _ExtentY        =   11827
      ScrollBars      =   3
      ScrollTrack     =   -1  'True
      RowsFixed       =   2
      Cols            =   6
      RowHeight       =   35
      BackColor       =   12640511
      BackColorAlternate=   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CellPictureSize =   32
      Editable        =   -1  'True
      SelectionMode   =   1
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim i As Long
    
    For i = 1 To 150
        SuperGrid1.Cells(1, i).Text = CStr(i)
        SuperGrid1.Cells(1, i).Alignment = caRight
        SuperGrid1.Cells(1, i).FontBold = True
        SuperGrid1.Cells(2, i).Text = "CellPicture " + CStr(i)
        Set SuperGrid1.Cells(2, i).Picture = ImageList1.ListImages(Int(Rnd(1) * 5 + 1)).Picture
        SuperGrid1.Cells(3, i).Text = "CellTextWrap Test" + CStr(i)
        SuperGrid1.Cells(3, i).Alignment = caCenter
        SuperGrid1.Cells(4, i).Text = "ABC " + CStr(i)
        SuperGrid1.Cells(5, i).Text = "L " + CStr(i)
        SuperGrid1.Cells(5, i).ForeColor = RGB(255, 0, 0)
        SuperGrid1.Cells(5, i).FontItalic = True
        SuperGrid1.Cells(6, i).Text = "Some text AAAAAAAAAAABBBBBBBCCCCCC " + CStr(i)
    Next i
    SuperGrid1.Cells(3, 5).ForeColor = RGB(210, 30, 20)
    SuperGrid1.Cells(3, 5).BackColor = RGB(210, 220, 230)
    'SuperGrid1.Refresh
    SuperGrid1.AutoSize

End Sub

Private Sub Command2_Click()
    Dim myConn As New Connection
    Dim myRst As New Recordset
    
On Local Error GoTo errSkip

    myConn.Open "DSN=Grocery 2000"
    
    myRst.CursorType = adOpenDynamic
    myRst.CursorLocation = adUseClient
    myRst.LockType = adLockOptimistic
    myRst.Open "Select * from CountryQ", myConn
    
    Set SuperGrid1.DataSource = myRst
    myConn.Close
    Set myConn = Nothing
    SuperGrid1.AutoSize
Exit Sub
errSkip:
    MsgBox "Change the connection string property to a Database you have."
End Sub

Private Sub Form_Resize()
    SuperGrid1.Height = ScaleHeight - picToolbar.Height
End Sub


