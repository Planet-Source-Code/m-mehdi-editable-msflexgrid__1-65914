VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8235
   ClientLeft      =   2880
   ClientTop       =   1950
   ClientWidth     =   9645
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   9645
   Begin MSFlexGridLib.MSFlexGrid mfgDetail 
      Height          =   3240
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   5715
      _Version        =   393216
      Rows            =   6
      Cols            =   4
      FixedCols       =   0
      RowHeightMin    =   360
      BackColor       =   16777215
      BackColorFixed  =   -2147483648
      ForeColorFixed  =   8388608
      BackColorSel    =   16777088
      BackColorBkg    =   16777215
      ScrollBars      =   2
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    '///////////////////////////////////////////////////////////////////////////////////////
    '// Programmer : M. Mehdi                                                             //
    '// Email      : mehdipanjwani@msn.com,vb.dbase@yahoo.com                             //
    '// Profession : Software Developer at TechnoSys                                      //
    '// Place      : Karachi,Pakistan                                                     //
    '///////////////////////////////////////////////////////////////////////////////////////
    
    '// Variable for storing text . not necessary (mfgdetail.textmatrix can also be used) //
    '///////////////////////////////////////////////////////////////////////////////////////
    Dim XVal As String

Private Sub Form_Load()
    '// for setting columns width
    '////////////////////////////
    mfgDetail.ColWidth(0) = 2000
    mfgDetail.ColWidth(1) = 2000
    mfgDetail.ColWidth(2) = 2000
    mfgDetail.ColWidth(3) = 2000
End Sub

Private Sub mfgDetail_KeyPress(KeyAscii As Integer)
    
    '// for shifting cursor to column and row when enter pressed
    '///////////////////////////////////////////////////////////
    If KeyAscii = vbKeyReturn Then
        If mfgDetail.Col + 1 = mfgDetail.Cols Then
            If mfgDetail.Row + 1 = mfgDetail.Rows Then mfgDetail.Row = 0: mfgDetail.Col = 0
            mfgDetail.Row = mfgDetail.Row + 1
            mfgDetail.Col = 0
        Else
            mfgDetail.Col = mfgDetail.Col + 1
        End If
    End If
    
    '// 8 = backspace . for deleting characters
    '//////////////////////////////////////////
    If KeyAscii = 8 Then
    If Len(XVal) = 0 Then Exit Sub
        XVal = Left$(XVal, Len(XVal) - 1)
        Exit Sub
    End If
    
    '// for storing texts in Variable
    '////////////////////////////////
    XVal = XVal & Chr(KeyAscii)

End Sub

Private Sub mfgDetail_KeyUp(KeyCode As Integer, Shift As Integer)
    
    '// for storing texts in grid
    '////////////////////////////
    mfgDetail.Text = XVal
    
    '// for deleting whole text
    '//////////////////////////
    If KeyCode = vbKeyDelete Then
        mfgDetail.Text = ""
        XVal = ""
    End If
End Sub

Private Sub mfgDetail_RowColChange()
    '// for clearing variable when row or column changes
    '///////////////////////////////////////////////////
    XVal = ""
End Sub
