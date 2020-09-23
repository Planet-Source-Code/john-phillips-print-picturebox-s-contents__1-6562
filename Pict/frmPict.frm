VERSION 5.00
Begin VB.Form frmPict 
   Caption         =   "Afo Worksheet Printout"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   6930
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   2055
      Begin VB.CommandButton Command1 
         Caption         =   "Print"
         Height          =   255
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Done"
         Height          =   255
         Left            =   960
         TabIndex        =   3
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.VScrollBar vBar 
      Height          =   5895
      Left            =   6600
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.HScrollBar hBar 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   6120
      Width           =   6495
   End
   Begin VB.PictureBox OuterPict 
      BackColor       =   &H8000000E&
      Height          =   6015
      Left            =   0
      ScaleHeight     =   5955
      ScaleWidth      =   6435
      TabIndex        =   5
      Top             =   0
      Width           =   6495
      Begin VB.PictureBox InnerPict 
         BackColor       =   &H8000000E&
         Height          =   5895
         Left            =   0
         Picture         =   "frmPict.frx":0000
         ScaleHeight     =   5835
         ScaleWidth      =   6315
         TabIndex        =   7
         Top             =   0
         Width           =   6375
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   3840
            TabIndex        =   13
            Text            =   "Text1"
            Top             =   360
            Width           =   2055
         End
         Begin VB.ListBox List1 
            Height          =   840
            Left            =   1560
            TabIndex        =   12
            Top             =   4800
            Width           =   1695
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   2520
            TabIndex        =   11
            Text            =   "Combo1"
            Top             =   4200
            Width           =   2895
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H80000009&
            Caption         =   "Option1"
            Height          =   255
            Left            =   0
            TabIndex        =   9
            Top             =   4320
            Width           =   2055
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H80000009&
            Caption         =   "Check1"
            Height          =   255
            Left            =   0
            TabIndex        =   8
            Top             =   4080
            Width           =   1695
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "by John Phillips"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   3000
            Width           =   4335
         End
         Begin VB.Line Line1 
            X1              =   1440
            X2              =   4920
            Y1              =   4680
            Y2              =   4680
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Planet-Source-Code Print Picturebox Contents and controls"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   1575
            Left            =   0
            TabIndex        =   10
            Top             =   1200
            Width           =   4455
         End
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H8000000E&
      Height          =   2295
      Left            =   3720
      ScaleHeight     =   2235
      ScaleWidth      =   2955
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   3015
   End
End
Attribute VB_Name = "frmPict"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
   ' John Phillips, MCP
   ' VBJack1@aol.com
   ' Parts of this program were takein from source code
   ' on planet-source-code.com (Picturebox scroll bars)
   
   Private Const twipFactor = 1440
   Private Const WM_PAINT = &HF
   Private Const WM_PRINT = &H317
   Private Const PRF_CLIENT = &H4&    ' Draw the window's client area.
   Private Const PRF_CHILDREN = &H10& ' Draw all visible child windows.
   Private Const PRF_OWNED = &H20&    ' Draw all owned windows.

   Private Declare Function SendMessage Lib "user32" Alias _
      "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
      ByVal wParam As Long, ByVal lParam As Long) As Long
   
  Private Sub Command1_Click()
      Dim sWide As Single, sTall As Single
      Dim rv As Long

      Me.ScaleMode = vbTwips   ' default
      sWide = 8.5  ' set the size of the area to print in the picturebox
      sTall = 11   ' or 14, etc.
      Me.Width = twipFactor * sWide
      Me.Height = twipFactor * sTall
      With InnerPict
         .Top = 0
         .Left = 0
         .Width = twipFactor * sWide
         .Height = twipFactor * sTall
      End With
      With Picture2
         .Top = 0
         .Left = 0
         .Width = twipFactor * sWide
         .Height = twipFactor * sTall
      End With
      Me.Visible = True
      DoEvents
  
      InnerPict.SetFocus ' Set focus on the main PictureBox ie. where the controls and picture are
      Picture2.AutoRedraw = True 'Set Autoredraw to true
      
      ' send the contents of InnerPict to PictureBox2
      ' Essentially take a snapshot of the entire contents
      ' of the Picturebox we want to to print - including
      ' controls using sendmessage
      
      rv = SendMessage(InnerPict.hwnd, WM_PAINT, Picture2.hDC, 0)
      rv = SendMessage(InnerPict.hwnd, WM_PRINT, Picture2.hDC, _
      PRF_CHILDREN + PRF_CLIENT + PRF_OWNED)
      Picture2.Picture = Picture2.Image
   
      Picture2.AutoRedraw = False
      Picture2.Visible = True
      
      ' After setting the contents of Innerpict to Picture2
      ' we now need to set you the printer to print the contents
      ' of Picture2
      Printer.PrintQuality = 300
      Printer.Print ""
      'Printer.DrawMode = 7
      ' use the printer Paintpicture method to print the
      ' contents
      Printer.PaintPicture Picture2.Picture, 0, 0
      Printer.EndDoc ' close the print session
      End Sub

Private Sub Command2_Click()
Unload Me
End
End Sub

      Private Sub Form_Load()
       Me.Show
       Command1.Caption = "Print Form"
      End Sub
 
 Private Sub SetScrollBars()
    ' Set scroll bar properties.
    vBar.Min = 0
    vBar.Max = OuterPict.ScaleHeight - InnerPict.Height
    vBar.LargeChange = OuterPict.ScaleHeight
    vBar.SmallChange = OuterPict.ScaleHeight / 5
    
    hBar.Min = 0
    hBar.Max = OuterPict.ScaleWidth - InnerPict.Width
    hBar.LargeChange = OuterPict.ScaleWidth
    hBar.SmallChange = OuterPict.ScaleWidth / 5
End Sub
Private Sub Form_Resize()
Dim got_wid As Single
Dim got_hgt As Single
Dim need_wid As Single
Dim need_hgt As Single
Dim need_hbar As Boolean
Dim need_vbar As Boolean

    If WindowState = vbMinimized Then Exit Sub

    need_wid = InnerPict.Width + (OuterPict.Width - OuterPict.ScaleWidth)
    need_hgt = InnerPict.Height + (OuterPict.Height - OuterPict.ScaleHeight)
    got_wid = ScaleWidth
    got_hgt = ScaleHeight

    ' See which scroll bars we need.
    need_hbar = (need_wid > got_wid)
    If need_hbar Then got_hgt = got_hgt - hBar.Height

    need_vbar = (need_hgt > got_hgt)
    If need_vbar Then
        got_wid = got_wid - vBar.Width
        If Not need_hbar Then
            need_hbar = (need_wid > got_wid)
            If need_hbar Then got_hgt = got_hgt - hBar.Height
        End If
    End If

    OuterPict.Move 0, 0, got_wid, got_hgt

    If need_hbar Then
        hBar.Move 0, got_hgt, got_wid
        hBar.Visible = True
    Else
        hBar.Visible = False
    End If

    If need_vbar Then
        vBar.Move got_wid, 0, vBar.Width, got_hgt
        vBar.Visible = True
    Else
        vBar.Visible = False
    End If
    
    SetScrollBars
End Sub

Private Sub HBar_Change()
    InnerPict.Left = hBar.Value
End Sub


Private Sub HBar_Scroll()
    InnerPict.Left = hBar.Value
End Sub


Private Sub Label11_Click(Index As Integer)
End Sub

Private Sub VBar_Change()
    InnerPict.Top = vBar.Value
End Sub


Private Sub VBar_Scroll()
    InnerPict.Top = vBar.Value
End Sub





