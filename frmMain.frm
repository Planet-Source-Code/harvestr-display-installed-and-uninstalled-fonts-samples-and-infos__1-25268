VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Font Infos"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8625
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   8625
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   375
      Left            =   7800
      TabIndex        =   40
      Top             =   5175
      Visible         =   0   'False
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   661
      _Version        =   393216
      Rows            =   1
      FixedRows       =   0
      FixedCols       =   0
   End
   Begin VB.TextBox txtCustom 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   5265
      MaxLength       =   255
      TabIndex        =   4
      Text            =   "You can enter your own text here."
      Top             =   4455
      Width           =   3090
   End
   Begin VB.ListBox lstHidden 
      Height          =   255
      Left            =   6210
      Sorted          =   -1  'True
      TabIndex        =   36
      Top             =   5415
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.ListBox lstFiles 
      Height          =   255
      Left            =   6210
      TabIndex        =   35
      Top             =   5160
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton btnAbout 
      Height          =   435
      Left            =   7665
      Picture         =   "frmMain.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "About Font Infos"
      Top             =   90
      Width           =   435
   End
   Begin VB.CommandButton btnUnits 
      Height          =   435
      Left            =   75
      Picture         =   "frmMain.frx":0702
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Toggle inches/mm"
      Top             =   75
      Width           =   435
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   285
      Left            =   8310
      TabIndex        =   8
      Top             =   5745
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   503
      _Version        =   393216
      Value           =   20
      Max             =   72
      Min             =   6
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtSize 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   285
      Left            =   7980
      MaxLength       =   2
      TabIndex        =   7
      Text            =   "20"
      Top             =   5760
      Width           =   345
   End
   Begin VB.CheckBox chkItalic 
      Caption         =   "italic"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6195
      TabIndex        =   6
      Top             =   5895
      Width           =   930
   End
   Begin VB.CheckBox chkBold 
      Caption         =   "bold"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6195
      TabIndex        =   5
      Top             =   5655
      Width           =   825
   End
   Begin VB.CommandButton btnQuit 
      Height          =   435
      Left            =   8175
      Picture         =   "frmMain.frx":0E04
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Exit Program"
      Top             =   90
      Width           =   435
   End
   Begin VB.Frame frmSample 
      Caption         =   " Sample Text "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1140
      Left            =   60
      TabIndex        =   13
      Top             =   5235
      Width           =   6075
      Begin VB.Label lblSample 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   840
         Left            =   150
         TabIndex        =   14
         Top             =   225
         Width           =   5760
      End
   End
   Begin VB.CommandButton btnOpenFont 
      Height          =   435
      Left            =   585
      Picture         =   "frmMain.frx":1506
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Select Folder for Uninstalled Fonts"
      Top             =   75
      Width           =   435
   End
   Begin VB.ListBox lstOtherFonts 
      Height          =   4155
      Left            =   2580
      TabIndex        =   10
      Top             =   945
      Width           =   2430
   End
   Begin VB.ListBox lstSystemFonts 
      Height          =   4155
      Left            =   60
      Sorted          =   -1  'True
      TabIndex        =   9
      Top             =   945
      Width           =   2430
   End
   Begin VB.Line Line5 
      X1              =   5115
      X2              =   8490
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Label lblSize 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   4
      Left            =   6975
      TabIndex        =   39
      Top             =   4785
      Width           =   780
   End
   Begin VB.Label lblUnit 
      BackStyle       =   0  'Transparent
      Caption         =   "mm"
      Height          =   270
      Index           =   4
      Left            =   7815
      TabIndex        =   38
      Top             =   4785
      Width           =   615
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Custom Text Length"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   37
      Top             =   4185
      Width           =   3300
   End
   Begin VB.Label lblPath 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   285
      Left            =   1065
      TabIndex        =   34
      Top             =   165
      Width           =   6300
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00C0C0C0&
      X1              =   5130
      X2              =   7485
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00404040&
      X1              =   5115
      X2              =   8490
      Y1              =   2205
      Y2              =   2205
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Average Sizes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5265
      TabIndex        =   33
      Top             =   2235
      Width           =   3060
   End
   Begin VB.Label Label10 
      Caption         =   "Size"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   7485
      TabIndex        =   32
      Top             =   5790
      Width           =   405
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0C0C0&
      X1              =   5115
      X2              =   8505
      Y1              =   5085
      Y2              =   5085
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   8490
      X2              =   8490
      Y1              =   960
      Y2              =   5100
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Font Infos and Stats"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   5070
      TabIndex        =   31
      Top             =   705
      Width           =   3510
   End
   Begin VB.Label lblFileName 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   285
      Left            =   5265
      TabIndex        =   30
      Top             =   1800
      Width           =   3090
   End
   Begin VB.Label lblFontName 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   285
      Left            =   5265
      TabIndex        =   29
      Top             =   1230
      Width           =   3105
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "File Name"
      Height          =   315
      Left            =   5280
      TabIndex        =   28
      Top             =   1590
      Width           =   2100
   End
   Begin VB.Label lblUnit 
      BackStyle       =   0  'Transparent
      Caption         =   "mm"
      Height          =   270
      Index           =   3
      Left            =   7815
      TabIndex        =   26
      Top             =   3630
      Width           =   615
   End
   Begin VB.Label lblSize 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   3
      Left            =   6975
      TabIndex        =   25
      Top             =   3630
      Width           =   780
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "All Characters"
      Height          =   255
      Left            =   5160
      TabIndex        =   24
      Top             =   3630
      Width           =   2145
   End
   Begin VB.Label lblUnit 
      BackStyle       =   0  'Transparent
      Caption         =   "mm"
      Height          =   270
      Index           =   2
      Left            =   7830
      TabIndex        =   23
      Top             =   3255
      Width           =   630
   End
   Begin VB.Label lblSize 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   2
      Left            =   6975
      TabIndex        =   22
      Top             =   3240
      Width           =   780
   End
   Begin VB.Label lblUnit 
      BackStyle       =   0  'Transparent
      Caption         =   "mm"
      Height          =   270
      Index           =   1
      Left            =   7830
      TabIndex        =   21
      Top             =   2880
      Width           =   645
   End
   Begin VB.Label lblSize 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   6975
      TabIndex        =   20
      Top             =   2865
      Width           =   780
   End
   Begin VB.Label lblUnit 
      BackStyle       =   0  'Transparent
      Caption         =   "mm"
      Height          =   270
      Index           =   0
      Left            =   7845
      TabIndex        =   19
      Top             =   2505
      Width           =   630
   End
   Begin VB.Label lblSize 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   0
      Left            =   6975
      TabIndex        =   18
      Top             =   2505
      Width           =   780
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Numbers (0-9)"
      Height          =   255
      Left            =   5160
      TabIndex        =   17
      Top             =   3255
      Width           =   2190
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Lower Case Letters (a-z)"
      Height          =   255
      Left            =   5145
      TabIndex        =   16
      Top             =   2880
      Width           =   1905
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Upper Case Letters (A-Z)"
      Height          =   255
      Left            =   5145
      TabIndex        =   15
      Top             =   2505
      Width           =   1950
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Other Available Fonts"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   2580
      TabIndex        =   12
      Top             =   705
      Width           =   2430
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Installed Fonts"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   60
      TabIndex        =   11
      Top             =   705
      Width           =   2430
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Font Name"
      Height          =   255
      Left            =   5280
      TabIndex        =   27
      Top             =   1020
      Width           =   2100
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00999999&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      FillColor       =   &H00CCCCCC&
      FillStyle       =   0  'Solid
      Height          =   4155
      Left            =   5115
      Top             =   945
      Width           =   3390
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CurrentFont As String ' used to Store the selected Uninstalled Font
Dim InchUnits As Boolean ' used to store measurement Unit (True = inches, False = millimeters)
Dim BoldOn As Boolean, ItalicOn As Boolean ' used as "flags" for Font attributs
Dim FontPath As String ' used to store Uninstalled Fonts Folder
Dim ListDeSelect As Boolean

Private Sub GetSystemFonts()
' get all the fonts installed in Windows

lstSystemFonts.Clear 'Reset list
For i = 0 To Screen.FontCount - 1 'Cycle through every Font
    lstSystemFonts.AddItem Screen.Fonts(i) 'Add Font Name to List
Next i

End Sub

Private Sub GetOtherFonts()
Dim DirString As String ' used to list all TTF and FON files
Dim tmpStrg As String ' used to list all TTF and FON files
Dim FontFileName As String ' used to list all TTF and FON files
Dim tmpFont As String ' used to store uninstalled Font filename
Dim Result As Long ' Used for API Call
Dim FontFound As Boolean

lstOtherFonts.Clear 'Reset the list

' get all the font files (.ttf and .fon) in the FontPath folder
' and store them in the lstFiles hidden listbox
DirString = FontPath + "*.TTF" ' dir all TTF files
lstFiles.Clear ' clear the hidden listbox
tmpStrg = Dir$(DirString)
If tmpStrg <> "" Then
    FontFileName = tmpStrg ' one TTF file is found
Else
    GoTo Exit_GetOtherFonts1 ' no TTF is found
End If
lstFiles.AddItem FontFileName ' add found file to listbox
tmpStrg = Dir$
While Len(tmpStrg) > 0 ' cycle through all files
    FontFileName = tmpStrg
    lstFiles.AddItem FontFileName ' add file to listbox
    tmpStrg = Dir$
Wend
Exit_GetOtherFonts1:

DirString = FontPath + "*.FON" ' dir all FON files
tmpStrg = Dir$(DirString)
If tmpStrg <> "" Then
    FontFileName = tmpStrg ' one TTF file is found
Else
    GoTo Exit_GetOtherFonts2 ' no TTF is found
End If
lstFiles.AddItem FontFileName ' add found file to list
tmpStrg = Dir$
While Len(tmpStrg) > 0 ' cycle through all files
    FontFileName = tmpStrg
    lstFiles.AddItem FontFileName
    tmpStrg = Dir$
Wend
Exit_GetOtherFonts2:

If lstFiles.ListCount = 0 Then
    lstOtherFonts.AddItem "No Font found" ' inform user no font was found
    lstOtherFonts.Enabled = False ' disable the list (no need to click, it's empty !)
Else
    ' We now have to retrieve the Font name from the Font filename.
    ' Getting it from within the file itself is a bit hard, so
    ' we will use a litte trick.
    ' We load all System fonts in the hidden listbox (lstHidden)
    ' then we temporary add the uninstalled font to the system
    '(using the AddFontResource API), and we fill the hidden
    ' listbox with the installed fonts again. The result is
    ' the new font we added will appear only once in the list.
    ' As we set the Sorted Property of the hidden listbox on,
    ' we only have to check the list, and find where two lines
    ' are not matching, and we'll now we've found the new
    ' Font. We will store the Font name and filename in an array
    ' I used a MSFlexGrid as array (because I got lost with redim preserve function)
    ' MSFlexGrid is a VB component.
    ' There's surely a better way to do it, but I don't know how.
    MSFlexGrid1.Clear ' clear the hidden grid
    MSFlexGrid1.Rows = 1
    For i = 0 To lstFiles.ListCount - 1
        lstHidden.Clear
        For j = 0 To Screen.FontCount - 1 'Cycle through every Font
            lstHidden.AddItem Screen.Fonts(j) 'Add Font Name to List
        Next j
        Result = AddFontResource(FontPath + lstFiles.List(i))
        For j = 0 To Screen.FontCount - 1 'Cycle through every Font
            lstHidden.AddItem Screen.Fonts(j) 'Add Font Name to List
        Next j
        Result = RemoveFontResource(FontPath + lstFiles.List(i))
        For j = 0 To lstHidden.ListCount - 1 Step 2
            If lstHidden.List(j) <> lstHidden.List(j + 1) Then
                If i > 0 Then MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
                MSFlexGrid1.TextMatrix(i, 0) = lstHidden.List(j)
                MSFlexGrid1.TextMatrix(i, 1) = lstFiles.List(i)
                lstOtherFonts.AddItem lstHidden.List(j)
                Exit For
            End If
        Next j
    Next i
End If
End Sub

Private Sub CalcSizes()

Dim OriginalFont As String ' used to store default Form Font Name
Dim OriginalSize As Integer ' used to store default Form Font Size
Dim OriginalBold As Boolean ' used to store default Form Font Bold
Dim OriginalItalic As Boolean ' used to store default Form Font Italic

Dim StringLength As Double ' used to calculate the full string length
Dim UpperSize As Double ' used to calculate Upper Case Letters Average Size
Dim LowerSize As Double ' used to calculate Lower Case Letters Average Size
Dim NumberSize As Double ' used to calculate Numbers Average Size
Dim FullSize As Double ' used to calculate All Chars. Average Size
Dim SampleString As String ' used in all calculations above

' Store Form font values
OriginalFont = frmMain.FontName
OriginalSize = frmMain.FontSize
OriginalBold = frmMain.FontBold
OriginalItalic = frmMain.FontItalic

' Set new Form Font Attributes
frmMain.FontName = lblFontName.Caption
frmMain.FontSize = Val(Trim(txtSize.Text))
frmMain.FontBold = BoldOn
frmMain.FontItalic = ItalicOn

' Set the unit used (inches or millimeters)
If InchUnits = True Then
    frmMain.ScaleMode = vbInches
Else
    frmMain.ScaleMode = vbMillimeters
End If

' Get the length of the string ABCD...WXYZ (Upper Case)
SampleString = "" ' empty the string
For i = 1 To 26 ' 26 letters (A to Z)
    SampleString = SampleString + Chr(64 + i) ' ascii value for A is 65, for B 66, ...
Next i
StringLength = frmMain.TextWidth(SampleString)
UpperSize = StringLength / 26 ' Average size for 1 char.
UpperSize = Int(UpperSize * 100) / 100 ' rounded to .xx value
lblSize(0).Caption = Str(UpperSize) ' display value

' Get the length of the string abcd...wxyz (Lower Case)
SampleString = "" ' empty the string
For i = 1 To 26 ' 26 letters (a to z)
    SampleString = SampleString + Chr(96 + i) ' ascii value for a is 97, for b 98, ...
Next i
StringLength = frmMain.TextWidth(SampleString)
LowerSize = StringLength / 26 ' Average size for 1 char.
LowerSize = Int(LowerSize * 100) / 100 ' rounded to .xx value
lblSize(1).Caption = Str(LowerSize) ' display value

' Get the length of the string 0123456789 (Numbers)
SampleString = "0123456789"
StringLength = frmMain.TextWidth(SampleString)
NumberSize = StringLength / 10 ' Average size for 1 char.
NumberSize = Int(NumberSize * 100) / 100 ' rounded to .xx value
lblSize(2).Caption = Str(NumberSize) ' display value

' Get the length of all the PRINTABLE chars (224/255).
SampleString = "" ' empty the string
For i = 32 To 255 ' from space to last possible char
    SampleString = SampleString + Chr(i)
Next i
StringLength = frmMain.TextWidth(SampleString)
FullSize = StringLength / 224 ' Average size for 1 char.
FullSize = Int(FullSize * 100) / 100 ' rounded to .xx value
lblSize(3).Caption = Str(FullSize) ' display value

' restore original Form Font values
frmMain.FontName = OriginalFont
frmMain.FontSize = OriginalSize
frmMain.FontBold = OriginalBold
frmMain.FontItalic = OriginalItalic

' Call the Custom Text size calculation SUB (refer below)
CalcCustomSize
End Sub

Private Sub CalcCustomSize()
Dim SampleString As String
Dim CustomSize As Double ' used to calculate Custom Text total size

Dim OriginalFont As String ' used to store default Form Font Name
Dim OriginalSize As Integer ' used to store default Form Font Size
Dim OriginalBold As Boolean ' used to store default Form Font Bold
Dim OriginalItalic As Boolean ' used to store default Form Font Italic

' Store Form font values
OriginalFont = frmMain.FontName
OriginalSize = frmMain.FontSize
OriginalBold = frmMain.FontBold
OriginalItalic = frmMain.FontItalic

' Set new Form Font Attributes
frmMain.FontName = lblFontName.Caption
frmMain.FontSize = Val(Trim(txtSize.Text))
frmMain.FontBold = BoldOn
frmMain.FontItalic = ItalicOn

' Set the unit used (inches or millimeters)
If InchUnits = True Then
    frmMain.ScaleMode = vbInches
Else
    frmMain.ScaleMode = vbMillimeters
End If

' Get the length of the Custom Text size.
SampleString = txtCustom.Text ' Fill string with custom text
CustomSize = frmMain.TextWidth(SampleString)
CustomSize = Int(CustomSize * 100) / 100 ' rounded to .xx value
lblSize(4).Caption = Str(CustomSize) ' display value

' restore original Form Font values
frmMain.FontName = OriginalFont
frmMain.FontSize = OriginalSize
frmMain.FontBold = OriginalBold
frmMain.FontItalic = OriginalItalic

End Sub

Private Sub DisplaySample()
Dim DisplayString As String ' the string that will be displayed as sample
Dim StringHeight As Integer ' height of text
Dim NewTop As Integer ' label new top position
Dim NewHeight As Integer ' label new height
Dim MaxTop As Integer ' maximum label top position
Dim MaxHeight As Integer ' maximum label height

' set maximum positions
MaxTop = 150
MaxHeight = 850

Dim OriginalFont As String ' used to store default Form Font Name
Dim OriginalSize As Integer ' used to store default Form Font Size
Dim OriginalBold As Boolean ' used to store default Form Font Bold
Dim OriginalItalic As Boolean ' used to store default Form Font Italic

' set Font properties for the Sample Label
lblSample.FontName = lblFontName.Caption
lblSample.FontSize = Val(Trim(txtSize.Text))
lblSample.FontBold = BoldOn
lblSample.FontItalic = ItalicOn

' set label caption as font name
lblSample.Caption = lblFontName.Caption

' Depending the height of the label, center it in the Frame
' First: get the text height (based on form.TextWidth property)
'        refer to CalcSize Sub for more infos

' Store Form font values
OriginalFont = frmMain.FontName
OriginalSize = frmMain.FontSize
OriginalBold = frmMain.FontBold
OriginalItalic = frmMain.FontItalic

' Set new Form Font Attributes
frmMain.FontName = lblFontName.Caption
frmMain.FontSize = Val(Trim(txtSize.Text))
frmMain.FontBold = BoldOn
frmMain.FontItalic = ItalicOn

' set scalemode to twips (for position purpose)
frmMain.ScaleMode = vbTwips

' get text height then add 2 pixels up and down to it (so 4 * 15 = 60 twips)
StringHeight = frmMain.TextHeight(lblSample.Caption)
StringHeight = StringHeight + 60

' if text height < label max height then move the label
' to center it and redefine the label height
If StringHeight < MaxHeight Then
    NewHeight = StringHeight
    NewTop = (MaxHeight - NewHeight) / 2 ' algorythm for centering
    NewTop = Int(NewTop) ' we need an integer...
    lblSample.Top = 150 + NewTop
    lblSample.Height = NewHeight
End If

' restore original Form Font values
frmMain.FontName = OriginalFont
frmMain.FontSize = OriginalSize
frmMain.FontBold = OriginalBold
frmMain.FontItalic = OriginalItalic

End Sub

Private Sub btnAbout_Click()
' Display About Box
frmAbout.Show vbModal
End Sub

Private Sub btnOpenFont_Click()
Dim ResultTTF As String ' used to check if TTF file is found in folder
Dim ResultFON As String ' used to check if FON file is found in folder

' This Sub (in browse.bas) will bring up the Browse for Folder window
' Refer to browse_readme.txt for more infos
Display Me.hWnd, " " '" " is because we don't add a title to the browse window

If successful = False Then Exit Sub ' user cancelled folde browse

If successful = True Then ' user selected a folder
    ' be sure the folder name ends with \
    If Right(folderName, 1) <> "\" Then folderName = folderName + "\"
End If

ResultTTF = Dir(folderName + "*.TTF") ' will be empty if no TTF file found
ResultFON = Dir(folderName + "*.FON") ' will be empty if no FON file found
If Len(ResultTTF) = 0 And Len(ResultFON) = 0 Then ' no Font file found
    Beep ' a little noise ;)
    MsgBox "Sorry, No Font File was found in this Folder!", vbOKOnly + vbExclamation, "ERROR"
    Call btnOpenFont_Click ' reopen the browse for folder
End If

' don't allow user to select the Windows' Fonts folder
' It would work, but take long time for nothing... no Font
' would be added to the list.
' See the GetWinPath function in modMain.bas (uses API)
If UCase(folderName) = UCase(GetWinPath + "\fonts\") Then
    Beep ' a little noise ;)
    MsgBox "Sorry, you can't select the Windows Fonts folder!", vbOKOnly + vbExclamation, "ERROR"
    Call btnOpenFont_Click ' reopen the browse for folder
End If

' Remove current Uninstalled Font from memory, if any
If Len(CurrentFont) > 0 Then RemoveFontResource (FontPath + CurrentFont)

' refresh the Uninstalled Fonts list
FontPath = folderName ' set the UnInstalled Fonts folder
lblPath.Caption = FontPath ' display new path
GetOtherFonts ' refer to the sub above

' Deselect the Uninstalled Fonts lists
'     Because the .selected property generates a Click event
'     we use a flag (ListDeSelect)to ignore this click.
ListDeSelect = True ' avoid Click event in lstSystemFonts
If lstOtherFonts.ListIndex > -1 Then lstOtherFonts.Selected(lstOtherFonts.ListIndex) = False
ListDeSelect = False ' reset flag

End Sub

Private Sub btnQuit_Click()
' Remove current Uninstalled Font from memory, if any
If Len(CurrentFont) > 0 Then RemoveFontResource (FontPath + CurrentFont)

End
End Sub

Private Sub btnUnits_Click()

' If unit is mm, switch to inches, and vice-versa
If InchUnits = False Then
    InchUnits = True
    For i = 0 To 4 ' display "inches" in the 5 labels
        lblUnit(i).Caption = "inches"
    Next i
Else
    InchUnits = False
    For i = 0 To 4 ' display "mm" in the 5 labels
        lblUnit(i).Caption = "mm"
    Next i
End If

' Recalculate all Sizes
CalcSizes
End Sub

Private Sub chkBold_Click()
' Values have changed, so update:
If chkBold.Value = 1 Then
    BoldOn = True
Else
    BoldOn = False
End If

' Call the Sub that will display the Sample text (please refer above)
If lstSystemFonts.ListIndex <> -1 Or lstOtherFonts.ListIndex <> -1 Then
    DisplaySample
Else
    Exit Sub
End If

' Call the Sub that will calculate Average Sizes (please refer above)
CalcSizes

End Sub

Private Sub chkItalic_Click()
' Values have changed, so update:
If chkItalic.Value = 1 Then
    ItalicOn = True
Else
    ItalicOn = False
End If

' Call the Sub that will display the Sample text (please refer above)
If lstSystemFonts.ListIndex <> -1 Or lstOtherFonts.ListIndex <> -1 Then
    DisplaySample
Else
    Exit Sub
End If

' Call the Sub that will calculate Average Sizes (please refer above)
CalcSizes

End Sub

Private Sub Command1_Click()
RemoveFontResource (FontPath + "Bladrmf_.ttf")
RemoveFontResource (FontPath + "aapex.ttf")
RemoveFontResource (FontPath + "american.ttf")


End Sub

Private Sub Form_Load()

' Call the sub that fill first list with installed fonts
GetSystemFonts

' set default Font Size to 20
txtSize.Text = 20

' set default Uninstalled Fonts Folder to App path
FontPath = App.path
If Right(FontPath, 1) <> "\" Then FontPath = FontPath + "\"
lblPath.Caption = FontPath

' grey the Font File Name label (filename are used for
' uninstalled Fonts only)
lblFileName.BackColor = RGB(102, 102, 102)

' Call the sub that fill the Uninstalled Fonts list
GetOtherFonts

' Display first installed font as default
lblFontName.Caption = lstSystemFonts.List(0)
lblFileName.Caption = "" ' empty and Grey Font file name label
lblFileName.BackColor = RGB(102, 102, 102)
DisplaySample ' Call the Sub that will display the Sample text (please refer above)
CalcSizes ' Call the Sub that will calculate Average Sizes (please refer above)
Unload frmOpen
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
' Remove current Uninstalled Font from memory, if any
If Len(CurrentFont) > 0 Then RemoveFontResource (FontPath + CurrentFont)

End Sub

Private Sub lstOtherFonts_Click()

' if click event is generated by .selected property, ignore it
'     (also refer to lstSystemFonts_Click)
If ListDeSelect = True Then Exit Sub

' Deselect the system fonts list
'     Because the .selected property generates a Click event
'     we use a flag (ListDeSelect)to ignore this click.
ListDeSelect = True ' avoid Click event in lstSystemFonts
If lstSystemFonts.ListIndex > -1 Then lstSystemFonts.Selected(lstSystemFonts.ListIndex) = False
ListDeSelect = False ' reset flag

' Remove previous Uninstalled Font from System, if any
If Len(CurrentFont) > 0 Then RemoveFontResource (FontPath + CurrentFont)

' Show Font name, if none selected, exit sub
If Len(lblFontName.Caption) = 0 Then Exit Sub
lblFontName.Caption = lstOtherFonts.List(lstOtherFonts.ListIndex)

' Show Font filename and set label's background to white
lblFileName.BackColor = RGB(255, 255, 255)
lblFileName.Caption = MSFlexGrid1.TextMatrix(lstOtherFonts.ListIndex, 1)
CurrentFont = lblFileName.Caption

' Add Font to system
AddFontResource (FontPath + CurrentFont)

' Call the Sub that will display the Sample text (please refer above)
DisplaySample

' Call the Sub that will calculate Average Sizes (please refer above)
CalcSizes

End Sub

Private Sub lstSystemFonts_Click()
' Remove current Uninstalled Font from memory, if any
If Len(CurrentFont) > 0 Then RemoveFontResource (FontPath + CurrentFont)

' if click event is generated by .selected property, ignore it
'     (also refer to lstOtherFonts_Click)
If ListDeSelect = True Then Exit Sub

' Deselect the system fonts list
'     Because the .selected property generates a Click event
'     we use a flag (ListDeSelect)to ignore this click.
ListDeSelect = True ' avoid Click event in lstSystemFonts
If lstOtherFonts.ListIndex > -1 Then lstOtherFonts.Selected(lstOtherFonts.ListIndex) = False
ListDeSelect = False ' reset flag

' Display Font Name in the Stats label
lblFontName.Caption = lstSystemFonts.List(lstSystemFonts.ListIndex)

' empty and Grey Font file name label
lblFileName.Caption = ""
lblFileName.BackColor = RGB(102, 102, 102)

' Call the Sub that will display the Sample text (please refer above)
DisplaySample

' Call the Sub that will calculate Average Sizes (please refer above)
CalcSizes

End Sub

Private Sub txtCustom_Change()

' If Custom Text is not empty, calculate new Custom size
If Len(txtCustom.Text) > 0 Then CalcCustomSize

End Sub

Private Sub txtSize_Change()
' calculate the averages Sizes and change Sample text
' automatically when size value is changed, unless Size
' is not within the permitted range (6 to 72)

' NOTE: txtSize MaxLength property is set to 2, so there can't
' be 3 digits numbers.

' because minimum value is set to 6, if first digit entered
' is <6 then wait for next one (must be a 2 digits number)
If Len(txtSize.Text) < 2 Then
    If Val(Trim(txtSize.Text)) < 6 Then
        Exit Sub ' no Average Sizes and no Sample displayed
    End If
End If

' because maximum value is set to 72, if first digit entered
' is 7 then second digit can't be > to 2
If Len(txtSize.Text) = 2 Then
    If Val(Trim(Left(txtSize.Text, 1))) = 7 Then
        If Val(Trim(Right(txtSize.Text, 1))) > 2 Then
            Beep
            txtSize.Text = Left(txtSize.Text, 1) 'remove second digit
            txtSize.SelStart = 1 'set cursor position after first digit
            Exit Sub ' no Average Sizes and no Sample displayed
        End If
    End If
End If

' set the updown value to the txtsize value
UpDown1.Value = Val(txtSize.Text)

' Call the Sub that will display the Sample text (please refer above)
If lstSystemFonts.ListIndex <> -1 Or lstOtherFonts.ListIndex <> -1 Then
    DisplaySample
Else
    Exit Sub
End If

' Call the Sub that will calculate Average Sizes (please refer above)
CalcSizes

End Sub

Private Sub txtSize_KeyPress(KeyAscii As Integer)

' Numbers only in the textbox
Select Case KeyAscii
    Case 48 To 57 ' allow characters from 0 to 9
    Case 8        ' allow backspace
    Case Else     ' nothing else permitted
        Beep
        KeyAscii = 0 'cancel the pressed key
End Select

End Sub

Private Sub UpDown1_Change()
' UpDown Control is part of Microsoft Common Control 2
' wich is a standard control coming with VB6.
' To use it in your other programs, open the Project menu
' and select Components, then check Microsoft Common Control 2
' in the components list.
' To change default, min and max values, click Custom in the
' Control properties and see in third tab.

' Put UpDown1 value in txtSize
txtSize.Text = UpDown1.Value
End Sub
