VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "agentctl.dll"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form srhMain 
   BackColor       =   &H80000003&
   Caption         =   "SEARCH "
   ClientHeight    =   7995
   ClientLeft      =   1530
   ClientTop       =   1935
   ClientWidth     =   9105
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":1CCA
   ScaleHeight     =   7995
   ScaleWidth      =   9105
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   300
      Top             =   4020
   End
   Begin VB.DirListBox Dir1 
      Height          =   1215
      Left            =   720
      TabIndex        =   6
      Top             =   2130
      Width           =   2655
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   720
      TabIndex        =   5
      Top             =   1590
      Width           =   2655
   End
   Begin ComctlLib.ListView ListView1 
      CausesValidation=   0   'False
      Height          =   7815
      Left            =   4080
      TabIndex        =   4
      Top             =   0
      Width           =   7880
      _ExtentX        =   13891
      _ExtentY        =   13785
      LabelEdit       =   1
      SortOrder       =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483641
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.FileListBox File1 
      Archive         =   0   'False
      Height          =   3405
      Left            =   7320
      System          =   -1  'True
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton cmdstop 
      BackColor       =   &H80000014&
      Caption         =   "Stop"
      Height          =   375
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtsrh 
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Top             =   510
      Width           =   2625
   End
   Begin VB.CommandButton cmdsrh 
      BackColor       =   &H80000014&
      Caption         =   "Search"
      Height          =   375
      Left            =   720
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label15 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Label15"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   120
      TabIndex        =   21
      Top             =   4200
      Width           =   3810
   End
   Begin AgentObjectsCtl.Agent Agent2 
      Left            =   8880
      Top             =   3840
   End
   Begin AgentObjectsCtl.Agent Agent1 
      Left            =   0
      Top             =   120
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   4440
      Width           =   3855
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   19
      Top             =   7320
      Width           =   1335
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Attributes : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   7320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label11 
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      Caption         =   "Label11"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   6840
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Label10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   6360
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   5880
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   5400
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   4920
      Width           =   3855
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000003&
      BackStyle       =   0  'Transparent
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
      Left            =   120
      TabIndex        =   12
      Top             =   3840
      Width           =   3765
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   375
      Left            =   720
      TabIndex        =   11
      Top             =   1560
      Width           =   3045
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   720
      TabIndex        =   10
      Top             =   3060
      Width           =   2955
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   720
      TabIndex        =   9
      Top             =   2760
      Width           =   2985
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   720
      TabIndex        =   8
      Top             =   2250
      Width           =   2985
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "All or part of file name"
      Height          =   375
      Left            =   1200
      TabIndex        =   7
      Top             =   180
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000004&
      FillColor       =   &H80000013&
      FillStyle       =   0  'Solid
      Height          =   3705
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   90
      Width           =   3615
   End
   Begin VB.Menu file 
      Caption         =   "Commands"
      Begin VB.Menu srh 
         Caption         =   "Search"
      End
      Begin VB.Menu stop 
         Caption         =   "Stop"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu adv_op 
      Caption         =   "Options"
      Begin VB.Menu srh_all_drv 
         Caption         =   "Search in all drives"
         Checked         =   -1  'True
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu srh_sub_fld 
         Caption         =   "Search in subfolders"
         Checked         =   -1  'True
      End
      Begin VB.Menu srh_arh 
         Caption         =   "Search in archives"
         Checked         =   -1  'True
      End
      Begin VB.Menu srh_hid_fil 
         Caption         =   "Search in hidden files"
         Checked         =   -1  'True
      End
      Begin VB.Menu srh_sys_fil 
         Caption         =   "Search in system files"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu ch_det 
      Caption         =   "View Details"
      Begin VB.Menu par_fld 
         Caption         =   "Parent folder"
         Checked         =   -1  'True
         Visible         =   0   'False
      End
      Begin VB.Menu size 
         Caption         =   "Size"
         Checked         =   -1  'True
      End
      Begin VB.Menu date_crd 
         Caption         =   "Date created"
         Checked         =   -1  'True
      End
      Begin VB.Menu date_acc 
         Caption         =   "Date last accessed"
         Checked         =   -1  'True
      End
      Begin VB.Menu date_mod 
         Caption         =   "Date last modified"
         Checked         =   -1  'True
      End
      Begin VB.Menu typ 
         Caption         =   "Type"
         Checked         =   -1  'True
      End
      Begin VB.Menu attr 
         Caption         =   "Attributes"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu sort 
      Caption         =   "Sort"
      Begin VB.Menu s_name 
         Caption         =   "name"
         Checked         =   -1  'True
      End
      Begin VB.Menu s_size 
         Caption         =   "Size"
         Checked         =   -1  'True
      End
      Begin VB.Menu s_date_crd 
         Caption         =   "Date created"
         Checked         =   -1  'True
      End
      Begin VB.Menu s_date_acc 
         Caption         =   "Date last accessed"
         Checked         =   -1  'True
      End
      Begin VB.Menu s_date_mod 
         Caption         =   "Date last modified"
         Checked         =   -1  'True
      End
      Begin VB.Menu s_par_fld 
         Caption         =   "Parent folder"
         Checked         =   -1  'True
      End
      Begin VB.Menu s_typ 
         Caption         =   "Type"
         Checked         =   -1  'True
      End
      Begin VB.Menu s_attr 
         Caption         =   "Attributes"
         Checked         =   -1  'True
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu asc 
         Caption         =   "Ascending"
         Checked         =   -1  'True
      End
      Begin VB.Menu dsc 
         Caption         =   "Descending"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuadop 
      Caption         =   "Advance Options"
      Begin VB.Menu pic_fil 
         Caption         =   "Search Picture files"
         Checked         =   -1  'True
      End
      Begin VB.Menu doc_fil 
         Caption         =   "Search Document files"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnusep5 
         Caption         =   "-"
      End
      Begin VB.Menu srh_by 
         Caption         =   "Search by"
         Begin VB.Menu typ1 
            Caption         =   "Type"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnusep4 
            Caption         =   "-"
         End
         Begin VB.Menu size1 
            Caption         =   "Size"
            Begin VB.Menu s_none 
               Caption         =   "None"
               Checked         =   -1  'True
            End
            Begin VB.Menu s_m_than 
               Caption         =   "More than"
               Checked         =   -1  'True
            End
            Begin VB.Menu s_l_than 
               Caption         =   "Less than"
               Checked         =   -1  'True
            End
            Begin VB.Menu s_eq_to 
               Caption         =   "Equal to"
               Checked         =   -1  'True
            End
         End
         Begin VB.Menu date_crd1 
            Caption         =   "Date created"
            Begin VB.Menu c_none 
               Caption         =   "None"
               Checked         =   -1  'True
            End
            Begin VB.Menu c_m_than 
               Caption         =   "More than"
               Checked         =   -1  'True
            End
            Begin VB.Menu c_l_than 
               Caption         =   "Less than"
               Checked         =   -1  'True
            End
            Begin VB.Menu c_eq_to 
               Caption         =   "Equal to"
               Checked         =   -1  'True
            End
         End
         Begin VB.Menu date_acc1 
            Caption         =   "Date last accessed"
            Begin VB.Menu a_none 
               Caption         =   "None"
               Checked         =   -1  'True
            End
            Begin VB.Menu a_m_than 
               Caption         =   "More than"
               Checked         =   -1  'True
            End
            Begin VB.Menu a_l_than 
               Caption         =   "Less than"
               Checked         =   -1  'True
            End
            Begin VB.Menu a_eq_to 
               Caption         =   "Equal to"
               Checked         =   -1  'True
            End
         End
         Begin VB.Menu date_mod1 
            Caption         =   "Date last modified"
            Begin VB.Menu m_none 
               Caption         =   "None"
               Checked         =   -1  'True
            End
            Begin VB.Menu m_m_than 
               Caption         =   "More than"
               Checked         =   -1  'True
            End
            Begin VB.Menu m_l_than 
               Caption         =   "Less than"
               Checked         =   -1  'True
            End
            Begin VB.Menu m_eq_to 
               Caption         =   "Equal to"
               Checked         =   -1  'True
            End
         End
         Begin VB.Menu attr1 
            Caption         =   "Attributes"
            Begin VB.Menu atr_none 
               Caption         =   "None"
               Checked         =   -1  'True
            End
            Begin VB.Menu atr_arh 
               Caption         =   "Archive"
               Checked         =   -1  'True
            End
            Begin VB.Menu atr_hid 
               Caption         =   "Hidden"
               Checked         =   -1  'True
            End
            Begin VB.Menu atr_sys 
               Caption         =   "System"
               Checked         =   -1  'True
            End
            Begin VB.Menu atr_read 
               Caption         =   "Read only"
               Checked         =   -1  'True
            End
         End
      End
      Begin VB.Menu mnusep6 
         Caption         =   "-"
      End
      Begin VB.Menu srh_speed 
         Caption         =   "Search speed"
         Begin VB.Menu sp_medium 
            Caption         =   "Normal"
            Checked         =   -1  'True
         End
         Begin VB.Menu sp_fast 
            Caption         =   "Fast"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu act_mer 
         Caption         =   "Activate Merlin"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "srhMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Brought to you by Ashish Kumar Methani
'   8th sem Comp. Science,RIT college,Raipur,
'   ashish_methani@yahoo.co.in

' Searches various files depending on their properties on the hard disk.
Option Explicit
Dim fso As New FileSystemObject
Dim fld As Folder
Dim counter, counter1, counter2, counter3, counter4 As Long
Dim i, j, k, lc, z, s, s_more, s_less, atr, attri, n, at, sort_var, hr As Integer
Dim searchflag, alldrv, pic, fl, fl1, fl2, sb, st As Boolean
Dim si, t_typ, sort_text, lit, lit1, txtsrh1 As String
Dim mod_date, a, a_more, a_less, c, c_more, c_less, m, m_more, m_less As Date
Dim li As ListItem
' For using microsoft agent merlin one has to include microsoft agent control 2.0 from references
Dim merlin As IAgentCtlCharacterEx
' Setting path for merlin (remember merlin.acs file must be in the same folder as of exe file
Const DATAPATH = ".\merlin.acs"
Private Sub a_eq_to_Click()
a_eq:
     a_more = InputBox("Enter the date :" & vbCrLf & "Format(mm,dd,yy)")
     If IsDate(a_more) = False Or CInt(Left(CStr(a_more), InStr(CStr(a_more), "/") - 1)) > 12 Then
      MsgBox "Enter a valid date"
      GoTo a_eq:
     End If
a_eq_to.Checked = True
a_none.Checked = False
a_l_than.Checked = False
a_m_than.Checked = False
Label9.Visible = True
Label9.Caption = "Last date of file access : " & a_more
End Sub
Private Sub a_l_than_Click()
a_l:
     a_less = InputBox("Enter the date :" & vbCrLf & "Format(mm,dd,yy)")
     If IsDate(a_less) = False Then
      MsgBox "Enter a valid date"
      GoTo a_l
     End If
     If CInt(Left(CStr(a_less), InStr(CStr(a_less), "/") - 1)) > 12 Then
     MsgBox "Enter a valid date"
      GoTo a_l
     End If
     If a_m_than.Checked = True And CDate(a_less) <= CDate(a_more) Then
     MsgBox "Check the condition"
     GoTo a_l
     End If
a_l_than.Checked = True
a_none.Checked = False
a_eq_to.Checked = False
Label9.Visible = True
     Label9.Caption = "Last date of file access range :" & " less than " & a_less
     If a_m_than.Checked = True Then
     Label9.Caption = "Last date of File access range :" & " greater than " & a_more & " but less than " & a_less
     End If
End Sub
Private Sub a_m_than_Click()
a_m:
     a_more = InputBox("Enter the date :" & vbCrLf & "Format(mm,dd,yy)")
     If IsDate(a_more) = False Then
      MsgBox "Enter a valid date"
      GoTo a_m
     End If
     If CInt(Left(CStr(a_more), InStr(CStr(a_more), "/") - 1)) > 12 Then
     MsgBox "Enter a valid date"
      GoTo a_m
     End If
     If a_l_than.Checked = True And CDate(a_less) <= CDate(a_more) Then
     MsgBox "Check the condition"
     GoTo a_m
     End If
a_m_than.Checked = True
a_none.Checked = False
a_eq_to.Checked = False
Label9.Visible = True
     Label9.Caption = "Last date of file access range :" & " greater than " & a_more
     If a_l_than.Checked = True Then
     Label9.Caption = Label9.Caption & " but less than " & a_less
     End If
End Sub
Private Sub a_none_Click()
a_none.Checked = True
a_m_than.Checked = False
a_l_than.Checked = False
a_eq_to.Checked = False
Label9.Visible = False
End Sub
Private Sub act_mer_Click()
act_mer.Checked = Not act_mer.Checked
End Sub
Private Sub asc_Click()
If ListView1.ListItems.Count > 0 Then
Asc.Checked = Not Asc.Checked
On Error GoTo ed
If Asc.Checked = True Then
For j = 1 To ListView1.ColumnHeaders.Count
If ListView1.ColumnHeaders.Item(j).Text = sort_text Then
sort_var = j - 1
End If
Next
If sort_text = "Size" Then
For j = 1 To ListView1.ListItems.Count
lit = ListView1.ListItems.Item(j).SubItems(sort_var)
For k = 1 To 10 - Len(lit)
lit = "0" + lit
Next
ListView1.ListItems.Item(j).SubItems(sort_var) = lit
Next
End If
If sort_text = "Date created" Or sort_text = "Date last accessed" Or sort_text = "Date last modified" Then
For j = 1 To ListView1.ListItems.Count
lit = ListView1.ListItems.Item(j).SubItems(sort_var)
If InStr(1, lit, "/") = 2 Then
lit = "0" + lit
End If
If Right(Left(lit, 5), 1) = "/" Then
lit = Left(lit, 3) + "0" + Right(lit, Len(lit) - 3)
End If
lit = Right(Left(lit, 10), 4) + Left(lit, 6) + Right(lit, Len(lit) - 10)
If Right(lit, 2) = "PM" Then
lit1 = Left(Right(lit, 11), 2)
hr = lit1
hr = hr + 12
lit1 = hr
lit = Left(lit, 11) + lit1 + Right(lit, 9)
End If
If Right(Left(lit, 13), 1) = ":" Then
lit = Left(lit, 11) + "0" + Right(lit, 10)
End If
ListView1.ListItems.Item(j).SubItems(sort_var) = lit
Next
End If
ListView1.Sorted = True
ListView1.SortKey = sort_var
ListView1.SortOrder = lvwAscending
If sort_text = "Size" Then
For j = 1 To ListView1.ListItems.Count
lit = ListView1.ListItems.Item(j).SubItems(sort_var)
While Left(lit, 1) = "0"
lit = Right(lit, Len(lit) - 1)
Wend
If Left(lit, 1) = " " Then
lit = "0" + lit
End If
ListView1.ListItems.Item(j).SubItems(sort_var) = lit
Next
End If
If sort_text = "Date created" Or sort_text = "Date last accessed" Or sort_text = "Date last modified" Then
For j = 1 To ListView1.ListItems.Count
lit = ListView1.ListItems.Item(j).SubItems(sort_var)
lit1 = Right(Left(lit, 10), 6)
lit = lit1 + Left(lit, 4) + Right(lit, Len(lit) - 10)
ListView1.ListItems.Item(j).SubItems(sort_var) = lit
Next
End If
End If
ed:
dsc.Checked = False
End If
End Sub
Private Sub atr_arh_Click()
If atr_arh.Checked = False Then
Label13.Caption = Label13.Caption & "A"
atr_none.Checked = False
Label12.Visible = True
Else
Label13.Caption = Replace(Label13.Caption, "A", "")
End If
atr_arh.Checked = Not atr_arh.Checked
If Label13.Caption = "" Then
atr_none_Click
End If
End Sub
Private Sub atr_hid_Click()
If atr_hid.Checked = False Then
Label13.Caption = Label13.Caption & "H"
atr_none.Checked = False
Label12.Visible = True
Else
Label13.Caption = Replace(Label13.Caption, "H", "")
End If
atr_hid.Checked = Not atr_hid.Checked
If Label13.Caption = "" Then
atr_none_Click
End If
End Sub
Private Sub atr_none_Click()
atr_none.Checked = True
atr_arh.Checked = False
atr_sys.Checked = False
atr_hid.Checked = False
atr_read.Checked = False
Label12.Visible = False
Label13.Caption = ""
End Sub
Private Sub atr_read_Click()
If atr_read.Checked = False Then
Label13.Caption = Label13.Caption & "R"
atr_none.Checked = False
Label12.Visible = True
Label13.Visible = True
Else
Label13.Caption = Replace(Label13.Caption, "R", "")
End If
atr_read.Checked = Not atr_read.Checked
If Label13.Caption = "" Then
atr_none_Click
End If
End Sub
Private Sub atr_sys_Click()
If atr_sys.Checked = False Then
Label13.Caption = Label13.Caption & "S"
atr_none.Checked = False
Label12.Visible = True
Else
Label13.Caption = Replace(Label13.Caption, "S", "")
End If
atr_sys.Checked = Not atr_sys.Checked
If Label13.Caption = "" Then
atr_none_Click
End If
End Sub
Private Sub attr_Click()
attr.Checked = Not attr.Checked
s_attr.Visible = attr.Checked
End Sub
Private Sub c_eq_to_Click()
c_eq:
     c_more = InputBox("Enter the date :" & vbCrLf & "Format(mm,dd,yy)")
     If IsDate(c_more) = False Or CInt(Left(CStr(c_more), InStr(CStr(c_more), "/") - 1)) > 12 Then
      MsgBox "Enter a valid date"
      GoTo c_eq:
     End If
c_eq_to.Checked = True
c_none.Checked = False
c_m_than.Checked = False
c_l_than.Checked = False
Label8.Visible = True
Label8.Caption = "File date creation : " & c_more
End Sub
Private Sub c_l_than_Click()
c_l:
     c_less = InputBox("Enter the date :" & vbCrLf & "Format(mm,dd,yy)")
     If IsDate(c_less) = False Then
      MsgBox "Enter a valid date"
      GoTo c_l
     End If
     If CInt(Left(CStr(c_less), InStr(CStr(c_less), "/") - 1)) > 12 Then
     MsgBox "Enter a valid date"
      GoTo c_l
     End If
     If c_m_than.Checked = True And CDate(c_less) <= CDate(c_more) Then
     MsgBox "Check the condition"
     GoTo c_l
     End If
c_l_than.Checked = True
c_none.Checked = False
c_eq_to.Checked = False
Label8.Visible = True
     Label8.Caption = "File creation date range :" & " less than " & c_less
     If c_m_than.Checked = True Then
     Label8.Caption = "File creation date range :" & " greater than " & c_more & " but less than " & c_less
     End If
End Sub
Private Sub c_m_than_Click()
c_m:
     c_more = InputBox("Enter the date :" & vbCrLf & "Format(mm,dd,yy)")
     If IsDate(c_more) = False Then
      MsgBox "Enter a valid date"
      GoTo c_m
     End If
     If CInt(Left(CStr(c_more), InStr(CStr(c_more), "/") - 1)) > 12 Then
     MsgBox "Enter a valid date"
      GoTo c_m
     End If
     If c_l_than.Checked = True And CDate(c_less) <= CDate(c_more) Then
     MsgBox "Check the condition"
     GoTo c_m
     End If
c_m_than.Checked = True
c_none.Checked = False
c_eq_to.Checked = False
Label8.Visible = True
     Label8.Caption = "File date creation range :" & " greater than " & c_more
     If c_l_than.Checked = True Then
     Label8.Caption = Label8.Caption & " but less than " & c_less
     End If
End Sub
Private Sub c_none_Click()
c_none.Checked = True
c_l_than.Checked = False
c_m_than.Checked = False
c_eq_to.Checked = False
Label8.Visible = False
End Sub
Private Sub date_acc_Click()
date_acc.Checked = Not date_acc.Checked
s_date_acc.Visible = date_acc.Checked
End Sub
Private Sub date_crd_Click()
date_crd.Checked = Not date_crd.Checked
s_date_crd.Visible = date_crd.Checked
End Sub
Private Sub date_mod_Click()
date_mod.Checked = Not date_mod.Checked
s_date_mod.Visible = date_mod.Checked
End Sub
Private Sub Dir1_Change()
Label15 = "Searching Path : " & Dir1.Path
End Sub
Private Sub doc_fil_Click()
doc_fil.Checked = Not doc_fil.Checked
If doc_fil.Checked = True Then
pic_fil.Checked = False
End If
End Sub
Private Sub dsc_Click()
If ListView1.ListItems.Count > 0 Then
On Error GoTo ed
dsc.Checked = Not dsc.Checked
If dsc.Checked = True Then
For j = 1 To ListView1.ColumnHeaders.Count
If ListView1.ColumnHeaders.Item(j).Text = sort_text Then
sort_var = j - 1
End If
Next
If sort_text = "Size" Then
For j = 1 To ListView1.ListItems.Count
lit = ListView1.ListItems.Item(j).SubItems(sort_var)
For k = 1 To 10 - Len(lit)
lit = "0" + lit
Next
ListView1.ListItems.Item(j).SubItems(sort_var) = lit
Next
End If
If sort_text = "Date created" Or sort_text = "Date last accessed" Or sort_text = "Date last modified" Then
For j = 1 To ListView1.ListItems.Count
lit = ListView1.ListItems.Item(j).SubItems(sort_var)
If InStr(1, lit, "/") = 2 Then
lit = "0" + lit
End If
If Right(Left(lit, 5), 1) = "/" Then
lit = Left(lit, 3) + "0" + Right(lit, Len(lit) - 3)
End If
lit = Right(Left(lit, 10), 4) + Left(lit, 6) + Right(lit, Len(lit) - 10)
If Right(lit, 2) = "PM" Then
lit1 = Left(Right(lit, 11), 2)
hr = lit1
hr = hr + 12
lit1 = hr
lit = Left(lit, 11) + lit1 + Right(lit, 9)
End If
If Right(Left(lit, 13), 1) = ":" Then
lit = Left(lit, 11) + "0" + Right(lit, 10)
End If
ListView1.ListItems.Item(j).SubItems(sort_var) = lit
Next
End If
ListView1.Sorted = True
ListView1.SortKey = sort_var
ListView1.SortOrder = lvwDescending
If sort_text = "Size" Then
For j = 1 To ListView1.ListItems.Count
lit = ListView1.ListItems.Item(j).SubItems(sort_var)
While Left(lit, 1) = "0"
lit = Right(lit, Len(lit) - 1)
Wend
If Left(lit, 1) = " " Then
lit = "0" + lit
End If
ListView1.ListItems.Item(j).SubItems(sort_var) = lit
Next
End If
If sort_text = "Date created" Or sort_text = "Date last accessed" Or sort_text = "Date last modified" Then
For j = 1 To ListView1.ListItems.Count
lit = ListView1.ListItems.Item(j).SubItems(sort_var)
lit1 = Right(Left(lit, 10), 6)
lit = lit1 + Left(lit, 4) + Right(lit, Len(lit) - 10)
ListView1.ListItems.Item(j).SubItems(sort_var) = lit
Next
End If
End If
ed:
Asc.Checked = False
End If
End Sub
Private Sub Form_Load()
Agent1.Characters.Load "merlin", DATAPATH
    Set merlin = Agent1.Characters("merlin")
    merlin.LanguageID = &H409
    merlin.Top = 400
    merlin.Left = 650
Drive1 = "c:"
srh_all_drv.Checked = False
srh_hid_fil.Checked = False
Timer1.Enabled = False
fl1 = False
fl2 = False
date_crd.Checked = False
date_acc.Checked = False
date_mod.Checked = False
typ.Checked = False
Asc.Checked = False
dsc.Checked = False
pic_fil.Checked = False
doc_fil.Checked = False
s_none_Click
c_none_Click
a_none_Click
m_none_Click
Label15 = "Searching Path : " & Dir1.Path
typ1.Checked = False
s_date_crd.Visible = False
s_date_acc.Visible = False
s_date_mod.Visible = False
s_typ.Visible = False
s_name.Checked = False
s_date_crd.Checked = False
s_size.Checked = False
s_date_acc.Checked = False
s_date_mod.Checked = False
s_par_fld.Checked = False
s_typ.Checked = False
s_attr.Checked = False
atr_none_Click
sp_medium_Click
End Sub
Private Sub cmdsrh_Click()
   merlin.Stop
   sb = True
   st = False
   counter2 = 1
   If s_none.Checked = False Or a_none.Checked = False _
   Or c_none.Checked = False Or pic_fil.Checked = True Then
   fl1 = True
   End If
   If m_none.Checked = False Or atr_none.Checked = False _
   Or typ1.Checked = True Or doc_fil.Checked = True Then
   fl2 = True
   End If
   MousePointer = vbHourglass
   Label15.Visible = False
   Timer1.Enabled = True
   cmdsrh.Enabled = False
   srh.Enabled = False
   txtsrh.Enabled = False
   Dir1.Visible = False
   Drive1.Visible = False
   adv_op.Enabled = False
   ch_det.Enabled = False
   sort.Enabled = False
   mnuadop.Enabled = False
   srh_by.Enabled = False
   srh_speed.Enabled = False
   Label3.Caption = ""
   Label4.Caption = "No file found"
   j = ListView1.ColumnHeaders.Count
   Do While j > 0
   ListView1.ColumnHeaders.Remove (j)
   j = j - 1
   Loop
   j = j + 1
   sort_var = 0
   ListView1.ColumnHeaders.Add j, , "Name", ListView1.Width / 4, lvwColumnLeft
   If Size.Checked = True Then
   If s_size.Checked = True Then
   sort_var = j
   End If
   j = j + 1
   ListView1.ColumnHeaders.Add j, , "Size", ListView1.Width / 4, lvwColumnRight
   End If
   If date_crd.Checked = True Then
   If s_date_crd.Checked = True Then
   sort_var = j
   End If
   j = j + 1
   ListView1.ColumnHeaders.Add j, , "Date created", ListView1.Width / 4, lvwColumnLeft
   End If
   If date_acc.Checked = True Then
   If s_date_acc.Checked = True Then
   sort_var = j
   End If
   j = j + 1
   ListView1.ColumnHeaders.Add j, , "Date last accessed", ListView1.Width / 4, lvwColumnLeft
   End If
   If date_mod.Checked = True Then
   If s_date_mod.Checked = True Then
   sort_var = j
   End If
   j = j + 1
   ListView1.ColumnHeaders.Add j, , "Date last modified", ListView1.Width / 4, lvwColumnLeft
   End If
   If par_fld.Checked = True Then
   If s_par_fld.Checked = True Then
   sort_var = j
   End If
   j = j + 1
   ListView1.ColumnHeaders.Add j, , "In folder", ListView1.Width / 4, lvwColumnLeft
   End If
   If typ.Checked = True Then
   If s_typ.Checked = True Then
   sort_var = j
   End If
   j = j + 1
   ListView1.ColumnHeaders.Add j, , "Type", ListView1.Width / 4, lvwColumnLeft
   End If
   If attr.Checked = True Then
   If s_attr.Checked = True Then
   sort_var = j
   End If
   j = j + 1
   ListView1.ColumnHeaders.Add j, , "Attributes", ListView1.Width / 4, lvwColumnLeft
   End If
   ListView1.View = lvwReport
   If srh_all_drv.Checked = True Then
   Label3.Caption = "Looking in all drives"
   ElseIf srh_sub_fld.Checked = False Then
   Label3.Caption = "Looking in " & Dir1.Path
   Else
   Label3.Caption = "Looking in " & Dir1.Path & " and subfolders"
   End If
   If srh_hid_fil.Checked = True Then
   Label5.Caption = "Looking in hidden files"
   Else
   Label5.Caption = "Not looking in hidden files"
   End If
   If srh_sys_fil.Checked = True Then
   Label6.Caption = "Looking in system files"
   Else
   Label6.Caption = "Not looking in system files"
   End If
   counter = 0
   counter1 = 0
   ListView1.ListItems.Clear
   searchflag = True
   txtsrh1 = txtsrh
   ' Searching process starts here
   If alldrv = True Then
    For i = 0 To Drive1.ListCount - 1
Y:
     sb = True
    On Error Resume Next
    Drive1 = Drive1.List(i)
    Dir1.Path = Drive1
     If Len(Dir1.Path) > 3 Then
      Dir1.Path = Left(Dir1.Path, 3)
     End If
     If Drive1 = Drive1.List(i) Then
     FindFile Dir1.Path, txtsrh
     End If
    Next
   Else
     FindFile Dir1.Path, txtsrh1
   End If
z:
   If alldrv = True Then
    Dir1.Path = Drive1
   End If
   Dir1.Visible = True
   Drive1.Visible = True
   Label2.Caption = ""
   Label6.Caption = ""
   cmdsrh.Enabled = True
   txtsrh.Enabled = True
   srh.Enabled = True
   adv_op.Enabled = True
   sort.Enabled = True
   ch_det.Enabled = True
   mnuadop.Enabled = True
   srh_speed.Enabled = True
   Timer1.Enabled = False
   srh_by.Enabled = True
   Label15.Visible = True
   MousePointer = vbDefault
   If st = False Then
   If act_mer.Checked = True Then
   merlin.Show
   merlin.Speak "Search finished"
   merlin.Hide
   End If
   End If
   End Sub
Private Function FindFile(ByVal sFol As String, txtsrh1 As String) As Long
   Dim tFld As Folder, tfil As File, li As ListItem
   Dim fp, fp1 As String
   Dim sz As Variant
   Set fld = fso.GetFolder(sFol)
     If sb = True Then
     File1.Path = Dir1.Path
     If InStrRev(txtsrh1, "*") = 0 And InStrRev(txtsrh1, "?") = 0 Then
      File1.Pattern = "*" & txtsrh1 & "*"
      Else
      File1.Pattern = txtsrh1
     End If
     fp1 = File1.Path
     For z = 0 To File1.ListCount - 1
      fp = fp1 & "\" & File1.List(z)
      fp = Replace(fp, "\\", "\")
      On Error Resume Next
      Set tfil = fso.GetFile(fp)
      If srh_arh.Checked = True Or tfil.Attributes < 32 Then
       If s_none.Checked = True Then GoTo s2
       sz = tfil.Size
       If s_m_than.Checked = False Or Round(sz / 1024, 0) > CInt(s_more) Then
        If s_l_than.Checked = False Or Round(sz / 1024, 0) < CInt(s_less) Then
         If s_eq_to.Checked = False Or Round(sz / 1024, 0) = CInt(s_more) Then
s2:
          If c_none.Checked = True Then GoTo c2
          If c_m_than.Checked = False Or Left(tfil.DateCreated, InStr(tfil.DateCreated, " ")) > CDate(c_more) Then
           If c_l_than.Checked = False Or Left(tfil.DateCreated, InStr(tfil.DateCreated, " ")) < CDate(c_less) Then
            If c_eq_to.Checked = False Or Left(tfil.DateCreated, InStr(tfil.DateCreated, " ")) = CDate(c_more) Then
c2:
             If a_none.Checked = True Then GoTo a2
             If a_m_than.Checked = False Or tfil.DateLastAccessed > CDate(a_more) Then
              If a_l_than.Checked = False Or tfil.DateLastAccessed < CDate(a_less) Then
               If a_eq_to.Checked = False Or tfil.DateLastAccessed = CDate(a_more) Then
a2:
                If m_none.Checked = True Then GoTo m2
                If m_m_than.Checked = False Or Left(tfil.DateLastModified, InStr(tfil.DateLastModified, " ")) > CDate(m_more) Then
                 If m_l_than.Checked = False Or Left(tfil.DateLastModified, InStr(tfil.DateLastModified, " ")) < CDate(m_less) Then
                  If m_eq_to.Checked = False Or Left(tfil.DateLastModified, InStr(tfil.DateLastModified, " ")) = CDate(m_more) Then
m2:
                   If typ1.Checked = False Or LCase(tfil.Type) = t_typ Then
                    If pic_fil.Checked = False Or LCase(Right(tfil.Type, 5)) = "image" Then
                     If doc_fil.Checked = False Or LCase(Right(tfil.Type, 8)) = "document" Then
                      If atr_none.Checked = True Then GoTo at2
                      attri = tfil.Attributes
                      If atr_arh.Checked = False Or attri >= 32 Then
                       If attri >= 32 Then
                        attri = attri - 32
                       End If
                       If atr_sys.Checked = False Or attri >= 6 Then
                        If attri >= 6 Then
                         attri = attri - 6
                        End If
                        If atr_hid.Checked = False Or attri >= 2 Then
                         If attri >= 2 Then
                          attri = attri - 2
                         End If
                         If atr_read.Checked = False Or attri >= 1 Then
at2:
                          Set li = ListView1.ListItems.Add(, , tfil.Name)
                          k = 1
                          If Size.Checked = True Then
                           li.SubItems(k) = Round(tfil.Size / 1024, 0) & " KB"
                           k = k + 1
                          End If
                          If date_crd.Checked = True Then
                           li.SubItems(k) = tfil.DateCreated
                           k = k + 1
                          End If
                          If date_acc.Checked = True Then
                           li.SubItems(k) = tfil.DateLastAccessed
                           k = k + 1
                          End If
                          If date_mod.Checked = True Then
                           li.SubItems(k) = tfil.DateLastModified
                           k = k + 1
                          End If
                          If par_fld.Checked = True Then
                           li.SubItems(k) = fp1
                           k = k + 1
                          End If
                          If typ.Checked = True Then
                           li.SubItems(k) = tfil.Type
                           k = k + 1
                          End If
                          If attr.Checked = True Then
                           atr = tfil.Attributes
                           If atr >= 32 Then
                            atr = atr - 32
                            li.SubItems(k) = "A"
                           End If
                           If atr >= 6 Then
                            atr = atr - 6
                            li.SubItems(k) = li.SubItems(k) & "S"
                           End If
                           If atr >= 2 Then
                            atr = atr - 2
                            li.SubItems(k) = li.SubItems(k) & "H"
                           End If
                           If atr = 1 Then
                            li.SubItems(k) = li.SubItems(k) & "R"
                           End If
                          End If
                          counter = counter + 1
                          Label4.Caption = "No. of files found " & counter
                         End If
                        End If
                       End If
                      End If
                     End If
                    End If
                   End If
                  End If
                 End If
                End If
               End If
              End If
             End If
            End If
           End If
          End If
         End If
        End If
       End If
       DoEvents
      End If
     Next
     sb = False
    End If
    If srh_sub_fld.Checked = True Then
    If fld.SubFolders.Count > 0 Then
     For Each tFld In fld.SubFolders
      If searchflag = True Then
       File1.Path = tFld.Path
       If InStrRev(txtsrh1, "*") = 0 And InStrRev(txtsrh1, "?") = 0 Then
        File1.Pattern = "*" & txtsrh1 & "*"
        Else
        File1.Pattern = txtsrh1
       End If
       Label2.Caption = tFld.Path
       counter1 = File1.ListCount
       If File1.ListCount < 100 Then
        counter2 = 1
       End If
       fp1 = File1.Path
       If Right(fp1, 1) = "\" Then
       fp1 = Left(fp1, Len(fp1) - 1)
       End If
       For z = 0 To File1.ListCount - 1
        If searchflag = True Then
         fp = fp1 & "\" & File1.List(z)
         If Not File1.List(z) = "?" Then
         On Error Resume Next
         Set tfil = fso.GetFile(fp)
         If srh_arh.Checked = True Or tfil.Attributes < 32 Then
          If fl1 = False Then GoTo ac
          If pic_fil.Checked = False Or LCase(Right(tfil.Type, 5)) = "image" Then
           If s_none.Checked = True Then GoTo se
           s = Round(tfil.Size / 1024, 0)
           If s_m_than.Checked = False Or s > CInt(s_more) Then
            If s_l_than.Checked = False Or s < CInt(s_less) Then
             If s_eq_to.Checked = False Or s = CInt(s_more) Then
se:
              If c_none.Checked = True Then GoTo cr
              c = Left(tfil.DateCreated, InStr(tfil.DateCreated, " "))
              If c_m_than.Checked = False Or c > CDate(c_more) Then
               If c_l_than.Checked = False Or c < CDate(c_less) Then
                If c_eq_to.Checked = False Or c = CDate(c_more) Then
cr:
                 If a_none.Checked = True Then GoTo ac
                 a = tfil.DateLastAccessed
                 If a_m_than.Checked = False Or a > CDate(a_more) Then
                  If a_l_than.Checked = False Or a < CDate(a_less) Then
                   If a_eq_to.Checked = False Or a = CDate(a_more) Then
ac:
                    attri = tfil.Attributes
                    atr = attri
                    If fl2 = False Then GoTo at
                    If m_none.Checked = True Then GoTo md
                    m = Left(tfil.DateLastModified, InStr(tfil.DateLastModified, " "))
                    If m_m_than.Checked = False Or m > CDate(m_more) Then
                     If m_l_than.Checked = False Or m < CDate(m_less) Then
                      If m_eq_to.Checked = False Or m = CDate(m_more) Then
md:
                       If typ1.Checked = False Or LCase(tfil.Type) = t_typ Then
                        If doc_fil.Checked = False Or LCase(Right(tfil.Type, 8)) = "document" Then
                         
                         If atr_none.Checked = True Then GoTo at
                         
                         If atr_arh.Checked = False Or attri >= 32 Then
                          If attri >= 32 Then
                           attri = attri - 32
                          End If
                          If atr_sys.Checked = False Or attri >= 6 Then
                           If attri >= 6 Then
                            attri = attri - 6
                           End If
                           If atr_hid.Checked = False Or attri >= 2 Then
                            If attri >= 2 Then
                             attri = attri - 2
                            End If
                            If atr_read.Checked = False Or attri >= 1 Then
at:
                             Set li = ListView1.ListItems.Add(, , tfil.Name)
                             counter = counter + 1
                             counter1 = counter1 + 1
                             k = 1
                             If Size.Checked = True Then
                              li.SubItems(k) = Round(tfil.Size / 1024, 0) & " KB"
                              k = k + 1
                             End If
                             If date_crd.Checked = True Then
                              li.SubItems(k) = tfil.DateCreated
                              k = k + 1
                             End If
                             If date_acc.Checked = True Then
                              li.SubItems(k) = tfil.DateLastAccessed
                              k = k + 1
                             End If
                             If date_mod.Checked = True Then
                              li.SubItems(k) = tfil.DateLastModified
                              k = k + 1
                             End If
                             li.SubItems(k) = fp1
                             k = k + 1
                              If typ.Checked = True Then
                              li.SubItems(k) = tfil.Type
                              k = k + 1
                             End If
                             If attr.Checked = True Then
                              If atr >= 32 Then
                               atr = atr - 32
                               li.SubItems(k) = "A"
                              End If
                              If atr >= 6 Then
                               atr = atr - 6
                               li.SubItems(k) = li.SubItems(k) & "S"
                              End If
                              If atr >= 2 Then
                               atr = atr - 2
                               li.SubItems(k) = li.SubItems(k) & "H"
                              End If
                              If atr = 1 Then
                               li.SubItems(k) = li.SubItems(k) & "R"
                              End If
                             End If
                             Label4.Caption = "No. of files found " & counter
                            End If
                           End If
                          End If
                         End If
                        End If
                       End If
                      End If
                     End If
                    End If
                   End If
                  End If
                 End If
                End If
               End If
              End If
             End If
            End If
           End If
           End If
          End If
         End If
        End If
ne:
        If counter Mod n = 0 Then
         DoEvents
        End If
        If txtsrh1 <> "*.*" Or fl1 = True Or fl2 = True Then
         If counter Mod (n / 10) = 0 Then
          DoEvents
        End If
        End If
       Next
       If txtsrh1 <> "*.*" Or fl1 = True Or fl2 = True Then
        DoEvents
     End If
       FindFile = FindFile + FindFile(tFld.Path, txtsrh1)
      End If
     Next
    End If
    End If
   End Function
Private Sub cmdstop_Click()
searchflag = False
st = True
End Sub
Private Sub Drive1_Change()
 On Error GoTo drivehandler
 Dir1.Path = Drive1
  If Len(Dir1.Path) > 3 Then
   Dir1.Path = Left(Dir1.Path, 3)
  End If
Exit Sub
drivehandler:
  MsgBox ("Drive is not ready")
  Drive1 = Dir1.Path
End Sub
Private Sub exit_Click()
Unload Me
End Sub
Private Sub m_eq_to_Click()
m_eq:
     m_more = InputBox("Enter the date :" & vbCrLf & "Format(mm,dd,yy)")
     If IsDate(m_more) = False Or CInt(Left(CStr(m_more), InStr(CStr(m_more), "/") - 1)) > 12 Then
      MsgBox "Enter a valid date"
      GoTo m_eq:
     End If
m_eq_to.Checked = True
m_l_than.Checked = False
m_m_than.Checked = False
m_none.Checked = False
Label10.Visible = True
Label10.Caption = "Last of file modification range : " & m_more
End Sub
Private Sub m_l_than_Click()
m_l:
     m_less = InputBox("Enter the date :" & vbCrLf & "Format(mm,dd,yy)")
     If IsDate(m_less) = False Then
      MsgBox "Enter a valid date"
      GoTo m_l
     End If
     If CInt(Left(CStr(m_less), InStr(CStr(m_less), "/") - 1)) > 12 Then
     MsgBox "Enter a valid date"
      GoTo m_l
     End If
     If m_m_than.Checked = True And CDate(m_less) <= CDate(m_more) Then
     MsgBox "Check the condition"
     GoTo m_l
     End If
m_l_than.Checked = True
m_none.Checked = False
m_eq_to.Checked = False
Label10.Visible = True
     Label10.Caption = "Last date of file modification range :" & " less than " & m_less
     If m_m_than.Checked = True Then
     Label10.Caption = "Last date of File modification range :" & " greater than " & m_more & " but less than " & m_less
     End If
End Sub
Private Sub m_m_than_Click()
m_m:
     m_more = InputBox("Enter the date :" & vbCrLf & "Format(mm,dd,yy)")
     If IsDate(m_more) = False Then
      MsgBox "Enter a valid date"
      GoTo m_m
     End If
     If CInt(Left(CStr(m_more), InStr(CStr(m_more), "/") - 1)) > 12 Then
     MsgBox "Enter a valid date"
      GoTo m_m
     End If
     If m_l_than.Checked = True And CDate(m_less) <= CDate(m_more) Then
     MsgBox "Check the condition"
     GoTo m_m
     End If
m_m_than.Checked = True
m_none.Checked = False
m_eq_to.Checked = False
Label10.Visible = True
     Label10.Caption = "Last date of file modification range :" & " greater than " & m_more
     If m_l_than.Checked = True Then
     Label10.Caption = Label10.Caption & " but less than " & m_less
     End If
End Sub
Private Sub m_none_Click()
m_none.Checked = True
m_l_than.Checked = False
m_m_than.Checked = False
m_eq_to.Checked = False
Label10.Visible = False
End Sub
Private Sub pic_fil_Click()
pic_fil.Checked = Not pic_fil.Checked
If pic_fil.Checked = True Then
doc_fil.Checked = False
End If
End Sub

Private Sub s_attr_Click()
s_attr.Checked = Not s_attr.Checked
If s_attr.Checked = True Then
s_name.Checked = False
s_date_crd.Checked = False
s_date_acc.Checked = False
s_date_mod.Checked = False
s_par_fld.Checked = False
s_typ.Checked = False
s_size.Checked = False
sort_text = "Attributes"
End If
If Asc.Checked = True Then
asc_Click
asc_Click
End If
If dsc.Checked = True Then
dsc_Click
dsc_Click
End If
End Sub

Private Sub s_date_acc_Click()
s_date_acc.Checked = Not s_date_acc.Checked
If s_date_acc.Checked = True Then
s_name.Checked = False
s_date_crd.Checked = False
s_size.Checked = False
s_date_mod.Checked = False
s_par_fld.Checked = False
s_typ.Checked = False
s_attr.Checked = False
sort_text = "Date last accessed"
End If
If Asc.Checked = True Then
asc_Click
asc_Click
End If
If dsc.Checked = True Then
dsc_Click
dsc_Click
End If
End Sub
Private Sub s_date_crd_Click()
s_date_crd.Checked = Not s_date_crd.Checked
If s_date_crd.Checked = True Then
s_name.Checked = False
s_size.Checked = False
s_date_acc.Checked = False
s_date_mod.Checked = False
s_par_fld.Checked = False
s_typ.Checked = False
s_attr.Checked = False
sort_text = "Date created"
End If
If Asc.Checked = True Then
asc_Click
asc_Click
End If
If dsc.Checked = True Then
dsc_Click
dsc_Click
End If
End Sub
Private Sub s_date_mod_Click()
s_date_mod.Checked = Not s_date_mod.Checked
If s_date_mod.Checked = True Then
s_name.Checked = False
s_date_crd.Checked = False
s_date_acc.Checked = False
s_size.Checked = False
s_par_fld.Checked = False
s_typ.Checked = False
s_attr.Checked = False
sort_text = "Date last modified"
End If
If Asc.Checked = True Then
asc_Click
asc_Click
End If
If dsc.Checked = True Then
dsc_Click
dsc_Click
End If
End Sub

Private Sub s_eq_to_Click()
s_eq:
     s_more = InputBox("Enter the size of file :")
     If IsNumeric(s_more) = False Then
      MsgBox "Enter a valid digit"
      GoTo s_eq:
     End If
s_eq_to.Checked = True
s_l_than.Checked = False
s_m_than.Checked = False
s_none.Checked = False
Label7.Visible = True
 Label7.Caption = "File size  : " & s_more & " KB"
     End Sub
Private Sub s_l_than_Click()
s_l:
    s_less = InputBox("Enter the size of file :")
    If IsNumeric(s_less) = False Then
     MsgBox "Enter a valid digit"
     GoTo s_l
    End If
    If s_m_than.Checked = True And CInt(s_less) <= CInt(s_more) Then
     MsgBox "Check the condition"
     GoTo s_l
     End If
    s_l_than.Checked = True
    s_eq_to.Checked = False
    s_none.Checked = False
    Label7.Visible = True
     Label7.Caption = "File size range :" & " less than " & s_less & " KB"
     If s_m_than.Checked = True Then
     Label7.Caption = "File size range :" & " greater than " & s_more & " KB  but less than " & s_less & "KB"
     End If
End Sub
Private Sub s_m_than_Click()
s_m:
     s_more = InputBox("Enter the size of file :")
     If IsNumeric(s_more) = False Then
      MsgBox "Enter a valid digit"
      GoTo s_m
     End If
     If s_l_than.Checked = True And CInt(s_less) <= CInt(s_more) Then
     MsgBox "Check the condition"
     GoTo s_m
     End If
     Label7.Caption = "File size range :" & " greater than " & s_more & " KB"
     If s_l_than.Checked = True Then
       Label7.Caption = Label7.Caption & " but less than " & s_less & " KB "
     End If
     s_m_than.Checked = True
     s_none.Checked = False
     s_eq_to.Checked = False
     Label7.Visible = True
End Sub

Private Sub s_name_Click()
s_name.Checked = Not s_name.Checked
If s_name.Checked = True Then
s_size.Checked = False
s_date_crd.Checked = False
s_date_acc.Checked = False
s_date_mod.Checked = False
s_size.Checked = False
s_typ.Checked = False
s_attr.Checked = False
sort_text = "Name"
End If
If Asc.Checked = True Then
asc_Click
asc_Click
End If
If dsc.Checked = True Then
dsc_Click
dsc_Click
End If
End Sub
Private Sub s_none_Click()
s_none.Checked = True
s_m_than.Checked = False
s_l_than.Checked = False
s_eq_to.Checked = False
Label7.Visible = False
End Sub

Private Sub s_par_fld_Click()
s_par_fld.Checked = Not s_par_fld.Checked
If s_par_fld.Checked = True Then
s_name.Checked = False
s_date_crd.Checked = False
s_date_acc.Checked = False
s_date_mod.Checked = False
s_size.Checked = False
s_typ.Checked = False
s_attr.Checked = False
sort_text = "In folder"
End If
If Asc.Checked = True Then
asc_Click
asc_Click
End If
If dsc.Checked = True Then
dsc_Click
dsc_Click
End If
End Sub

Private Sub s_size_Click()
s_size.Checked = Not s_size.Checked
If s_size.Checked = True Then
s_name.Checked = False
s_date_crd.Checked = False
s_date_acc.Checked = False
s_date_mod.Checked = False
s_par_fld.Checked = False
s_typ.Checked = False
s_attr.Checked = False
sort_text = "Size"
End If
If Asc.Checked = True Then
asc_Click
asc_Click
End If
If dsc.Checked = True Then
dsc_Click
dsc_Click
End If
End Sub

Private Sub s_typ_Click()
s_typ.Checked = Not s_typ.Checked
If s_typ.Checked = True Then
s_name.Checked = False
s_date_crd.Checked = False
s_date_acc.Checked = False
s_date_mod.Checked = False
s_par_fld.Checked = False
s_size.Checked = False
s_attr.Checked = False
sort_text = "Type"
End If
If Asc.Checked = True Then
asc_Click
asc_Click
End If
If dsc.Checked = True Then
dsc_Click
dsc_Click
End If
End Sub
Private Sub size_Click()
Size.Checked = Not Size.Checked
s_size.Visible = Size.Checked
End Sub
Private Sub sp_fast_Click()
sp_fast.Checked = True
sp_medium.Checked = False
n = 1000
End Sub
Private Sub sp_medium_Click()
sp_medium.Checked = True
sp_fast.Checked = False
n = 100
End Sub
Private Sub srh_arh_Click()
srh_arh.Checked = Not srh_arh.Checked
End Sub
Private Sub srh_Click()
cmdsrh_Click
End Sub
Private Sub srh_hid_fil_Click()
srh_hid_fil.Checked = Not srh_hid_fil.Checked
File1.Hidden = Not File1.Hidden
End Sub
Private Sub srh_sys_fil_Click()
srh_sys_fil.Checked = Not srh_sys_fil.Checked
File1.System = Not File1.System
End Sub
Private Sub srh_all_drv_Click()
srh_all_drv.Checked = Not srh_all_drv.Checked
If srh_all_drv.Checked = True Then
srh_sub_fld.Checked = True
Dir1.Enabled = False
Drive1.Enabled = False
sp_fast_Click
Label14.Caption = "Search in all drives is active"
Else
Dir1.Enabled = True
Drive1.Enabled = True
Label14.Caption = ""
End If
alldrv = Not alldrv
End Sub
Private Sub srh_sub_fld_Click()
If srh_all_drv.Checked = True Then
srh_sub_fld.Checked = True
Else
srh_sub_fld.Checked = Not srh_sub_fld.Checked
End If
sp_medium_Click
End Sub
Private Sub stop_Click()
cmdstop_Click
End Sub
Private Sub Timer1_Timer()
fl = False
counter3 = counter3 + 1
If counter3 Mod 10 = 0 Then
fl = True
Timer1.Enabled = False
End If
End Sub
Private Sub typ_Click()
typ.Checked = Not typ.Checked
s_typ.Visible = typ.Checked
End Sub
Private Sub typ1_Click()
typ1.Checked = Not typ1.Checked
If typ1.Checked = True Then
t_typ = InputBox("Enter the type of file")
Label11.Caption = "Type : " & t_typ
t_typ = LCase(t_typ)
End If
Label11.Visible = Not Label11.Visible
End Sub



