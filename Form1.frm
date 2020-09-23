VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   7005
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10320
   FillColor       =   &H00C0C0C0&
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C0C0C0&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   10320
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "1"
      ForeColor       =   &H80000008&
      Height          =   6015
      Index           =   1
      Left            =   2400
      MousePointer    =   9  'Size W E
      TabIndex        =   23
      Top             =   720
      Width           =   50
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      ForeColor       =   &H00000000&
      Height          =   35
      Left            =   2760
      TabIndex        =   22
      Top             =   1200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00000000&
      Height          =   35
      Left            =   2760
      TabIndex        =   21
      Top             =   890
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ListBox List6 
      Height          =   1110
      ItemData        =   "Form1.frx":08CA
      Left            =   2760
      List            =   "Form1.frx":08CC
      TabIndex        =   16
      Top             =   2040
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.ListBox List5 
      Height          =   1320
      ItemData        =   "Form1.frx":08CE
      Left            =   2760
      List            =   "Form1.frx":08D0
      Sorted          =   -1  'True
      TabIndex        =   15
      Top             =   3240
      Visible         =   0   'False
      Width           =   2535
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1095
      Left            =   2760
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1931
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"Form1.frx":08D2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox CmdView 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   2520
      Picture         =   "Form1.frx":09B4
      ScaleHeight     =   210
      ScaleWidth      =   240
      TabIndex        =   20
      TabStop         =   0   'False
      ToolTipText     =   "Auto Syntax OFF"
      Top             =   1680
      Width           =   240
   End
   Begin VB.PictureBox cmdFind 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   2450
      Picture         =   "Form1.frx":0C96
      ScaleHeight     =   195
      ScaleWidth      =   270
      TabIndex        =   19
      TabStop         =   0   'False
      ToolTipText     =   "Find"
      Top             =   1080
      Width           =   270
   End
   Begin VB.PictureBox cmdReplace 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   2450
      Picture         =   "Form1.frx":0FB0
      ScaleHeight     =   240
      ScaleWidth      =   270
      TabIndex        =   18
      TabStop         =   0   'False
      ToolTipText     =   "Replace"
      Top             =   1320
      Width           =   270
   End
   Begin VB.PictureBox CmdClose 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   2520
      Picture         =   "Form1.frx":1372
      ScaleHeight     =   135
      ScaleWidth      =   180
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "Close Document"
      Top             =   840
      Width           =   180
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   9720
      Top             =   1440
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   6750
      Width           =   10320
      _ExtentX        =   18203
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   10001
            Picture         =   "Form1.frx":14F8
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox RichTextBox2 
      Height          =   375
      Left            =   9000
      TabIndex        =   13
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form1.frx":1A3A
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   36000
      Left            =   9720
      Top             =   960
   End
   Begin VB.ListBox List4 
      Height          =   270
      Left            =   9600
      TabIndex        =   12
      Top             =   3000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3210
      Left            =   120
      System          =   -1  'True
      TabIndex        =   6
      Top             =   3120
      Width           =   2175
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2205
      ItemData        =   "Form1.frx":1B23
      Left            =   120
      List            =   "Form1.frx":1B2A
      TabIndex        =   9
      Top             =   840
      Width           =   2175
   End
   Begin VB.ListBox List3 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   1815
      Left            =   120
      TabIndex        =   11
      Top             =   4680
      Width           =   2175
   End
   Begin VB.FileListBox File2 
      Height          =   300
      Left            =   9360
      Pattern         =   "*.ftp"
      TabIndex        =   10
      Top             =   2640
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSComctlLib.ImageList ToolBarIcons 
      Left            =   9120
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   19
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1B3E
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1C50
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1D62
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1E74
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1F86
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2098
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":21AA
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":22BC
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":23CE
            Key             =   "Redo"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":24EE
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2600
            Key             =   "Align Left"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2712
            Key             =   "Align Right"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2824
            Key             =   "Center"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2936
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2A48
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2B5A
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2C6C
            Key             =   "Small Caps"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2D7E
            Key             =   "Strike Through"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2E90
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   1425
      Left            =   120
      TabIndex        =   8
      Top             =   3120
      Width           =   2175
   End
   Begin VB.DriveListBox Drive1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   2175
   End
   Begin MSComDlg.CommonDialog SaveDlg 
      Left            =   8640
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog OpenDlg 
      Left            =   9120
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.DirListBox Dir1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1890
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   2175
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   1095
      Left            =   2760
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   1455
      ExtentX         =   2566
      ExtentY         =   1931
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   6255
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   11033
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Local"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Remote"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.TabStrip TabStrip2 
      Height          =   6240
      Left            =   2400
      TabIndex        =   2
      Top             =   480
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   11007
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Edit"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Browse"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   10320
      _ExtentX        =   18203
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ToolBarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   18
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Undo"
            Object.ToolTipText     =   "Undo"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Redo"
            Object.ToolTipText     =   "Redo"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Find"
            Object.ToolTipText     =   "Find"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Align Left"
            Object.ToolTipText     =   "Align Left"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Align Right"
            Object.ToolTipText     =   "Align Right"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Center"
            Object.ToolTipText     =   "Center"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bold"
            Object.ToolTipText     =   "Bold"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Italic"
            Object.ToolTipText     =   "Italic"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Underline"
            Object.ToolTipText     =   "Underline"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Small Caps"
            Object.ToolTipText     =   "Small Caps"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Strike Through"
            Object.ToolTipText     =   "Strike Through"
            ImageIndex      =   18
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "hello"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
   End
   Begin VB.Menu MnuFile 
      Caption         =   "&File"
      Begin VB.Menu MnuNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu s1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu MnuOpenFromWeb 
         Caption         =   "&Open From Web"
      End
      Begin VB.Menu s2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu MnuSaveAs 
         Caption         =   "Save &As"
      End
      Begin VB.Menu MnuSaveToServer 
         Caption         =   "Save &To Server"
      End
      Begin VB.Menu MnuSaveAsTemplate 
         Caption         =   "Save As &Template"
      End
      Begin VB.Menu s3 
         Caption         =   "-"
      End
      Begin VB.Menu MnuClose 
         Caption         =   "&Close File"
         Shortcut        =   ^W
      End
      Begin VB.Menu s4 
         Caption         =   "-"
      End
      Begin VB.Menu MnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu MnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu MnuUndo 
         Caption         =   "&Undo"
      End
      Begin VB.Menu MnuRedo 
         Caption         =   "&Redo"
      End
      Begin VB.Menu s5 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCut 
         Caption         =   "&Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu MnuCopy 
         Caption         =   "C&opy"
         Shortcut        =   ^C
      End
      Begin VB.Menu MnuPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^Z
      End
      Begin VB.Menu MnuDelete 
         Caption         =   "&Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu MnuSelectAll 
         Caption         =   "&Select All"
         Shortcut        =   ^A
      End
      Begin VB.Menu s6 
         Caption         =   "-"
      End
      Begin VB.Menu MnuIndent 
         Caption         =   "&Indent"
         Shortcut        =   ^I
      End
      Begin VB.Menu MnuUnIndent 
         Caption         =   "&UnIndent"
         Shortcut        =   ^U
      End
      Begin VB.Menu s7 
         Caption         =   "-"
      End
      Begin VB.Menu MnuWordWrap 
         Caption         =   "&WordWrap"
         Checked         =   -1  'True
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu MnuSearch 
      Caption         =   "&Search"
      Begin VB.Menu MnuFind 
         Caption         =   "&Find"
         Shortcut        =   ^F
      End
      Begin VB.Menu MnuReplace 
         Caption         =   "&Replace"
         Shortcut        =   ^R
      End
   End
   Begin VB.Menu MnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu MnuTagChooser 
         Caption         =   "&Tag Chooser"
         Shortcut        =   {F8}
      End
      Begin VB.Menu MnuSqlBuiler 
         Caption         =   "&Sql Builder"
         Shortcut        =   {F7}
      End
      Begin VB.Menu MnuDocumentWeight 
         Caption         =   "&Document Weight"
      End
   End
   Begin VB.Menu MnuView 
      Caption         =   "&View"
      Begin VB.Menu MnuResourceTab 
         Caption         =   "&Resource Tab"
         Checked         =   -1  'True
         Shortcut        =   {F9}
      End
      Begin VB.Menu MnuToggleEB 
         Caption         =   "&Toggle Edit/Browse"
         Shortcut        =   {F12}
      End
      Begin VB.Menu MnuViewWithExternalBrowsers 
         Caption         =   "&View With External Browsers"
         Shortcut        =   {F11}
      End
      Begin VB.Menu MnuSettings 
         Caption         =   "&Settings"
      End
   End
   Begin VB.Menu MnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu MnuMerkatumOnline 
         Caption         =   "&Merkatum Home Page"
      End
      Begin VB.Menu MnuOnlineSupport 
         Caption         =   "&Online Support"
      End
      Begin VB.Menu MnuAbout 
         Caption         =   "&About"
      End
   End
   Begin VB.Menu MnuA 
      Caption         =   "A"
      Visible         =   0   'False
      Begin VB.Menu MnuUndoPop 
         Caption         =   "&Undo"
      End
      Begin VB.Menu MnuARedo 
         Caption         =   "&Redo"
      End
      Begin VB.Menu s8 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCutPop 
         Caption         =   "&Cut"
      End
      Begin VB.Menu MnuCopyPop 
         Caption         =   "C&opy"
      End
      Begin VB.Menu MnuPastePop 
         Caption         =   "&Paste"
      End
      Begin VB.Menu MnuDeletePop 
         Caption         =   "&Delete"
      End
      Begin VB.Menu MnuFilePop 
         Caption         =   "&File"
         Begin VB.Menu MnuNewPop 
            Caption         =   "&New"
         End
         Begin VB.Menu s9 
            Caption         =   "-"
         End
         Begin VB.Menu MnuOpenPop 
            Caption         =   "&Open"
         End
         Begin VB.Menu MnuOpenFromWebPop 
            Caption         =   "Open &From Web"
         End
         Begin VB.Menu s10 
            Caption         =   "-"
         End
         Begin VB.Menu MnuSavePop 
            Caption         =   "&Save"
         End
         Begin VB.Menu MnuSaveAsPop 
            Caption         =   "&Save As"
         End
         Begin VB.Menu MnuDSaveToServer 
            Caption         =   "&Save To Server"
         End
         Begin VB.Menu MnuAutoSave 
            Caption         =   "&Last AutoSaved Copy"
         End
         Begin VB.Menu s11 
            Caption         =   "-"
         End
         Begin VB.Menu MnuClosePop 
            Caption         =   "&Close"
         End
      End
      Begin VB.Menu MnuInsertTagPop 
         Caption         =   "&Insert Taq"
      End
      Begin VB.Menu MnuPopViewWithExternal 
         Caption         =   "&View With External Browsers"
      End
      Begin VB.Menu MnuSelection 
         Caption         =   "&Selection"
         Begin VB.Menu MnuUpperCase 
            Caption         =   "&UpperCase"
         End
         Begin VB.Menu MnuLowerCase 
            Caption         =   "&LowerCase"
         End
         Begin VB.Menu MnuConvertToTable 
            Caption         =   "&Convert To Table"
         End
      End
   End
   Begin VB.Menu MnuB 
      Caption         =   "B"
      Visible         =   0   'False
      Begin VB.Menu MnuAddServer 
         Caption         =   "&Add Server"
      End
      Begin VB.Menu MnuConnect 
         Caption         =   "&Connect"
      End
      Begin VB.Menu MnuDisconnect 
         Caption         =   "D&isconnect"
      End
      Begin VB.Menu MnuDeleteServer 
         Caption         =   "&Delete Server"
      End
      Begin VB.Menu MnuServerProperties 
         Caption         =   "&Properties"
      End
      Begin VB.Menu MnuRefreshServes 
         Caption         =   "&Refresh"
      End
   End
   Begin VB.Menu MnuC 
      Caption         =   "C"
      Visible         =   0   'False
      Begin VB.Menu MnuEditPop 
         Caption         =   "&Edit"
      End
      Begin VB.Menu MnuPropertiesPop 
         Caption         =   "&Properties"
      End
      Begin VB.Menu MnuInsertAsLinkPop 
         Caption         =   "&Insert As Link"
      End
      Begin VB.Menu MnuInsertAsImagePop 
         Caption         =   "&Insert As Image"
      End
      Begin VB.Menu MnuFilter 
         Caption         =   "&Filter"
      End
      Begin VB.Menu MnuCRefresh 
         Caption         =   "&Refresh"
      End
   End
   Begin VB.Menu MnuD 
      Caption         =   "D"
      Visible         =   0   'False
      Begin VB.Menu MnuEditDPop 
         Caption         =   "&Edit"
      End
      Begin VB.Menu MnuDInsertAsLink 
         Caption         =   "&Insert As Link"
      End
      Begin VB.Menu MnuDInsertAsImage 
         Caption         =   "&Insert As Image"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Thanks to all that I got code and or ideas from on PSC
'Especially David Smejkal for his YZYFTP article and code AWESOME!!!!
'Also a thanks to the guy who made the resizing possible do not know your name sorry.
'Jason Shimkoski for his undo and redo code
'And a big thanks to everyone else who has contributed to helping me build
'If you feel that any ideas came from your code please send an email to me at erpa14@aol.com
'I will review what you have sent in and will be glad to include your name in the credits
'If your code was in fact included in making this project

Option Explicit
Dim FileName As String, OldName As String 'global filename used for saving and uploading of files to and from the server
Dim lngStart As Long
Const sReadBuffer = 1024
Private hFile As Long
Private fPat As String
Private Const SW_SHOW = 5
Private Const SEE_MASK_INVOKEIDLIST = &HC
Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function ShellExecuteEx Lib "shell32.dll" (ByRef S As SHELLEXECUTEINFO) As Long
Const SW_SHOWNORMAL = 1
Dim FormHeight As Single
Dim FormWidth As Single

'Stuff for undo
Dim gblnIgnoreChange As Boolean
Dim gintIndex As Integer
Dim gstrStack(10000) As String
'Global save
Public Saved As Boolean
Public CtlKey As Boolean

Private Declare Function SendMessageLong Lib "user32" Alias _
        "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
        ByVal wParam As Long, lParam As Long) As Long

Private Const EM_LINEFROMCHAR = &HC9
Private Const EM_LINEINDEX = &HBB
Private Const EM_LINELENGTH = &HC1

Dim X1 As Single

Dim Y1 As Single

Dim Start As Boolean

Private Sub CmdClose_Click()
MnuClose_Click
End Sub

Private Sub CmdClose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
CmdClose.BorderStyle = 1
End Sub

Private Sub CmdClose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
CmdClose.BorderStyle = 0
End Sub

Private Sub cmdFind_Click()
Form2.Show
End Sub

Private Sub cmdFind_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdFind.BorderStyle = 1
End Sub

Private Sub cmdFind_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdFind.BorderStyle = 0
End Sub

Private Sub CmdReplace_Click()
Form3.Show
End Sub

Private Sub cmdReplace_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
CmdReplace.BorderStyle = 1
End Sub

Private Sub cmdReplace_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
CmdReplace.BorderStyle = 0
End Sub

Private Sub CmdView_Click()
MnuResourceTab_Click
End Sub

Private Sub CmdView_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
CmdView.BorderStyle = 1
End Sub

Private Sub CmdView_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
CmdView.BorderStyle = 0
End Sub

Private Sub Drive1_Change()
On Error Resume Next ' if you dont use this and you choose a drive that is not ready it will return an error
Dir1.Path = Drive1.Drive 'this does not work in design mode!
End Sub

Private Sub File1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then ' used to popup a menu
PopupMenu MnuC
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyControl Then
CtlKey = True
ElseIf KeyCode = vbKeyF6 And (Shift And vbAltMask) Then
KeyCode = 0
End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyControl Then
CtlKey = False
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Make sure everything gets unloaded so that we do not have a form still running, taking up resources
DeleteAutosave
Unload Me
Unload Form2
Unload Form3
Unload Form4
Unload Form5
Unload Form6
Unload Form7
Unload Form8
Unload Form9
Unload Form10
Unload Form11
Unload Form12
Unload Form13
Unload Form14
Unload Form15
Unload Form16
Unload Form17
Unload Form18
End 'Make sure everything gets unloaded just to make sure
End Sub

Private Sub Frame3_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
X1 = X
Y1 = Y
Start = True
End Sub

Private Sub Frame3_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim X2, Y2 As Single
On Error GoTo MoveErr:
X2 = X
Y2 = Y
If Button = 1 Then

With Frame3(Index)

.Move Frame3(Index).Left - X1 + X2


If Frame3(Index).Left <= 1000 Then

Frame3(Index).Left = 1000
TabStrip1.Width = Frame3(Index).Left
TabStrip2.Left = Frame3(Index).Left + Frame3(Index).Width
TabStrip2.Width = Form1.Width - TabStrip1.Width - Frame3(Index).Width - 100
Drive1.Width = TabStrip1.Width - 200
Dir1.Width = TabStrip1.Width - 200
File1.Width = TabStrip1.Width - 200
List1.Width = TabStrip1.Width - 200
List2.Width = TabStrip1.Width - 200
List3.Width = TabStrip1.Width - 200
WebBrowser1.Width = TabStrip2.Width - 500
RichTextBox1.Width = TabStrip2.Width - 500
WebBrowser1.Left = TabStrip2.Left + 400
RichTextBox1.Left = TabStrip2.Left + 400
cmdFind.Left = WebBrowser1.Left - 300
CmdReplace.Left = WebBrowser1.Left - 300
CmdClose.Left = WebBrowser1.Left - 300
CmdView.Left = WebBrowser1.Left - 300



Else


If Frame3(Index).Left >= 8000 Then

Frame3(Index).Left = 7000
TabStrip1.Width = Frame3(Index).Left
TabStrip2.Left = Frame3(Index).Left + Frame3(Index).Width
TabStrip2.Width = Form1.Width - TabStrip1.Width - Frame3(Index).Width - 100
Drive1.Width = TabStrip1.Width - 200
Dir1.Width = TabStrip1.Width - 200
File1.Width = TabStrip1.Width - 200
List1.Width = TabStrip1.Width - 200
List2.Width = TabStrip1.Width - 200
List3.Width = TabStrip1.Width - 200
WebBrowser1.Width = TabStrip2.Width - 500
RichTextBox1.Width = TabStrip2.Width - 500
WebBrowser1.Left = TabStrip2.Left + 400
RichTextBox1.Left = TabStrip2.Left + 400
cmdFind.Left = WebBrowser1.Left - 300
CmdReplace.Left = WebBrowser1.Left - 300
CmdClose.Left = WebBrowser1.Left - 300
CmdView.Left = WebBrowser1.Left - 300



Else
TabStrip1.Width = Frame3(Index).Left
TabStrip2.Left = Frame3(Index).Left + Frame3(Index).Width
TabStrip2.Width = Form1.Width - TabStrip1.Width - Frame3(Index).Width - 100
Drive1.Width = TabStrip1.Width - 200
Dir1.Width = TabStrip1.Width - 200
File1.Width = TabStrip1.Width - 200
List1.Width = TabStrip1.Width - 200
List2.Width = TabStrip1.Width - 200
List3.Width = TabStrip1.Width - 200
WebBrowser1.Width = TabStrip2.Width - 500
RichTextBox1.Width = TabStrip2.Width - 500
WebBrowser1.Left = TabStrip2.Left + 400
RichTextBox1.Left = TabStrip2.Left + 400
cmdFind.Left = WebBrowser1.Left - 300
CmdReplace.Left = WebBrowser1.Left - 300
CmdClose.Left = WebBrowser1.Left - 300
CmdView.Left = WebBrowser1.Left - 300
End If

End If

End With

End If


If TabStrip1.Visible = True Then

If Frame2.Visible = True Then
Frame1.Width = RichTextBox1.Width - 650
Frame2.Width = RichTextBox1.Width - 650
Frame1.Left = RichTextBox1.Left + 100
Frame2.Left = RichTextBox1.Left + 100
End If
Else
Frame1.Width = RichTextBox1.Width - 350
Frame2.Width = RichTextBox1.Width - 350
Frame2.Left = RichTextBox1.Left + 100
End If

MoveErr:
Exit Sub

End Sub

Private Sub List1_DblClick() ' this changes the directory that you are currently at will connected via ftp
Screen.MousePointer = vbHourglass
If List1.List(List1.ListIndex) = "/" Then
Klic = "/"
adr = Klic & fPat
Else
Klic = List1.List(List1.ListIndex)
adr = Klic & fPat
End If
List3.Clear
Form8.List
FtpSetCurrentDirectory session, adr
Screen.MousePointer = vbDefault
End Sub

Private Sub List2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then ' used for popup menu for the list of ftp sites saved for reuse utilizing the writeprivateprofile string
If List2.ListCount = 1 Then
MnuConnect.Enabled = False
MnuDeleteServer.Enabled = False
MnuServerProperties.Enabled = False
PopupMenu MnuB
Else
MnuConnect.Enabled = True
MnuDeleteServer.Enabled = True
MnuServerProperties.Enabled = True
PopupMenu MnuB
End If
End If
End Sub

Private Sub List3_DblClick() 'used for downloading a file from an open session of ftp
If session = 0 Or server = 0 Then
MsgBox "Not connected to any server!", vbInformation, App.ProductName
Exit Sub
End If
If List3.List(List3.ListIndex) = "" Then
MsgBox "No file selected!!", vbOKOnly, App.ProductName
Exit Sub
Else
Form9.List1.AddItem (List3.List(List3.ListIndex))
Form9.Command3.Caption = "Download"
Form9.Show
End If
End Sub

Private Sub List3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then 'creates a popup menu of available functions when connected to a session via ftp
PopupMenu MnuD
End If
End Sub

Private Sub List5_DblClick()
RichTextBox1.SelText = RichTextBox1.SelText & List5.List(List5.ListIndex) & ">"
List5.Visible = False
RichTextBox1.SetFocus
End Sub

Private Sub List5_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
List5_DblClick
End If
End Sub

Private Sub List5_KeyPress(KeyAscii As Integer)
If KeyAscii = 47 Then
List5.Visible = False
RichTextBox1.SetFocus
RichTextBox1.SelText = "/"
End If
End Sub

Private Sub List5_LostFocus()
List5.Visible = False
RichTextBox1.SetFocus
End Sub

Private Sub List6_DblClick()
RichTextBox1.SelText = RichTextBox1.SelText & List6.List(List6.ListIndex)
List6.Visible = False
RichTextBox1.SetFocus
End Sub

Private Sub List6_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
List6_DblClick
End If
End Sub

Private Sub List6_LostFocus()
List6.Visible = False
RichTextBox1.SetFocus
End Sub

Private Sub MnuAbout_Click()
Form11.Show 'shows the about form
End Sub

Private Sub MnuAddServer_Click()
Form5.Show 'shows the addserver form
End Sub

Private Sub MnuARedo_Click()
'This is the basic redo stuff.
    gblnIgnoreChange = True
    gintIndex = gintIndex + 1
    On Error Resume Next
    RichTextBox1.TextRTF = gstrStack(gintIndex)
    gblnIgnoreChange = False
End Sub

Private Sub MnuAutoSave_Click()
RichTextBox1.LoadFile (App.Path & "\" & "Temp.Save") 'loads the last saved file via autosave
End Sub

Private Sub MnuClose_Click()
Dim RetVal As String
If Len(RichTextBox1.Text) = Len(RichTextBox2.Text) Or Len(RichTextBox1.Text) = 0 Then
RichTextBox1.Text = ""
RichTextBox2.Text = ""
Form1.Caption = App.ProductName & " - Untitled Document"
FileName = ""
GetGlobals
Exit Sub
Else
RetVal = MsgBox("This document has changed do you wish to save it?", vbYesNoCancel + vbInformation, App.ProductName)
If RetVal = vbYes Then
SaveFile
Else
If RetVal = vbNo Then
RichTextBox1.Text = ""
RichTextBox2.Text = ""
Form1.Caption = App.ProductName & " - Untitled Document"
FileName = ""
GetGlobals
Else
If RetVal = vbCancel Then
Exit Sub
End If
End If
End If
End If
End Sub

Private Sub MnuClosePop_Click()
Dim RetVal As String
If Len(RichTextBox1.Text) = Len(RichTextBox2.Text) Or Len(RichTextBox1.Text) = 0 Then
RichTextBox1.Text = ""
RichTextBox2.Text = ""
Form1.Caption = App.ProductName & " - Untitled Document"
FileName = ""
GetGlobals
Exit Sub
Else
RetVal = MsgBox("This document has changed do you wish to save it?", vbYesNoCancel + vbInformation, App.ProductName)
If RetVal = vbYes Then
SaveFile
Else
If RetVal = vbNo Then
RichTextBox1.Text = ""
RichTextBox2.Text = ""
Form1.Caption = App.ProductName & " - Untitled Document"
FileName = ""
GetGlobals
Else
If RetVal = vbCancel Then
Exit Sub
End If
End If
End If
End If
End Sub

Private Sub MnuConnect_Click() ' shows the connection form for ftp

If List2.List(List2.ListIndex) <> "Remote Servers" Then
List1.Clear
List3.Clear
Form8.Show
End If

End Sub


Private Sub MnuConvertToTable_Click()
If RichTextBox1.SelText = "" Or Len(RichTextBox1.Text) = 0 Then
Exit Sub
Else
RichTextBox1.SelText = "<Table>" & vbCrLf & "    <Tr><Td>" & RichTextBox1.SelText & "</Td></Tr>" & vbCrLf & "</Table>"
End If
End Sub

Private Sub MnuCopyPop_Click()
Clipboard.Clear
Clipboard.SetText (RichTextBox1.SelText)
End Sub

Private Sub MnuCRefresh_Click()
File1.Refresh
End Sub

Private Sub MnuCutPop_Click()
Clipboard.SetText (RichTextBox1.SelText)
RichTextBox1.SelText = ""
End Sub

Private Sub MnuDeletePop_Click()
RichTextBox1.SelText = ""
Clipboard.Clear
End Sub

Private Sub MnuDeleteServer_Click()
Dim RetVal As String
If List2.ListCount < 1 Or List2.List(List2.ListIndex) = "Remote Servers" Then
MsgBox ("You cannot remove this item"), vbOKOnly + vbInformation, App.ProductName
Else
If List2.List(List2.ListIndex) = "" Then
MsgBox ("Please select a server to delete!"), vbOKOnly, App.ProductName
Else
RetVal = MsgBox("Delete server " & Left(List2.List(List2.ListIndex), Len(List2.List(List2.ListIndex)) - 4) & " ?", vbOKCancel + vbCritical, App.ProductName)
If RetVal = vbOK Then
DeleteServer
File2.Refresh
List2.Refresh
MnuRefreshServes_Click
Else
'Nada
End If
End If
End If
End Sub

Private Sub MnuDInsertAsImage_Click()
If List3.ListIndex = -1 Then
MsgBox ("Insert an image of what?"), vbInformation, App.ProductName
Exit Sub
Else
RichTextBox1.SelText = "<img src=""" & List3.List(List3.ListIndex) & """>"
End If
End Sub

Private Sub MnuDInsertAsLink_Click()
If List3.ListIndex = -1 Then
MsgBox ("Insert a link of what?"), vbInformation, App.ProductName
Exit Sub
Else
RichTextBox1.SelText = "<a href=""" & List3.List(List3.ListIndex) & """>" & List3.List(List3.ListIndex) & "</a>"
End If
End Sub

Private Sub MnuDisconnect_Click()
InternetCloseHandle server
InternetCloseHandle session
server = 0
session = 0
List1.Clear
List3.Clear
List1.BackColor = &H8000000F
List1.ForeColor = &H8000000F
List3.BackColor = &H8000000F
List3.ForeColor = &H8000000F
End Sub

Private Sub MnuDSaveToServer_Click()
Dim I As Integer
Dim RetVal As String
For I = 0 To List4.ListCount - 1
RetVal = MsgBox("Save to server: " & Adresa & " as " & List4.List(I), vbOKCancel + vbInformation, App.ProductName)
If RetVal = vbOK Then
SaveFW
Upload
Else
Exit Sub
End If
Next I
End Sub

Private Sub MnuLowerCase_Click()
RichTextBox1.SelText = LCase(RichTextBox1.SelText)
End Sub

Private Sub MnuPopViewWithExternal_Click()
Form15.Show
End Sub

Private Sub MnuRedo_Click()
'This is the basic redo stuff.
gblnIgnoreChange = True
gintIndex = gintIndex + 1
On Error Resume Next
RichTextBox1.TextRTF = gstrStack(gintIndex)
gblnIgnoreChange = False
End Sub

Private Sub MnuSaveAsTemplate_Click()
Form12.Show
Form12.TabStrip1.Tabs(2).Selected = True
End Sub

Private Sub MnuSaveToServer_Click()
Dim I As Integer
Dim RetVal As String
For I = 0 To List4.ListCount - 1
RetVal = MsgBox("Save to server: " & Adresa & " as " & List4.List(I), vbOKCancel + vbInformation, App.ProductName)
If RetVal = vbOK Then
SaveFW
Upload
Else
Exit Sub
End If
Next I
End Sub

Private Sub MnuEditDPop_Click()
If session = 0 Or server = 0 Then
MsgBox "Not connected to any server!", vbInformation, App.ProductName
Exit Sub
End If
If List3.List(List3.ListIndex) = "" Then
MsgBox "No file selected!!", vbOKOnly, App.ProductName
Exit Sub
Else
Form9.List1.AddItem (List3.List(List3.ListIndex))
Form9.Command3.Caption = "Download"
Form9.Show
End If
End Sub

Private Sub MnuEditPop_Click()
File1_DblClick
End Sub

Private Sub MnuExit_Click()
End
End Sub

Private Sub MnuFilter_Click()
Form10.Show
End Sub

Private Sub MnuFind_Click()
Form2.Show
End Sub

Private Sub MnuInsertAsImagePop_Click()
If File1.ListIndex = -1 Then
MsgBox ("Insert an image of what?"), vbInformation, App.ProductName
Exit Sub
Else
RichTextBox1.SelText = "<img src=""" & File1.FileName & """>"
End If
End Sub

Private Sub MnuInsertAsLinkPop_Click()
If File1.ListIndex = -1 Then
MsgBox ("Insert a link of what?"), vbInformation, App.ProductName
Exit Sub
Else
RichTextBox1.SelText = "<a href=""" & File1.FileName & """>" & File1.FileName & "</a>"
End If
End Sub

Private Sub MnuInsertTagPop_Click()
Form4.Show
End Sub

Private Sub MnuMerkatumOnline_Click()
Dim sURL As String
sURL = "http://www.merkatum.com"
If Not StartNewBrowser(sURL) Then
MsgBox ("Your default browser could not be opened. & vbcrlf & Please navigate to: http://www.merkatum.com"), vbInformation, App.ProductName
End If
End Sub
Private Sub MnuNewPop_Click()
Dim RetVal As String
If Saved = True Then
RichTextBox1.Text = ""
RichTextBox2.Text = ""
GetGlobals
Form1.Caption = App.ProductName & " - Untitled Document"
FileName = ""
Else
RetVal = MsgBox("This document has changed do you wish to save it?", vbOKCancel + vbInformation, App.ProductName)
If RetVal = vbOK Then
SaveFile
Else
Form1.Caption = App.ProductName & " - Untitled Document"
FileName = ""
End If
End If
End Sub

Private Sub MnuOnlineSupport_Click()
Dim sURL As String
sURL = "http://www.merkatum.com/support/login.asp"
If Not StartNewBrowser(sURL) Then
MsgBox ("Your default browser could not be opened. & vbcrlf & Please navigate to: http://www.merkatum.com"), vbInformation, App.ProductName
End If
End Sub

Private Sub MnuOpenFromWeb_Click()
Form7.Show
End Sub

Private Sub MnuOpenFromWebPop_Click()
Form7.Show
End Sub

Private Sub MnuOpenPop_Click()
MnuOpen_Click
End Sub

Private Sub MnuPastePop_Click()
Dim p As String
p = Clipboard.GetText
RichTextBox1.SelText = p
End Sub

Private Sub MnuPropertiesPop_Click()
Dim ShInfo As SHELLEXECUTEINFO
Dim strPath As String
strPath = File1.Path
If File1.FileName = "" Then
MsgBox ("Properties of what?"), vbInformation, App.ProductName
Exit Sub
End If
With ShInfo
.cbSize = LenB(ShInfo)
.lpFile = strPath & "\" & File1.FileName
.nShow = SW_SHOW
.fMask = SEE_MASK_INVOKEIDLIST
.lpVerb = "properties"
End With
ShellExecuteEx ShInfo
End Sub

Private Sub MnuRefreshServes_Click()
List2.Clear
List2.AddItem ("Remote Servers")
GetGlobals
End Sub

Private Sub MnuReplace_Click()
Form3.Show
End Sub

Private Sub MnuSave_Click()
Call SaveFile
End Sub

Private Sub MnuSaveAs_Click()
On Error GoTo Nada
SaveDlg.CancelError = False
SaveDlg.Filter = "Auctioneer Files (*.tem)|*.tem|Vtrader Files (*.vtf)|*.vtf |Text Files (*.txt)|*.txt|Html Files (*.html)|*.html|All Files (*.*)|*.*"
SaveDlg.ShowSave
If SaveDlg.FileName = "" Then
SaveDlg.FileName = "c:\" & "Error.tmp"
Else
Open SaveDlg.FileName For Output As 1
Print #1, RichTextBox1.Text
Close
Saved = True
End If
Nada:
Exit Sub
End Sub

Private Sub MnuSaveAsPop_Click()
MnuSaveAs_Click
End Sub

Private Sub MnuSavePop_Click()
Call SaveFile
End Sub
Private Sub MnuServerProperties_Click()
Form6.Show
End Sub

Private Sub MnuSettings_Click()
Form12.Show
End Sub

Private Sub MnuSqlBuiler_Click()
Form14.Show
End Sub

Private Sub MnuTagChooser_Click()
Form4.Show
End Sub

Private Sub MnuToggleEB_Click()
If TabStrip2.Tabs(1).Selected = True Then
TabStrip2.Tabs(2).Selected = True
Else
TabStrip2.Tabs(1).Selected = True
End If
End Sub

Private Sub MnuUndoPop_Click()
'SendMessage RichTextBox1.hwnd, EM_UNDO, 0, 0&
If gintIndex = 0 Then Exit Sub
'This is the basic undo stuff.
gblnIgnoreChange = True
gintIndex = gintIndex - 1
On Error Resume Next
RichTextBox1.TextRTF = gstrStack(gintIndex)
gblnIgnoreChange = False
End Sub

Private Sub MnuUpperCase_Click()
RichTextBox1.SelText = UCase(RichTextBox1.SelText)
End Sub

Private Sub MnuViewWithEXternalBrowsers_Click()
Form15.Show
End Sub

Private Sub RichTextBox1_Change()
'updates the Undo and Redo
If Not gblnIgnoreChange Then
gintIndex = gintIndex + 1
gstrStack(gintIndex) = RichTextBox1.TextRTF
End If
StatusBar1.Panels(4).Text = Form1.Caption
End Sub

Private Sub RichTextBox1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode <> 58 And KeyCode <> 60 Then
If Form12.C = True Then
GetCaretPos A
Frame1.Top = A.Y * 15 + RichTextBox1.Top + 300
Frame2.Top = A.Y * 15 + RichTextBox1.Top
Else
' the user does not want the cursor to appear
End If
Else
' do not allow it to to do anything as the keypress event takes over.
End If
End Sub

Private Sub RichTextBox1_KeyPress(KeyAscii As Integer)

If KeyAscii = 58 Then ' this is for the : character
GetCaretPos A
List6.Left = A.X * 15 + RichTextBox1.Left + 50
List6.Top = A.Y * 15 + RichTextBox1.Top + 300
List6.Visible = True
List6.SetFocus
End If

If KeyAscii = 60 Then 'This is for the < character
GetCaretPos A
List5.Left = A.X * 15 + RichTextBox1.Left + 50
List5.Top = A.Y * 15 + RichTextBox1.Top + 300
List5.Visible = True
List5.SetFocus
End If
End Sub

Private Sub RichTextBox1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
PopupMenu MnuA
Else
If Form12.C = True Then
GetCaretPos A
Frame1.Top = A.Y * 15 + RichTextBox1.Top + 300
Frame2.Top = A.Y * 15 + RichTextBox1.Top
Else
End If
End If
End Sub

Private Sub RichTextBox1_SelChange()
GetEditStatus
End Sub

Public Sub GetEditStatus()
Dim lLine As Long, lCol As Long
Dim cCol As Long, lChar As Long, I As Long
lChar = RichTextBox1.SelStart + 1
lLine = 1 + SendMessageLong(RichTextBox1.hwnd, EM_LINEFROMCHAR, _
RichTextBox1.SelStart, 0&)
cCol = SendMessageLong(RichTextBox1.hwnd, EM_LINELENGTH, lChar - 1, 0&)
I = SendMessageLong(RichTextBox1.hwnd, EM_LINEINDEX, lLine - 1, 0&)
lCol = lChar - I
StatusBar1.Panels(1).Text = "Line: " & lLine & ", Col: " & lCol
End Sub

Private Sub Timer1_Timer()
AutoSave
End Sub

Private Sub Timer2_Timer()
StatusBar1.Panels(2).Text = Time & " " & Date
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
Select Case Button.Key
Case "New"
MnuNew_Click
Case "Open"
MnuOpen_Click
Case "Save"
MnuSave_Click
Case "Find"
MnuFind_Click
Case "Replace"
MnuReplace_Click
Case "Cut"
MnuCut_Click
Case "Copy"
MnuCopy_Click
Case "Paste"
MnuPaste_Click
Case "Delete"
MnuDelete_Click
Case "Undo"
MnuUndo_Click
Case "Redo"
MnuRedo_Click
Case "Align Left"
RichTextBox1.SelAlignment = rtfLeft
Case "Center"
RichTextBox1.SelAlignment = rtfCenter
Case "Align Right"
RichTextBox1.SelAlignment = rtfRight
Case "Bold"
If RichTextBox1.SelBold = False Then
RichTextBox1.SelBold = True
Toolbar1.Buttons("Bold").Value = tbrPressed
Else
RichTextBox1.SelBold = False
Toolbar1.Buttons("Bold").Value = tbrUnpressed
End If
Case "Italic"
If RichTextBox1.SelItalic = False Then
RichTextBox1.SelItalic = True
Toolbar1.Buttons("Italic").Value = tbrPressed
Else
RichTextBox1.SelItalic = False
Toolbar1.Buttons("Italic").Value = tbrUnpressed
End If
Case "Underline"
If RichTextBox1.SelUnderline = False Then
RichTextBox1.SelUnderline = True
Toolbar1.Buttons("Underline").Value = tbrPressed
Else
RichTextBox1.SelUnderline = False
Toolbar1.Buttons("Underline").Value = tbrUnpressed
End If
Case "Small Caps"
RichTextBox1.SelText = LCase(RichTextBox1.SelText)
Case "Strike Through"
If RichTextBox1.SelStrikeThru = False Then
RichTextBox1.SelStrikeThru = True
Toolbar1.Buttons("Strike Through").Value = tbrPressed
Else
RichTextBox1.SelStrikeThru = False
Toolbar1.Buttons("Strike Through").Value = tbrUnpressed
End If
Case "Properties"
MnuDocumentWeight_Click
End Select
End Sub

Private Sub MnuCopy_Click()
Clipboard.Clear
Clipboard.SetText (RichTextBox1.SelText)
End Sub

Private Sub MnuCut_Click()
Clipboard.SetText (RichTextBox1.SelText)
RichTextBox1.SelText = ""
End Sub

Private Sub MnuDelete_Click()
RichTextBox1.SelText = ""
Clipboard.Clear
End Sub

Private Sub MnuDocumentWeight_Click()
MsgBox (Len(RichTextBox1.Text) & " Bytes" & vbCrLf & "Note This Does Not Take Into Consideration Any Images!"), vbInformation + vbOKOnly, App.ProductName
End Sub

Private Sub MnuIndent_Click()
SendKeys "{Tab}"
End Sub

Private Sub MnuNew_Click()
Dim RetVal As String
If Saved = True Then
RichTextBox1.Text = ""
RichTextBox2.Text = ""
GetGlobals
Form1.Caption = App.ProductName & " - Untitled Document"
FileName = ""
Else
RetVal = MsgBox("This document has changed do you wish to save it?", vbOKCancel + vbInformation, App.ProductName)
If RetVal = vbOK Then
SaveFile
Else
Form1.Caption = App.ProductName & " - Untitled Document"
FileName = ""
End If
End If
End Sub

Private Sub MnuOpen_Click()
Dim Buffer1 As String, Buffer2 As String, CRLF As String
Dim FileNum As Integer
CRLF = Chr$(13) + Chr$(10)
OpenDlg.Filter = "Auctioneer Files (*.tem)|*.tem|Vtrader Files (*.vtf)|*.vtf |Text Files (*.txt)|*.txt|Html Files (*.html)|*.html|All Files (*.*)|*.*"
OpenDlg.ShowOpen
If OpenDlg.FileName = "" Then Exit Sub
FileName = OpenDlg.FileName
FileNum = FreeFile
Open FileName For Input As FileNum
Do While Not EOF(FileNum)
Line Input #FileNum, Buffer1
Buffer2 = Buffer2 & Buffer1 & CRLF
Loop
Close FileNum
RichTextBox1.Text = Buffer2
RichTextBox2.Text = Buffer2
Form1.Caption = App.ProductName & " - [ " & OpenDlg.FileName & " ]"
Saved = False
StatusBar1.Panels(4).Text = Form1.Caption
End Sub

Private Sub MnuPaste_Click()
Dim p As String
p = Clipboard.GetText
RichTextBox1.SelText = p

End Sub

Private Sub MnuSelectAll_Click()
RichTextBox1.SelStart = 0
RichTextBox1.SelLength = Len(RichTextBox1.Text)
End Sub

Private Sub MnuUndo_Click()
'SendMessage RichTextBox1.hwnd, EM_UNDO, 0, 0&
If gintIndex = 0 Then Exit Sub
    
    'This is the basic undo stuff.
    gblnIgnoreChange = True
    gintIndex = gintIndex - 1
    On Error Resume Next
    RichTextBox1.TextRTF = gstrStack(gintIndex)
   gblnIgnoreChange = False

End Sub

Private Sub MnuUnIndent_Click()
SendKeys "{BackSpace}"
End Sub

Private Sub MnuWordWrap_Click()
If MnuWordWrap.Checked = False Then
MnuWordWrap.Checked = True
RichTextBox1.RightMargin = RichTextBox1.Width - 500
Else
MnuWordWrap.Checked = False
RichTextBox1.RightMargin = Form1.Width + 100000
End If
StatusBar1.Panels(1).Text = "Line:" & _
RichTextBox1.GetLineFromChar(Len(RichTextBox1.Text))
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub File1_DblClick()
Saved = False
Form1.Caption = App.ProductName
Dim Fil As String
FileName = File1.FileName
If File1.ListIndex = -1 Then
MsgBox ("Please select an item!"), vbInformation, App.ProductName
Exit Sub
End If
If Right(File1.Path, 1) <> "\" Then ' if the end of the filename ends with a \
If Left(File1.Path, 1) <> "\" Then 'if the file is in the root directory
Fil = File1.Path & "\" & File1.FileName 'then set the string to not end with a \
Screen.MousePointer = vbHourglass
RichTextBox1.LoadFile (Fil)
RichTextBox2.LoadFile (Fil)
Screen.MousePointer = vbDefault
Form1.Caption = App.ProductName & " - [ " & File1.Path & "\" & File1.FileName & " ] "
Else
Fil = File1.Path & File1.FileName
Screen.MousePointer = vbHourglass
RichTextBox1.LoadFile (Fil)
RichTextBox2.LoadFile (Fil)
Screen.MousePointer = vbDefault
Form1.Caption = App.ProductName & " -  [ " & File1.Path & "\" & File1.FileName & " ] "
End If
Else
Fil = Dir1.Path & File1.FileName
Screen.MousePointer = vbHourglass
RichTextBox1.LoadFile (Fil)
RichTextBox2.LoadFile (Fil)
Screen.MousePointer = vbDefault
Form1.Caption = App.ProductName & " - [ " & Drive1.Drive & "\" & File1.FileName & " ] "
End If
RichTextBox1.SetFocus
StatusBar1.Panels(4).Text = Form1.Caption
End Sub

Private Sub Form_Load()
FormHeight = Form1.Height
FormWidth = Form1.Width
Form1.File2.Path = App.Path
GetGlobals
LoadLists
GetSettings
Form12.SetCursor
TabStrip1.Tabs(1).Selected = True
TabStrip2.Tabs(1).Selected = True
FileName = ""
OldName = ""
Form1.Caption = App.ProductName & " - Untitled Document"
Saved = True
MnuAutoSave.Enabled = False
StatusBar1.Panels(4).Text = Form1.Caption
StatusBar1.Panels(2).Text = Time & " " & Date
If Form12.Check1.Value = 0 Then
StatusBar1.Panels(3).Text = "Autosave is off"
Else
StatusBar1.Panels(3).Text = "Autosave is on"
End If
GetEditStatus
cmdFind.Left = WebBrowser1.Left - 300
CmdReplace.Left = WebBrowser1.Left - 300
CmdClose.Left = WebBrowser1.Left - 300
CmdView.Left = WebBrowser1.Left - 300
End Sub

Private Sub Form_Resize()
On Error Resume Next
If Form1.WindowState = vbMinimized Then
Exit Sub
End If
If Form1.Height < 7000 Then
Form1.Height = 7000
End If
If Form1.Width < 8000 Then
Form1.Width = 8000
End If

If Form1.WindowState <> 1 Then
If T1V = True Then
TabStrip1.Height = Form1.Height - Toolbar1.Height - StatusBar1.Height - 800
TabStrip2.Height = Form1.Height - Toolbar1.Height - StatusBar1.Height - 800

TabStrip2.Width = Form1.Width - TabStrip1.Width - 100
RichTextBox1.Width = TabStrip2.Width - 500
RichTextBox1.Height = TabStrip2.Height - 500
WebBrowser1.Width = TabStrip2.Width - 500
WebBrowser1.Height = TabStrip2.Height - 500

Frame1.Width = RichTextBox1.Width - 350
Frame1.Left = RichTextBox1.Left + 100
Frame2.Width = RichTextBox1.Width - 350
Frame2.Left = RichTextBox1.Left + 100

File1.Height = TabStrip1.Height - Drive1.Height - Dir1.Height - StatusBar1.Height - 200
List3.Height = TabStrip1.Height - Drive1.Height - Dir1.Height - List1.Height - StatusBar1.Height - 400

Else

TabStrip2.Height = Form1.Height - Toolbar1.Height - StatusBar1.Height - 800
TabStrip2.Width = Form1.Width - 100
RichTextBox1.Width = TabStrip2.Width - 500
RichTextBox1.Height = TabStrip2.Height - 500
WebBrowser1.Width = TabStrip2.Width - 500
WebBrowser1.Height = TabStrip2.Height - 500

End If
Else
' Nothing
End If
Frame3(1).Height = Form1.Height - Toolbar1.Height - StatusBar1.Height - 1200
End Sub

Private Sub MnuResourceTab_Click()
If MnuResourceTab.Checked = True Then
MnuResourceTab.Checked = False
TabStrip1.Visible = False
Drive1.Visible = False
Dir1.Visible = False
File1.Visible = False
List1.Visible = False
List2.Visible = False
List3.Visible = False
TabStrip2.Width = Form1.Width - 200
TabStrip2.Left = 0
WebBrowser1.Left = TabStrip2.Left + 400
RichTextBox1.Left = TabStrip2.Left + 400
WebBrowser1.Width = TabStrip2.Width - 500
RichTextBox1.Width = TabStrip2.Width - 500

Frame1.Width = RichTextBox1.Width - 650
Frame1.Left = RichTextBox1.Left + 100
Frame2.Width = RichTextBox1.Width - 650
Frame2.Left = RichTextBox1.Left + 100

Frame3(1).Visible = False

Else
MnuResourceTab.Checked = True
TabStrip1.Visible = True
Frame3(1).Visible = True
If TabStrip1.Tabs(1).Selected = True Then
Drive1.Visible = True
Dir1.Visible = True
File1.Visible = True
List1.Visible = False
List2.Visible = False
List3.Visible = False
Else
Drive1.Visible = False
Dir1.Visible = False
File1.Visible = False
List1.Visible = True
List2.Visible = True
List3.Visible = True
End If
TabStrip2.Width = Form1.Width - 200 - TabStrip1.Width
TabStrip2.Left = TabStrip1.Width
WebBrowser1.Left = TabStrip2.Left + 400
RichTextBox1.Left = TabStrip2.Left + 400
WebBrowser1.Width = TabStrip2.Width - 500
RichTextBox1.Width = TabStrip2.Width - 500

Frame1.Width = RichTextBox1.Width - 650
Frame1.Left = RichTextBox1.Left + 100
Frame2.Width = RichTextBox1.Width - 650
Frame2.Left = RichTextBox1.Left + 100

End If
cmdFind.Left = WebBrowser1.Left - 300
CmdReplace.Left = WebBrowser1.Left - 300
CmdClose.Left = WebBrowser1.Left - 300
CmdView.Left = WebBrowser1.Left - 300
End Sub

Private Sub TabStrip1_Click()
If TabStrip1.Tabs(1).Selected = True Then
Drive1.Visible = True
Dir1.Visible = True
File1.Visible = True
List1.Visible = False
List2.Visible = False
List3.Visible = False
Else
Drive1.Visible = False
Dir1.Visible = False
File1.Visible = False
List1.Visible = True
List2.Visible = True
List3.Visible = True
End If
End Sub

Private Sub TabStrip2_Click()
If TabStrip2.Tabs(1).Selected = True Then
RichTextBox1.Visible = True
WebBrowser1.Visible = False

If Form12.C = True Then
Frame1.Visible = True
Frame2.Visible = True
Else
Frame1.Visible = False
Frame2.Visible = False
End If

Else
SaveTmp
WebBrowser1.Navigate ("c:\e.tmp")
WebBrowser1.Visible = True
RichTextBox1.Visible = False
List5.Visible = False

Frame1.Visible = False
Frame2.Visible = False

End If
End Sub

Public Function T1V() As Boolean
If TabStrip1.Visible = True Then
T1V = True
Else
T1V = False
End If
End Function

Private Sub SaveFile()
Dim FileNum As Integer, Buffer As String
Buffer = RichTextBox1.Text
If FileName = "" Then
If Form1.List4.List(0) = "" Then
SaveDlg.FileName = ""
SaveDlg.Filter = "Text Files (*.txt)|*.txt|Html Files (*.html)|*.html|Merkatum Files (*.vtf)|*.vtf"
SaveDlg.ShowSave
FileName = SaveDlg.FileName
If FileName = "" Then
FileName = OldName
Exit Sub
End If
Else
FileName = List4.List(0)
End If
Else
'
End If
FileNum = FreeFile
Open FileName For Output As FileNum
Print #FileNum, Buffer
Close FileNum
Saved = True
End Sub

Public Sub SaveTmp()
Dim FileNum As Integer, Buffer As String
Buffer = "<Html><Body>" & RichTextBox1.Text & "</Body></Html>" ' the opening and closing body and html tags are for files that may not begin with these tags like [!html:header!] & [!html:footer!]
FileName = "c:\" & "e.tmp"
FileNum = FreeFile
Open FileName For Output As FileNum
Print #FileNum, Buffer
Close FileNum
FileName = ""
End Sub

Function GetGlobals()
On Error Resume Next
Dim I As Integer
Dim StrFileName
Dim SzReturn As String
Dim SName As String
Dim TName As String
Dim Cursor As String
TName = Space(100)
SName = Space(5)
Cursor = Space(5)
List2.Clear
List2.AddItem ("Remote Servers")
For I = 0 To File2.ListCount - 1
List2.AddItem (File2.List(I))
Next I
File2.Refresh
List2.Refresh

StrFileName = App.Path & "\Satellite.ini"
GetPrivateProfileString "StartWithTemplate", "Yes,No", SzReturn, SName, Len(SName), StrFileName
GetPrivateProfileString "Template", "Template", SzReturn, TName, Len(TName), StrFileName
Form12.Check2.Value = SName
Form12.Text2.Text = TName

If SName = 1 Then
RichTextBox1.LoadFile TName
RichTextBox2.LoadFile TName

Else

RichTextBox1.Text = ""
RichTextBox2.Text = ""

End If
FileName = ""
End Function
Public Function LoadLists()
Dim ListBuffer As String
Dim FileNumber As Integer
FileName = App.Path & "\HtmlTags.dat"
FileNumber = FreeFile

Open FileName For Input As FileNumber

Do While Not EOF(FileNumber)

Line Input #FileNumber, ListBuffer
List5.AddItem ListBuffer
Form4.List1.AddItem ListBuffer
Loop
Close FileNumber
FileName = ""



Dim ListBuffer1 As String
Dim FileNumber1 As Integer
FileName = App.Path & "\VTags.dat"
FileNumber1 = FreeFile

Open FileName For Input As FileNumber1

Do While Not EOF(FileNumber1)

Line Input #FileNumber1, ListBuffer1
List6.AddItem ListBuffer1
Form4.List3.AddItem ListBuffer1
Loop
Close FileNumber1
FileName = ""

End Function

Public Function GetSettings()
Dim Path As String
Dim SzReturn As String
Dim Color As String
Dim Color2 As String
Dim Size As String
Dim Font As String
Color = Space(2)
Color2 = Space(2)
Size = Space(4)
Path = App.Path & "\Satellite.ini"

GetPrivateProfileString "BackColor", "Color", SzReturn, Color, Len(Color), Path

Form1.RichTextBox1.BackColor = Form12.Combo3.List(Color)


GetPrivateProfileString "ForeColor", "Color", SzReturn, Color2, Len(Color2), Path

Form1.RichTextBox1.SelStart = 0
Form1.RichTextBox1.SelLength = Len(Form1.RichTextBox1.Text)
Form1.RichTextBox1.SelColor = Form12.Combo3.List(Color2)

GetPrivateProfileString "FontSize", "Size", SzReturn, Size, Len(Size), Path

Form1.RichTextBox1.SelStart = 0
Form1.RichTextBox1.SelLength = Len(Form1.RichTextBox1.Text)
Form1.RichTextBox1.SelFontSize = Size

GetPrivateProfileString "FontName", "Font", SzReturn, Font, Len(Font), Path

Form1.RichTextBox1.SelStart = 0
Form1.RichTextBox1.SelLength = Len(Form1.RichTextBox1.Text)
Form1.RichTextBox1.SelFontName = Font
Form1.RichTextBox1.SelStart = 0
End Function


Function DeleteServer()
Kill App.Path & "\" & List2.List(List2.ListIndex)
File2.Refresh
List2.Refresh
End Function

Private Sub Upload()
Dim Cnt As Long, nFileLen As Long, nRet As Long, nTotFileLen As Long
Dim sBuffer As String * 1024
Dim Ret As Long, SentBytes As Long, sAllBytes As Long, z As Long
Dim I As Integer
Dim Kam As String, Ode As String
Dim FS As Long, StartT As Long, T As Long, p As Long
Dim spRate As Single
Screen.MousePointer = vbHourglass
spRate = 0
sAllBytes = 0
p = 0
For I = 0 To List4.ListCount - 1
Ode = App.Path & "\" & List4.List(I)
Kam = Klic & List4.List(I)
hFile = FtpOpenFile(server, Kam, GENERIC_WRITE, FTP_TRANSFER_TYPE_BINARY, 0)
If hFile = 0 Then
MsgBox "Cant create requested file on server", vbExclamation, App.ProductName
Screen.MousePointer = vbDefault
Exit Sub
End If
SentBytes = 0
nFileLen = 0
StartT = GetTickCount
Open Ode For Binary As #1
nTotFileLen = LOF(1)
Do
Get #1, , sBuffer
If nFileLen < nTotFileLen - sReadBuffer Then
If InternetWriteFile(hFile, sBuffer, sReadBuffer, nRet) = 0 Then
MsgBox "Could not upload file!", vbExclamation, App.ProductName
Exit Do
End If
SentBytes = SentBytes + sReadBuffer
sAllBytes = sAllBytes + sReadBuffer
nFileLen = nFileLen + sReadBuffer
Else
If InternetWriteFile(hFile, sBuffer, nTotFileLen - nFileLen, nRet) = 0 Then
MsgBox "Could not upload file!", vbExclamation, App.Title
Exit Do
End If
SentBytes = SentBytes + (nTotFileLen - nFileLen)
sAllBytes = sAllBytes + (nTotFileLen - nFileLen)
nFileLen = nTotFileLen
End If
If SentBytes <> 0 Then
T = GetTickCount - StartT
If T <> 0 Then
spRate = (spRate + ((SentBytes / 1000) / (T / 1000))) / 2
End If
End If
Loop Until nFileLen >= nTotFileLen
Close
p = T / 1000
InternetCloseHandle hFile
Next I
MsgBox ("Data transfer completed."), vbInformation, App.ProductName
Screen.MousePointer = vbDefault
Saved = True
End Sub

Private Sub SaveFW()
Dim FileNum As Integer, Buffer As String
Dim I As Integer
For I = 0 To List4.ListCount - 1
Buffer = RichTextBox1.Text
FileName = App.Path & "\" & List4.List(I)
FileNum = FreeFile
Open FileName For Output As FileNum
Print #FileNum, Buffer
Close FileNum
FileName = ""
RichTextBox1.Text = ""
RichTextBox1.LoadFile (App.Path & "\" & List4.List(I))
Next I
Saved = True
End Sub

Public Function AutoSave()
Screen.MousePointer = vbHourglass
Dim FileNum As Integer, Buffer As String
Buffer = RichTextBox1.Text
FileName = App.Path & "\" & "Temp.Save"
FileNum = FreeFile
Open FileName For Output As FileNum
Print #FileNum, Buffer
Close FileNum
FileName = ""
Screen.MousePointer = vbDefault
Beep
End Function

Public Function DeleteAutosave()
Dim FileNum As Integer, Buffer As String
Buffer = ""
FileName = App.Path & "\" & "Temp.Save"
FileNum = FreeFile
Open FileName For Output As FileNum
Print #FileNum, Buffer
Close FileNum
FileName = ""
End Function
'''''
'''''
''''
Private Function StartNewBrowser(sURL As String) As Boolean
       
  'start a new instance of the user's browser
  'at the page passed as sURL
   Dim success As Long
   Dim hProcess As Long
   Dim sBrowser As String
   Dim Start As STARTUPINFO
   Dim proc As PROCESS_INFORMATION
   
   sBrowser = GetBrowserName(success)
   
  'did sBrowser get correctly filled?
   If success >= ERROR_FILE_SUCCESS Then
      
     'prepare STARTUPINFO members
      With Start
         .cb = Len(Start)
         .dwFlags = STARTF_USESHOWWINDOW
         .wShowWindow = SW_SHOWNORMAL
      End With
      
     'start a new instance of the default
     'browser at the specified URL. The
     'lpCommandLine member (second parameter)
     'requires a leading space or the call
     'will fail to open the specified page.
      success = CreateProcess(sBrowser, _
                              " " & sURL, _
                              0&, 0&, 0&, _
                              NORMAL_PRIORITY_CLASS, _
                              0&, 0&, Start, proc)
                                  
     'if the process handle is valid, return success
      StartNewBrowser = proc.hProcess <> 0
     
     'don't need the process
     'handle anymore, so close it
      Call CloseHandle(proc.hProcess)

     'and close the handle to the thread created
      Call CloseHandle(proc.hThread)

   End If

End Function

Private Function GetBrowserName(dwFlagReturned As Long) As String

  'find the full path and name of the user's
  'associated browser
   Dim hFile As Long
   Dim sResult As String
   Dim sTempFolder As String
        
  'get the user's temp folder
   sTempFolder = GetTempDir()
   
  'create a dummy html file in the temp dir
   hFile = FreeFile
      Open sTempFolder & "dummy.html" For Output As #hFile
   Close #hFile

  'get the file path & name associated with the file
   sResult = Space$(MAX_PATH)
   dwFlagReturned = FindExecutable("dummy.html", sTempFolder, sResult)
  
  'clean up
   Kill sTempFolder & "dummy.html"
   
  'return result
   GetBrowserName = TrimNull(sResult)
   
End Function


Private Function TrimNull(Item As String)

  'remove string before the terminating null(s)
   Dim pos As Integer
   
   pos = InStr(Item, Chr$(0))
   
   If pos Then
         TrimNull = Left$(Item, pos - 1)
   Else: TrimNull = Item
   End If
   
End Function


Public Function GetTempDir() As String

  'retrieve the user's system temp folder
   Dim tmp As String
   
   tmp = Space$(256)
   Call GetTempPath(Len(tmp), tmp)
   
   GetTempDir = TrimNull(tmp)
    
End Function


