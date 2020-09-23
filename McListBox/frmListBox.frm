VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmListBox 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "McListBox 2.1 - Test Form !!"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8850
   Icon            =   "frmListBox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   407
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   590
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picTop 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   8850
      TabIndex        =   0
      Top             =   0
      Width           =   8850
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "McListBox 2.1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   120
         Width           =   1680
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   240
         Picture         =   "frmListBox.frx":000C
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A perfect replacement for vb's 'ListBox' control, with 'Item HighLight' and 'ListIcons'"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   960
         TabIndex        =   1
         Top             =   480
         Width           =   7095
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4695
      Left            =   3240
      TabIndex        =   3
      Top             =   1200
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   8281
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      Tab             =   3
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmListBox.frx":08D6
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label7"
      Tab(0).Control(1)=   "lbSelCount"
      Tab(0).Control(2)=   "lbSeltext"
      Tab(0).Control(3)=   "lbSelItem"
      Tab(0).Control(4)=   "lbCount"
      Tab(0).Control(5)=   "lstSelect"
      Tab(0).Control(6)=   "Frame2"
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Appearance"
      TabPicture(1)   =   "frmListBox.frx":08F2
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Operations"
      TabPicture(2)   =   "frmListBox.frx":090E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Style"
      TabPicture(3)   =   "frmListBox.frx":092A
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Label11"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label12"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "McListBox2"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "McListBox3"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "Speed"
      TabPicture(4)   =   "frmListBox.frx":0946
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label16"
      Tab(4).Control(1)=   "Label10"
      Tab(4).Control(2)=   "chkHide"
      Tab(4).Control(3)=   "McListBox4"
      Tab(4).Control(4)=   "List1"
      Tab(4).Control(5)=   "cmdTest"
      Tab(4).Control(6)=   "txtCount"
      Tab(4).ControlCount=   7
      Begin VB.TextBox txtCount 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72840
         TabIndex        =   75
         Text            =   "10000"
         Top             =   4080
         Width           =   855
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "Test Speed"
         Height          =   375
         Left            =   -74880
         TabIndex        =   67
         Top             =   4080
         Width           =   2055
      End
      Begin VB.ListBox List1 
         Height          =   3180
         Left            =   -72360
         TabIndex        =   66
         Top             =   840
         Width           =   2535
      End
      Begin VB.Frame Frame2 
         Height          =   2775
         Left            =   -74880
         TabIndex        =   40
         Top             =   1800
         Width           =   5175
         Begin VB.CheckBox chkFlat 
            Caption         =   "Flat ScrollBars"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   78
            Top             =   1920
            Width           =   2295
         End
         Begin VB.CheckBox chkAppear 
            Caption         =   "3D - Appearence"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3120
            TabIndex        =   49
            Top             =   360
            Value           =   1  'Checked
            Width           =   1935
         End
         Begin VB.CheckBox chkBorder 
            Caption         =   "Border"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3120
            TabIndex        =   48
            Top             =   720
            Value           =   1  'Checked
            Width           =   1935
         End
         Begin VB.CheckBox chkIcon 
            Caption         =   "Icon Focus"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1800
            TabIndex        =   47
            Top             =   720
            Value           =   1  'Checked
            Width           =   1815
         End
         Begin VB.CheckBox chkShowicon 
            Caption         =   "Show Icon"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1800
            TabIndex        =   41
            Top             =   360
            Value           =   1  'Checked
            Width           =   1815
         End
         Begin VB.CheckBox chkFull 
            Caption         =   "Full Row Select"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   72
            Top             =   1560
            Width           =   1935
         End
         Begin VB.CheckBox chkStrech 
            Caption         =   "Strech Icon"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   1200
            Width           =   1935
         End
         Begin VB.ComboBox cmbSort 
            Height          =   360
            ItemData        =   "frmListBox.frx":0962
            Left            =   2880
            List            =   "frmListBox.frx":096F
            TabIndex        =   46
            Text            =   "Sort_None"
            Top             =   1440
            Width           =   1935
         End
         Begin VB.ComboBox cmbAlline 
            Height          =   360
            ItemData        =   "frmListBox.frx":099E
            Left            =   2880
            List            =   "frmListBox.frx":09AB
            TabIndex        =   45
            Text            =   "vbLeftJustify"
            Top             =   2280
            Width           =   1935
         End
         Begin VB.TextBox txtRow 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   1560
            TabIndex        =   44
            Text            =   "26"
            Top             =   2280
            Width           =   855
         End
         Begin VB.CheckBox chkFocus 
            Caption         =   "Focus Rectangle"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   43
            Top             =   720
            Value           =   1  'Checked
            Width           =   1815
         End
         Begin VB.CheckBox chkMulti 
            Caption         =   "Multi Select"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   42
            Top             =   360
            Value           =   1  'Checked
            Width           =   1815
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Text Allign"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2880
            TabIndex        =   53
            Top             =   2040
            Width           =   750
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Sort Order"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2880
            TabIndex        =   52
            Top             =   1200
            Width           =   765
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Row Height"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   480
            TabIndex        =   51
            Top             =   2400
            Width           =   825
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Operations"
         Height          =   4215
         Left            =   -74880
         TabIndex        =   27
         Top             =   360
         Width           =   5175
         Begin VB.CommandButton cmdNew 
            Caption         =   "<Create List>"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   76
            Top             =   1320
            Width           =   1815
         End
         Begin VB.CommandButton cmdSelClear 
            BackColor       =   &H00FFD6AC&
            Caption         =   "Clear Selection"
            Height          =   360
            Left            =   240
            TabIndex        =   71
            Top             =   1800
            Width           =   1815
         End
         Begin VB.CommandButton cmdSelectAll 
            BackColor       =   &H00FFD6AC&
            Caption         =   "Select All"
            Height          =   360
            Left            =   240
            TabIndex        =   70
            Top             =   2280
            Width           =   1815
         End
         Begin VB.CheckBox chkItemBold 
            Caption         =   "Item Bold"
            Height          =   255
            Left            =   360
            TabIndex        =   62
            Top             =   3000
            Width           =   1935
         End
         Begin VB.TextBox txtIcon 
            Alignment       =   2  'Center
            Height          =   360
            Left            =   4200
            TabIndex        =   61
            Text            =   "1"
            Top             =   2880
            Width           =   375
         End
         Begin VB.CommandButton cmdIcon 
            BackColor       =   &H00FFD6AC&
            Caption         =   "Set new Icon"
            Height          =   360
            Left            =   2400
            TabIndex        =   60
            Top             =   2880
            Width           =   1575
         End
         Begin VB.CommandButton cmdClear 
            BackColor       =   &H00FFD6AC&
            Caption         =   "Clear All"
            Height          =   360
            Left            =   2400
            TabIndex        =   37
            Top             =   2280
            Width           =   1575
         End
         Begin VB.TextBox txtAdd 
            Height          =   375
            Left            =   240
            TabIndex        =   36
            Text            =   "Test Add"
            Top             =   600
            Width           =   2175
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add Text"
            Height          =   360
            Left            =   2400
            TabIndex        =   35
            Top             =   600
            Width           =   1575
         End
         Begin VB.TextBox txtIndex 
            Alignment       =   2  'Center
            Height          =   360
            Left            =   4200
            TabIndex        =   34
            Text            =   "-1"
            Top             =   600
            Width           =   375
         End
         Begin VB.CommandButton cmdRemove 
            Caption         =   "Remove"
            Height          =   375
            Left            =   2400
            TabIndex        =   33
            Top             =   1800
            Width           =   1575
         End
         Begin VB.TextBox txtSelect 
            Alignment       =   2  'Center
            Height          =   360
            Left            =   4200
            TabIndex        =   32
            Text            =   "30"
            Top             =   1320
            Width           =   375
         End
         Begin VB.CommandButton cmdSelect 
            Caption         =   "Set Selected"
            Height          =   360
            Left            =   2400
            TabIndex        =   31
            Top             =   1320
            Width           =   1575
         End
         Begin VB.TextBox txtRemove 
            Alignment       =   2  'Center
            Height          =   360
            Left            =   4200
            TabIndex        =   30
            Text            =   "-1"
            Top             =   1800
            Width           =   375
         End
         Begin VB.TextBox txtImage 
            Alignment       =   2  'Center
            Height          =   360
            Left            =   4680
            TabIndex        =   29
            Text            =   "0"
            Top             =   600
            Width           =   375
         End
         Begin VB.CheckBox chkBold 
            Caption         =   "Bold"
            Height          =   240
            Left            =   4200
            TabIndex        =   28
            Top             =   960
            Width           =   855
         End
         Begin ListBox.McImageList McImageList1 
            Left            =   360
            Top             =   3480
            _extentx        =   7646
            _extenty        =   873
            imagecount      =   13
            images0         =   "frmListBox.frx":09D8
            images1         =   "frmListBox.frx":0D72
            images2         =   "frmListBox.frx":110C
            images3         =   "frmListBox.frx":14A6
            images4         =   "frmListBox.frx":1840
            images5         =   "frmListBox.frx":1BDA
            images6         =   "frmListBox.frx":1F74
            images7         =   "frmListBox.frx":230E
            images8         =   "frmListBox.frx":26A8
            images9         =   "frmListBox.frx":2A42
            images10        =   "frmListBox.frx":2DDC
            images11        =   "frmListBox.frx":3178
            images12        =   "frmListBox.frx":3514
            currentimage    =   12
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Index"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4080
            TabIndex        =   39
            Top             =   360
            Width           =   420
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Image"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4680
            TabIndex        =   38
            Top             =   360
            Width           =   450
         End
      End
      Begin VB.Frame Frame3 
         Height          =   4215
         Left            =   -74880
         TabIndex        =   5
         Top             =   360
         Width           =   5175
         Begin VB.ComboBox cmbSelection 
            Height          =   360
            ItemData        =   "frmListBox.frx":38B0
            Left            =   2640
            List            =   "frmListBox.frx":38BA
            TabIndex        =   73
            Text            =   "[Style_XP] = 1"
            Top             =   480
            Width           =   1935
         End
         Begin VB.PictureBox picSelGrad 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1800
            ScaleHeight     =   225
            ScaleWidth      =   345
            TabIndex        =   16
            Top             =   2880
            Width           =   375
         End
         Begin VB.PictureBox picSelCol 
            Appearance      =   0  'Flat
            BackColor       =   &H00FBAF4A&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1800
            ScaleHeight     =   225
            ScaleWidth      =   345
            TabIndex        =   15
            Top             =   2640
            Width           =   375
         End
         Begin VB.PictureBox picBack1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1800
            ScaleHeight     =   225
            ScaleWidth      =   345
            TabIndex        =   14
            Top             =   2280
            Width           =   375
         End
         Begin VB.PictureBox picGrid 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1800
            ScaleHeight     =   225
            ScaleWidth      =   345
            TabIndex        =   13
            Top             =   3840
            Width           =   375
         End
         Begin VB.PictureBox picSelFore 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1800
            ScaleHeight     =   225
            ScaleWidth      =   345
            TabIndex        =   12
            Top             =   3600
            Width           =   375
         End
         Begin VB.PictureBox picFore 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1800
            ScaleHeight     =   225
            ScaleWidth      =   345
            TabIndex        =   11
            Top             =   3360
            Width           =   375
         End
         Begin VB.PictureBox picBack 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1800
            ScaleHeight     =   225
            ScaleWidth      =   345
            TabIndex        =   10
            Top             =   2040
            Width           =   375
         End
         Begin VB.PictureBox picCol 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   2130
            Left            =   2640
            Picture         =   "frmListBox.frx":38E2
            ScaleHeight     =   2100
            ScaleWidth      =   2100
            TabIndex        =   9
            Top             =   1920
            Width           =   2130
         End
         Begin VB.CheckBox chkGrid 
            Caption         =   "GridLines"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   8
            Top             =   480
            Width           =   1935
         End
         Begin VB.ComboBox cmbGradient 
            Height          =   360
            ItemData        =   "frmListBox.frx":11ED4
            Left            =   360
            List            =   "frmListBox.frx":11EED
            TabIndex        =   7
            Text            =   "[Fill_None] = 0"
            Top             =   1080
            Width           =   1935
         End
         Begin VB.ComboBox cmbSelGradient 
            Height          =   360
            ItemData        =   "frmListBox.frx":11FA4
            Left            =   2640
            List            =   "frmListBox.frx":11FBD
            TabIndex        =   6
            Text            =   "[Fill_None] = 0"
            Top             =   1080
            Width           =   1935
         End
         Begin VB.OptionButton optselGrad 
            Caption         =   "SelGradCol"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   360
            TabIndex        =   17
            Top             =   2880
            Width           =   2055
         End
         Begin VB.OptionButton optselCol 
            Caption         =   "SelColor"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   360
            TabIndex        =   18
            Top             =   2640
            Width           =   2055
         End
         Begin VB.OptionButton optBackGrad 
            Caption         =   "BackGradCol"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   360
            TabIndex        =   19
            Top             =   2280
            Width           =   2055
         End
         Begin VB.OptionButton optGridCol 
            Caption         =   "GridCol"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   360
            TabIndex        =   20
            Top             =   3840
            Width           =   2055
         End
         Begin VB.OptionButton optSelFore 
            Caption         =   "SelForeColor"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   360
            TabIndex        =   21
            Top             =   3600
            Width           =   2055
         End
         Begin VB.OptionButton optFore 
            Caption         =   "ForeColor"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   360
            TabIndex        =   22
            Top             =   3360
            Width           =   2055
         End
         Begin VB.OptionButton optBackColor 
            Caption         =   "BackColor"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   360
            TabIndex        =   23
            Top             =   2040
            Value           =   -1  'True
            Width           =   2055
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Selection Style"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2640
            TabIndex        =   74
            Top             =   240
            Width           =   1050
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Select color >"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   360
            TabIndex        =   26
            Top             =   1680
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "BackGradient"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   360
            TabIndex        =   25
            Top             =   840
            Width           =   945
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "SelGradient "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2640
            TabIndex        =   24
            Top             =   840
            Width           =   870
         End
      End
      Begin VB.ListBox lstSelect 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         ItemData        =   "frmListBox.frx":12074
         Left            =   -72240
         List            =   "frmListBox.frx":12076
         TabIndex        =   4
         Top             =   720
         Width           =   2535
      End
      Begin ListBox.McListBox McListBox4 
         Height          =   3135
         Left            =   -74880
         TabIndex        =   54
         Top             =   840
         Width           =   2415
         _extentx        =   4260
         _extenty        =   5530
         picture         =   "frmListBox.frx":12078
         font            =   "frmListBox.frx":12096
         selcolor        =   16777215
         selforecolor    =   16576
         fullrowselect   =   -1  'True
         selgradient     =   4
         selgradientcol  =   16494410
         selectionstyle  =   0
      End
      Begin ListBox.McListBox McListBox3 
         Height          =   3615
         Left            =   2760
         TabIndex        =   63
         Top             =   840
         Width           =   2415
         _extentx        =   4260
         _extenty        =   6376
         picture         =   "frmListBox.frx":120BE
         font            =   "frmListBox.frx":120DC
         gridlines       =   -1  'True
         multiselect     =   -1  'True
         showicon        =   -1  'True
         selectionstyle  =   0
         flatscrollbar   =   -1  'True
      End
      Begin ListBox.McListBox McListBox2 
         Height          =   3615
         Left            =   120
         TabIndex        =   77
         Top             =   840
         Width           =   2415
         _extentx        =   4260
         _extenty        =   6376
         picture         =   "frmListBox.frx":12104
         font            =   "frmListBox.frx":12122
         selcolor        =   16777215
         selforecolor    =   0
         fullrowselect   =   -1  'True
         backgradient    =   5
         selgradient     =   4
         backgradientcol =   8421504
         selgradientcol  =   0
         showicon        =   -1  'True
         selectionstyle  =   0
      End
      Begin VB.CheckBox chkHide 
         Caption         =   "Auto Hide ScrollBars"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   79
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Fill 10,000 items!!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -72360
         TabIndex        =   69
         Top             =   600
         Width           =   1275
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Check Immediate Window!!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -71880
         TabIndex        =   68
         Top             =   4200
         Width           =   1980
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Solid !!"
         Height          =   240
         Left            =   120
         TabIndex        =   65
         Top             =   600
         Width           =   600
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Office !! (multi select)"
         Height          =   240
         Left            =   2760
         TabIndex        =   64
         Top             =   600
         Width           =   1860
      End
      Begin VB.Label lbCount 
         AutoSize        =   -1  'True
         Caption         =   "ListCount"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74760
         TabIndex        =   59
         Top             =   540
         Width           =   795
      End
      Begin VB.Label lbSelItem 
         AutoSize        =   -1  'True
         Caption         =   "Selected Item"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74760
         TabIndex        =   58
         Top             =   1020
         Width           =   1200
      End
      Begin VB.Label lbSeltext 
         AutoSize        =   -1  'True
         Caption         =   "Sel Text"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74760
         TabIndex        =   57
         Top             =   780
         Width           =   690
      End
      Begin VB.Label lbSelCount 
         AutoSize        =   -1  'True
         Caption         =   "SelCount"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74760
         TabIndex        =   56
         Top             =   1440
         Width           =   750
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Selected Items"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -72240
         TabIndex        =   55
         Top             =   480
         Width           =   1290
      End
   End
   Begin ListBox.McListBox McListBox1 
      Height          =   4695
      Left            =   120
      TabIndex        =   80
      Top             =   1200
      Width           =   3015
      _extentx        =   5318
      _extenty        =   8281
      picture         =   "frmListBox.frx":1214A
      font            =   "frmListBox.frx":12168
      backgradientcol =   -2147483628
      multiselect     =   -1  'True
      showicon        =   -1  'True
   End
End
Attribute VB_Name = "frmListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkAppear_Click()
    McListBox1.Appearance = chkAppear
End Sub

Private Sub chkBorder_Click()
    McListBox1.BorderStyle = chkBorder
End Sub

Private Sub chkFlat_Click()
    McListBox1.FlatScrollBar = chkFlat
End Sub

Private Sub chkFocus_Click()
    McListBox1.FocusRectangle = chkFocus
End Sub

Private Sub chkFull_Click()
    McListBox1.FullRowSelect = chkFull
End Sub

Private Sub chkGrid_Click()
    McListBox1.GridLines = chkGrid
End Sub

Private Sub chkHide_Click()
    McListBox4.AutoHideScrollBars = chkHide
End Sub

Private Sub chkIcon_Click()
    McListBox1.IconFocus = chkIcon
End Sub

Private Sub chkItemBold_Click()
    McListBox1.ListBold(McListBox1.ListIndex) = chkItemBold.Value
End Sub

Private Sub chkMulti_Click()
    McListBox1.MultiSelect = chkMulti
End Sub

Private Sub chkShowicon_Click()
    McListBox1.ShowIcon = chkShowicon.Value
End Sub

Private Sub chkStrech_Click()
    McListBox1.StrechIcon = chkStrech
End Sub

Private Sub cmbAlline_Click()
    McListBox1.TextAlignment = cmbAlline.ListIndex
End Sub

Private Sub cmbGradient_Click()
    McListBox1.BackGradient = cmbGradient.ListIndex
End Sub

Private Sub cmbSelection_Click()
    McListBox1.SelectionStyle = cmbSelection.ListIndex
End Sub

Private Sub cmbSelGradient_Click()
    McListBox1.SelGradient = cmbSelGradient.ListIndex
End Sub

Private Sub cmbSort_Click()
    If cmbSort.ListIndex = 2 Then
        McListBox1.SortOrder = Sort_Ascending
    Else
        McListBox1.SortOrder = cmbSort.ListIndex
    End If
End Sub

Private Sub cmdAdd_Click()
    McListBox1.AddItem txtAdd, txtIndex, Val(txtImage), chkBold
    McListBox1.Refresh
End Sub

Private Sub cmdBold_Click()

End Sub

Private Sub cmdClear_Click()
    McListBox1.Clear
End Sub

Private Sub cmdIcon_Click()
    McListBox1.ListIcon(McListBox1.ListIndex) = Val(txtIcon)
End Sub

Private Sub cmdNew_Click()
Dim NewList As Control
Dim x As Long

    Set NewList = Controls.Add("ListBox.McListBox", "Test")
    With NewList
        .Visible = True
        .Move 0, 0, ScaleWidth, picTop.Height
        .ZOrder (0)
        .RowHeight = 16
        Set .ImageList = McImageList1
        For x = 0 To 100
            .AddItem "New Item " & x, -1, Rnd(12), Rnd * 1
        Next x
        .Refresh
    End With
    
    MsgBox "McListBox : Dynamically created!", vbInformation, "Created!"
End Sub

Private Sub cmdRemove_Click()
    McListBox1.Remove txtRemove
End Sub

Private Sub cmdSelClear_Click()
    McListBox1.ClearSelection
End Sub

Private Sub cmdSelect_Click()
    McListBox1.ListIndex = txtSelect
End Sub

Private Sub cmdSelectAll_Click()
    McListBox1.SelectAll
End Sub

Private Sub cmdTest_Click()
Dim tStart As Double
Dim x As Long

    List1.Clear
    McListBox4.Clear
    Debug.Print vbCrLf
    

    tStart = Timer
    For x = 0 To Val(txtCount)
        McListBox4.AddItem "ListItem " & x
    Next x
    McListBox4.Refresh
    Debug.Print "McListBox Took " & (Timer - tStart) * 10 & " ms to fill!!"
    
    DoEvents
    tStart = Timer
    For x = 0 To Val(txtCount)
        List1.AddItem "ListItem " & x
    Next x
    Debug.Print "VBListBox Took " & (Timer - tStart) * 10 & " ms to fill!!"


End Sub






Private Sub Form_Load()
Dim x As Long
Dim sText As String
Dim mBold As Boolean

    Randomize
    Set McListBox1.ImageList = McImageList1
    Set McListBox2.ImageList = McImageList1
    Set McListBox3.ImageList = McImageList1
    Set McListBox4.ImageList = McImageList1
    
    For x = 0 To 50
    
        mBold = True
        Select Case x
            Case 5
                sText = "Move Mouse Here...         And now you see the full Text !!"
            Case 6
                sText = "With Re-arrage...          You can see...          More...            More...           More...            More...            More...            More...            More..."
            Case 7
                sText = "Item Completer with...      " & String(33, ChrW(5000)) & " ...Unicode!!"
            Case Else
                sText = "ListItem " & x
                mBold = False
        End Select
            
        McListBox1.AddItem sText, -1, Rnd * 12, mBold
        McListBox2.AddItem sText, -1, Rnd * 12, mBold
        McListBox3.AddItem sText, -1, Rnd * 12, mBold
        
    Next x
    
    McListBox1.Refresh
    McListBox2.Refresh
    McListBox3.Refresh
    McListBox4.Refresh

End Sub

Private Sub McListBox1_SelChange()
Dim x As Long
On Error GoTo Handle

    lbCount = "List Count = " & McListBox1.ListCount
    lbSelItem = "Sel Index = " & McListBox1.ListIndex
    lbSeltext = "Sel Text = " & McListBox1.Text
    lbSelCount = "SelCount = " & McListBox1.SelCount
    
    If McListBox1.ListBold(McListBox1.ListIndex) Then
        chkItemBold.Value = 1
    Else
        chkItemBold.Value = 0
    End If
    
    lstSelect.Clear
    For x = 0 To McListBox1.SelCount - 1
        lstSelect.AddItem McListBox1.List(McListBox1.SelItem(x))
    Next x
    
Handle:
End Sub


Private Sub picCol_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim lnCol As Long
    lnCol = picCol.Point(x, Y)
    If optBackGrad Then McListBox1.BackGradientCol = lnCol: picBack1.BackColor = lnCol
    If optBackColor Then McListBox1.BackColor = lnCol: picBack.BackColor = lnCol
    If optFore Then McListBox1.ForeColor = lnCol: picFore.BackColor = lnCol
    If optSelFore Then McListBox1.SelForeColor = lnCol: picSelFore.BackColor = lnCol

    If optGridCol Then McListBox1.GridColor = lnCol: picGrid.BackColor = lnCol
    
    If optselCol Then McListBox1.SelColor = lnCol: picSelCol.BackColor = lnCol
    If optselGrad Then McListBox1.SelGradientCol = lnCol: picSelGrad.BackColor = lnCol

End Sub

Private Sub txtRow_Change()
    McListBox1.RowHeight = Val(txtRow)
End Sub

Private Sub Command1_Click()
    MsgBox Join(SplitToLines(Me, Text1, 200), vbCrLf)
End Sub

Public Function SplitToLines(hdcObject As Object, ByVal sText As String, _
                    ByVal lLength As Long, Optional ByVal bFilterLines As Boolean = True) As String()

 Dim mArray() As String
 Dim mChar As String
 Dim mLine As String
 Dim lnCount As Long
 Dim xMax As String
 Dim mPos As Long
 Dim x As Long
 Dim lDone As Long

    If bFilterLines Then sText = Replace(sText, vbNewLine, vbNullString)
    xMax = Len(sText)
    
    For x = 1 To xMax
    
        mChar = Mid(sText, x, 1)

        If IsDelim(mChar) Then mPos = x - (lDone + 1)
        If hdcObject.TextWidth(mLine & mChar) >= lLength Or x = xMax Then
            If mPos = 0 Then mPos = x - (lDone + 1)
            ReDim Preserve mArray(lnCount)
            mArray(lnCount) = RTrim(LTrim(Mid(mLine, 1, mPos)))
            mLine = Mid(mLine, mPos + 1, Len(mLine) - mPos)
            lDone = lDone + mPos: mPos = 0
            lnCount = lnCount + 1
        End If
        
        mLine = mLine & mChar
        
    Next x

    mArray(lnCount - 1) = mArray(lnCount - 1) & mChar
    SplitToLines = mArray
    
End Function

Public Function IsDelim(Char As String) As Boolean
    Select Case Asc(Char) ' Upper/Lowercase letters,Underscore Not delimiters
    Case 65 To 90, 95, 97 To 122
        IsDelim = False
    Case Else: IsDelim = True ' Another Character Is delimiter
    End Select
End Function
