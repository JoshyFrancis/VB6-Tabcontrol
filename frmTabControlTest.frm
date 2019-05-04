VERSION 5.00
Begin VB.Form frmTabControlTest 
   Caption         =   "Tab Control Test"
   ClientHeight    =   7485
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12705
   LinkTopic       =   "Form1"
   ScaleHeight     =   7485
   ScaleWidth      =   12705
   StartUpPosition =   3  'Windows Default
   Begin TabControlTest.TabControl TabControl5 
      Height          =   3135
      Left            =   360
      TabIndex        =   11
      Top             =   4080
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   5530
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AllowReorder    =   -1  'True
      SelectedItem    =   15
      ItemCount       =   16
      Item(0).Control(0)=   "TabControl6"
      Item(0).ControlCount=   1
      Item(1).Control(0)=   "TabControl7"
      Item(1).ControlCount=   1
      Item(2).Control(0)=   "TabControl8"
      Item(2).ControlCount=   1
      Item(15).Control(0)=   "Dir1"
      Item(15).ControlCount=   1
      ItemMax         =   15
      Begin VB.DirListBox Dir1 
         Height          =   1440
         Left            =   3480
         TabIndex        =   15
         Top             =   840
         Width           =   2295
      End
      Begin TabControlTest.TabControl TabControl8 
         Height          =   735
         Left            =   -67240
         TabIndex        =   14
         Top             =   1200
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   1296
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin TabControlTest.TabControl TabControl7 
         Height          =   1695
         Left            =   -68200
         TabIndex        =   13
         Top             =   1080
         Visible         =   0   'False
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   2990
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin TabControlTest.TabControl TabControl6 
         Height          =   1455
         Left            =   -68080
         TabIndex        =   12
         Top             =   1320
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   2566
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin TabControlTest.TabControl TabControl2 
      Height          =   3615
      Left            =   5280
      TabIndex        =   8
      Top             =   360
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   6376
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Item(0).Control(0)=   "TabControl3"
      Item(0).ControlCount=   1
      Item(1).Control(0)=   "TabControl4"
      Item(1).ControlCount=   1
      ItemMax         =   1
      Begin TabControlTest.TabControl TabControl4 
         Height          =   1815
         Left            =   -69400
         TabIndex        =   10
         Top             =   840
         Visible         =   0   'False
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   3201
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin TabControlTest.TabControl TabControl3 
         Height          =   1455
         Left            =   960
         TabIndex        =   9
         Top             =   1080
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   2566
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin TabControlTest.TabControl TabControl1 
      Height          =   3615
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   6376
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AllowReorder    =   -1  'True
      TabOrder0       =   1
      Item(0).Control(0)=   "Frame1"
      Item(0).ControlCount=   1
      TabOrder1       =   2
      Item(1).Control(0)=   "Command3"
      Item(1).ControlCount=   1
      TabOrder2       =   3
      TabEnabled2     =   0   'False
      Item(2).Control(0)=   "Check1"
      Item(2).ControlCount=   1
      TabOrder3       =   0
      Item(3).Control(0)=   "Text1"
      Item(3).Control(1)=   "Option2"
      Item(3).Control(2)=   "Option1"
      Item(3).Control(3)=   "Combo1"
      Item(3).ControlCount=   4
      ItemMax         =   3
      Begin VB.OptionButton Option2 
         Caption         =   "Option2"
         Height          =   855
         Left            =   3840
         TabIndex        =   7
         Top             =   1560
         Width           =   495
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   495
         Left            =   1680
         TabIndex        =   6
         Top             =   2760
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   600
         TabIndex        =   5
         Text            =   "Combo1"
         Top             =   720
         Width           =   1935
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   1335
         Left            =   -68680
         TabIndex        =   4
         Top             =   1320
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Command3"
         Height          =   1455
         Left            =   -68920
         TabIndex        =   3
         Top             =   1320
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   1695
         Left            =   -69280
         TabIndex        =   2
         Top             =   1080
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Height          =   855
         Left            =   840
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   1440
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmTabControlTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  
Private Sub Form_Click()
TabControl5.Refresh

End Sub

Private Sub Form_Load()
'TabControl1.LoadTabOrder App.Path & "\TabOrder.txt" 'Not working properly now
'    TabControl1.SelectedItem = 4
'    TabControl1.SetTabEnabled 2, False
'    TabControl1.MoveItem = 1
'    TabControl1.RemoveTab 6
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    TabControl1.SaveTabOrder App.Path & "\TabOrder.txt" 'Not working properly now
End Sub

Private Sub TabControl1_BeforeTabChange(ByVal LastTab As Long, NewTab As Long)
'    If NewTab = 2 Then
'        NewTab = 3
'    End If
End Sub

Private Sub TabControl1_TabClick(ByVal TabIndex As Long)
'    MsgBox TabIndex
End Sub

Private Sub TabControl1_TabOrderChanged(ByVal LastTab As Long, NewTab As Long)
'    If LastTab = 2 Then
'        NewTab = 3
'    End If
        
End Sub
