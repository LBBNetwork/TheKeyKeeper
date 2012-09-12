VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "The Key Keeper"
   ClientHeight    =   3345
   ClientLeft      =   2835
   ClientTop       =   2745
   ClientWidth     =   5175
   Icon            =   "KeyKeepMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   5175
   Begin VB.ComboBox List2 
      Height          =   315
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   600
      Width           =   4215
   End
   Begin VB.ComboBox List1 
      Height          =   315
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   120
      Width           =   4215
   End
   Begin VB.ListBox List3 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   600
      TabIndex        =   7
      Top             =   1080
      Width           =   4455
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2160
      Width           =   3975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1800
      Width           =   4455
   End
   Begin VB.Label PreRelLabel 
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3000
      Width           =   4215
   End
   Begin VB.Label Label7 
      Height          =   255
      Left            =   3000
      TabIndex        =   9
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Does Windows need to be preinstalled? "
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2640
      Width           =   2895
   End
   Begin VB.Label Label5 
      Caption         =   "BIOS Date:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Pass:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Key:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   615
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   5040
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label2 
      Caption         =   "Build:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Product:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Menu tkkFile 
      Caption         =   "&File"
      Begin VB.Menu makeBootimage 
         Caption         =   "&Make bootable disk image"
      End
      Begin VB.Menu ViewClipBoard 
         Caption         =   "&View clipboard contents"
      End
      Begin VB.Menu bar 
         Caption         =   "-"
      End
      Begin VB.Menu exitcmd 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu aboutmenu 
      Caption         =   "&Help"
      Begin VB.Menu aboutKeyKeeper 
         Caption         =   "&About The Key Keeper"
      End
   End
   Begin VB.Menu PopUpBlocker 
      Caption         =   "PopUpMenu"
      Visible         =   0   'False
      Begin VB.Menu CopyPasta 
         Caption         =   "&Copy to clipboard"
      End
      Begin VB.Menu ViewCBContents 
         Caption         =   "&View clipboard contents"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CurrentKey As String
Private Sub Command1_Click()
frmAbout.Show
End Sub

Private Sub Form_Load()
'On Error GoTo ErrorHandling
'PreRelLabel.Caption = "The Key Keeper | Build " & App.Revision & " | v" & App.Major & "." & App.Minor
List1.AddItem "Chicago (Windows 95)"
List1.AddItem "Memphis (Windows 98)"
List1.AddItem "Memphis NT (Windows 2000)"
List1.AddItem "Me (Windows Me)"
List1.AddItem "Whistler (Windows XP)"
List1.AddItem ".NET Server (Windows Server 2003)"
List1.AddItem "Longhorn (not released)"
List1.AddItem "Longhorn Omega-13 (Early Windows Vista)"
List1.AddItem "Vista (Windows Vista)"
List1.AddItem "7 (Windows 7)"

List1.ListIndex = 0
List2.ListIndex = 0
End Sub

Private Sub viewcbcontents_click()
Clipboardviewer.Show
End Sub
Private Sub CopyPasta_Click()
   Clipboard.Clear
   Clipboard.SetText List3.List(List3.ListIndex)
   
End Sub
Private Sub List3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Button 'Button = vbRightButton Then
Case vbRightButton
 Select Case List3
 Case ""
 
 Case Else
  PopupMenu PopUpBlocker, vbPopupMenuRightButton
  'End If
 End Select
End Select
End Sub

Private Sub makeBootimage_Click()
On Error GoTo ErrorHandling
' Errbox = MsgBox("Place holder.", vbCritical, "Error")
RaWrite = Shell("rawrite -f boot.img -d a", vbNormalFocus)

ErrorHandling:
Select Case Err
Case Is <> 0
 ErrDialog = MsgBox("Could not launch RAWRITE." & vbNewLine & "Check to make sure that RAWRITE.EXE is in the Key Keeper's program directory." & vbNewLine & "If it is not there, re-install The Key Keeper." & vbNewLine & "If it is there, then another error occurred.", vbCritical, "Could not load RAWRITE")
End Select
End Sub
Private Sub exitcmd_Click()
 End
End Sub
Private Sub aboutKeyKeeper_Click()
 frmAbout.Show
End Sub

Private Sub List1_Click()
Select Case List1.ListIndex
Case 0
List2.Clear

 List2.AddItem "56"
 List2.AddItem "58s"
 List2.AddItem "112"
 List2.AddItem "189"
 List2.AddItem "480 (German)"
 List2.ListIndex = 0
Case 1
 List2.Clear

 List2.AddItem "1691"
 List2.AddItem "1900"
 List2.AddItem "2106"
 List2.AddItem "2120"
 List2.AddItem "2183"
 List2.AddItem "2185"
 List2.ListIndex = 0
Case 2
 List2.Clear

 List2.AddItem "2183"
 List2.ListIndex = 0
Case 3
 List2.Clear
 
 List2.AddItem "2332"
 List2.AddItem "2348"
 List2.AddItem "2380"
 List2.AddItem "2394"
 List2.AddItem "2419"
 List2.AddItem "2452"
 List2.AddItem "2460"
 List2.AddItem "2470"
 List2.AddItem "2476"
 List2.AddItem "2481"
 List2.AddItem "2491"
 List2.AddItem "2499"
 List2.AddItem "2513"
 List2.ListIndex = 0
Case 4
 List2.Clear

 List2.AddItem "2410"
 List2.AddItem "2419"
 List2.AddItem "2428"
 List2.AddItem "2430"
 List2.AddItem "2446"
 List2.AddItem "2454"
 List2.AddItem "2457"
 List2.AddItem "2458"
 List2.AddItem "2462"
 List2.AddItem "2465"
 List2.AddItem "2469"
 List2.AddItem "2474"
 List2.AddItem "2475"
 List2.AddItem "2481"
 List2.AddItem "2485"
 List2.AddItem "2486"
 List2.AddItem "2494"
 List2.AddItem "2495"
 List2.AddItem "2502"
 List2.AddItem "2504"
 List2.AddItem "2505"
 List2.AddItem "2517"
 List2.AddItem "2520"
 List2.AddItem "2526"
 List2.AddItem "2531"
 List2.AddItem "2532"
 List2.AddItem "2535"
 List2.AddItem "2542"
 List2.ListIndex = 0
Case 5
 List2.Clear

 List2.AddItem "2410"
 List2.AddItem "2430"
 List2.AddItem "2433"
 List2.AddItem "2455"
 List2.AddItem "2462"
 List2.AddItem "2463"
 List2.AddItem "2464"
 List2.AddItem "2465"
 List2.AddItem "2467"
 List2.AddItem "2493"
 List2.AddItem "3505"
 List2.AddItem "3531"
 List2.AddItem "3541"
 List2.AddItem "3621"
 List2.AddItem "3718"
 List2.ListIndex = 0
Case 6
 List2.Clear

 List2.AddItem "3683"
 List2.AddItem "3706"
 List2.AddItem "3718"
 List2.AddItem "4008"
 List2.AddItem "4011"
 List2.AddItem "4015"
 List2.AddItem "4029"
 List2.AddItem "4033"
 List2.AddItem "4039"
 List2.AddItem "4051"
 List2.AddItem "4053"
 List2.AddItem "4074"
 List2.AddItem "4083 (64-bit only)"
 List2.AddItem "4093"
 List2.ListIndex = 0
Case 7
 List2.Clear
 
 List2.AddItem "5048"
 List2.ListIndex = 0
Case 8
 List2.Clear
 List2.AddItem "5112"
 List2.AddItem "5219"
 List2.AddItem "5231.0 (32-bit)"
 List2.AddItem "5231.2 (64-bit)"
 List2.AddItem "5259"
 List2.AddItem "5270.9"
 List2.AddItem "5308.17"
 List2.AddItem "5308.60"
 List2.AddItem "5342.2"
 List2.AddItem "5365.8"
 List2.AddItem "5381"
 List2.AddItem "5384.4"
 List2.AddItem "5456.5"
 List2.AddItem "5472.5"
 List2.AddItem "5536"
 List2.AddItem "5552"
 List2.AddItem "5600"
 List2.AddItem "5712 (Chinese)"
 List2.AddItem "5728"
 List2.AddItem "5744"
 List2.AddItem "5754"
 List2.AddItem "5840"
 List2.AddItem "6000 (Spanish)"
 
 List2.ListIndex = 0
 
Case 9
 List2.Clear
 
 List2.AddItem "6956"
 List2.AddItem "7000"
 List2.AddItem "7022"
 List2.AddItem "7048"
 List2.AddItem "7057"
 List2.AddItem "7068"
 List2.AddItem "7077"
 List2.AddItem "7106 (Chinese)"
 List2.AddItem "7100"
 List2.AddItem "7127"
 List2.AddItem "7137"
 List2.AddItem "7201"
 List2.AddItem "7227 (SP1 Beta)"
 List2.AddItem "7229"
 List2.AddItem "7231"
 List2.AddItem "7232"
 List2.AddItem "7260"
  
 List2.ListIndex = 0
End Select

End Sub

Private Sub List2_Click()

Select Case List2.ListIndex
Case 0
 Select Case List1.ListIndex
 Case 0
  List3.Clear
  List3.AddItem "11111-11111-11111-11111-11111"
  Text2.Text = "11111-11111-11111-11111-11111"
  Text3.Text = "07-01-1993"
  Label7.Caption = "Yes, Windows 3.1"
  CurrentKey = "11111-11111-11111-11111-11111"
 Case 1
  List3.Clear
  List3.AddItem "11111-11111-11111-11111-11111"
  Text2.Text = "n/a"
  Text3.Text = "03-01-1998"
  Label7.Caption = "No"
  CurrentKey = "11111-11111-11111-11111-11111"
 Case 2
  List3.Clear
  List3.AddItem "11111-11111-11111-11111-11111"
  Text2.Text = "n/a"
  Text3.Text = "11-16-1999"
  Label7.Caption = "No"
  CurrentKey = "11111-11111-11111-11111-11111"
Case 3
 List3.Clear
 List3.AddItem "11111-11111-11111-11111-11111"
 Text2.Text = "n/a"
 Text3.Text = "07-20-1999"
 Label7.Caption = "No"
 CurrentKey = "11111-11111-11111-11111-11111"
Case 4
 List3.Clear
 List3.AddItem "11111-11111-11111-11111-11111"
 Text2.Text = "n/a"
 Text3.Text = "12-09-2000"
 Label7.Caption = "No"
 'CurrentKey = "11111-11111-11111-11111-11111"
Case 5
 List3.Clear
 List3.AddItem "11111-11111-11111-11111-11111"
 Text2.Text = "n/a"
 Text3.Text = "12-09-2000"
 Label7.Caption = "No"
Case 6
 List3.Clear
 List3.AddItem "11111-11111-11111-11111-11111"
 Text2.Text = "n/a"
 Text3.Text = "09-24-2002"
 Label7.Caption = "No"
Case 7
 List3.Clear
 List3.AddItem "11111-11111-11111-11111-11111"
 Text2.Text = "n/a"
 Text3.Text = "04-02-2005"
 Label7.Caption = "No"
Case 8
 List3.Clear
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 Text2.Text = "n/a"
 Text3.Text = "07-21-2005"
 Label7.Caption = "No"
Case 9
 List3.Clear
 List3.AddItem "11111-11111-11111-11111-11111"
 Text2.Text = "n/a"
 Text3.Text = "11-23-08"
 Label7.Caption = "No"
End Select
 
Case 1
 Select Case List1.ListIndex
  Case 0
  List3.Clear
  List3.AddItem "11111-11111-11111-11111-11111"
  Text2.Text = "11111-11111-11111-11111-11111"
  Text3.Text = "11-01-1993"
  Label7.Caption = "Yes, Windows 3.1"
  CurrentKey = "11111-11111-11111-11111-11111"
 Case 1
  List3.Clear
  List3.AddItem "11111-11111-11111-11111-11111"
  Text2.Text = "n/a"
  Text3.Text = "04-28-1998"
  Label7.Caption = "No"
 ' CurrentKey = "11111-11111-11111-11111-11111"
 Case 3
  List3.Clear
  List3.AddItem "11111-11111-11111-11111-11111"
  Text2.Text = "n/a"
  Text3.Text = "08-03-1999"
  Label7.Caption = "No"
  CurrentKey = "11111-11111-11111-11111-11111"
Case 4
  List3.Clear
  List3.AddItem "11111-11111-11111-11111-11111"
  Text2.Text = "n/a"
  Text3.Text = "01-14-2001"
  Label7.Caption = "No"
 ' CurrentKey = "11111-11111-11111-11111-11111"
Case 5
 List3.Clear
 List3.AddItem "11111-11111-11111-11111-11111"
 Text2.Text = "n/a"
 Text3.Text = "01-31-2001"
 Label7.Caption = "No"
Case 6
 List3.Clear
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 Text2.Text = "n/a"
 Text3.Text = "10-30-2002"
 Label7.Caption = "No"
Case 8
 List3.Clear
 List3.AddItem "11111-11111-11111-11111-11111"
 Text2.Text = "n/a"
 Text3.Text = "08-31-2005"
 Label7.Caption = "No"
Case 9
 List3.Clear
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 Text2.Text = "n/a"
 Text3.Text = "12-13-2008"
 Label7.Caption = "No"

End Select
 
Case 2
 Select Case List1.ListIndex
 Case 0
  List3.Clear
  List3.AddItem "11111-11111-11111-11111-11111"
  Text2.Text = "11111-11111-11111-11111-11111"
  Text3.Text = "06-02-1994"
  Label7.Caption = "No"
 ' CurrentKey = "11111-11111-11111-11111-11111"
 Case 1
  List3.Clear
  List3.AddItem "11111-11111-11111-11111-11111"
  Text2.Text = "n/a"
  Text3.Text = "12-12-1998"
  Label7.Caption = "No"
 ' CurrentKey = "11111-11111-11111-11111-11111"
Case 3
  List3.Clear
  List3.AddItem "11111-11111-11111-11111-11111"
  Text2.Text = "n/a"
  Text3.Text = "09-24-1999"
  Label7.Caption = "No"
'  CurrentKey = "11111-11111-11111-11111-11111"
Case 4
  List3.Clear
  List3.AddItem "11111-11111-11111-11111-11111"
  Text2.Text = "n/a"
  Text3.Text = "01-30-2001"
  Label7.Caption = "No"
 ' CurrentKey = "11111-11111-11111-11111-11111"
Case 5
 List3.Clear
 List3.AddItem "11111-11111-11111-11111-11111"
 Text2.Text = "n/a"
 Text3.Text = "02-07-2001"
 Label7.Caption = "No"
Case 6
 List3.Clear
 List3.AddItem "11111-11111-11111-11111-11111"
 Text2.Text = "n/a"
 Text3.Text = "11-20-2002"
 Label7.Caption = "No"
Case 8
 List3.Clear
 List3.AddItem "11111-11111-11111-11111-11111"
 Text2.Text = "n/a"
 Text3.Text = "09-13-2005"
 Label7.Caption = "No"
Case 9
 List3.Clear
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 Text2.Text = "n/a"
 Text3.Text = "01-16-2009"
 Label7.Caption = "No"

End Select
 
Case 3
 Select Case List1.ListIndex
 Case 0
  List3.Clear
  List3.AddItem "101907"
  Text2.Text = "999b70c9e"
  Text3.Text = "09-30-1994"
  Label7.Caption = "No"
 ' CurrentKey = "101907"
 Case 1
  List3.Clear
  List3.AddItem "11111-11111-11111-11111-11111"
  Text2.Text = "n/a"
  Text3.Text = "12-12-1998"
  Label7.Caption = "No"
 ' CurrentKey = "11111-11111-11111-11111-11111"
 Case 3
  List3.Clear
  List3.AddItem "11111-11111-11111-11111-11111"
  Text2.Text = "n/a"
  Text3.Text = "10-15-1999"
  Label7.Caption = "No"
 ' CurrentKey = "11111-11111-11111-11111-11111"
Case 4
 List3.Clear
 List3.AddItem "11111-11111-11111-11111-11111"
 Text2.Text = "n/a"
 Text3.Text = "01-31-2001"
 Label7.Caption = "No"
 CurrentKey = "11111-11111-11111-11111-11111"
Case 5
 List3.Clear
 List3.AddItem "11111-11111-11111-11111-11111"
 Text2.Text = "n/a"
 Text3.Text = "03-08-2001"
 Label7.Caption = "No"
Case 6
 List3.Clear
 List3.AddItem "11111-11111-11111-11111-11111"
 Text2.Text = "n/a"
 Text3.Text = "02-20-2003"
 Label7.Caption = "No"
Case 8
 List3.Clear
 List3.AddItem "11111-11111-11111-11111-11111"
 Text2.Text = "n/a"
 Text3.Text = "10-05-2005"
 Label7.Caption = "No"
Case 9
 List3.Clear
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 Text2.Text = "n/a"
 Text3.Text = "02-20-2009"
 Label7.Caption = "No"

End Select
 
Case 4
Select Case List1.ListIndex
 Case 0
  List3.Clear
  List3.AddItem "11111-11111-11111-11111-11111"
  Text2.Text = "n/a"
  Text3.Text = "06-01-1995"
  Label7.Caption = "Yes, Windows 3.1"
 ' CurrentKey = "11111-11111-11111-11111-11111"
 Case 1
  List3.Clear
  List3.AddItem "11111-11111-11111-11111-11111"
  Text2.Text = "n/a"
  Text3.Text = "04-06-1999"
  Label7.Caption = "No"
Case 3
  List3.Clear
  List3.AddItem "11111-11111-11111-11111-11111"
  Text2.Text = "n/a"
  Text3.Text = "11-20-1999"
  Label7.Caption = "No"
 ' CurrentKey = "11111-11111-11111-11111-11111"
Case 4
  List3.Clear
  List3.AddItem "11111-11111-11111-11111-11111"
  Text2.Text = "n/a"
  Text3.Text = "02-25-2001"
  Label7.Caption = "No"
Case 5
 List3.Clear
 List3.AddItem "11111-11111-11111-11111-11111"
 Text2.Text = "n/a"
 Text3.Text = "03-16-2001"
 Label7.Caption = "No"
Case 6
 List3.Clear
 List3.AddItem "11111-11111-11111-11111-11111"
 Text2.Text = "n/a"
 Text3.Text = "03-05-2003"
 Label7.Caption = "No"
Case 8
 List3.Clear
 List3.AddItem "11111-11111-11111-11111-11111"
 Text2.Text = "n/a"
 Text3.Text = "11-14-2005"
 Label7.Caption = "No"
Case 9
 List3.Clear
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 Text2.Text = "n/a"
 Text3.Text = "03-06-2009"
 Label7.Caption = "No"

End Select
 
Case 5
Select Case List1.ListIndex
 Case 1
  List3.Clear
  List3.AddItem "11111-11111-11111-11111-11111"
  Text2.Text = "n/a"
  Text3.Text = "04-25-1999"
  Label7.Caption = "No"
 ' CurrentKey = "11111-11111-11111-11111-11111"
 Case 3
  List3.Clear
  List3.AddItem "11111-11111-11111-11111-11111"
  Text2.Text = "n/a"
  Text3.Text = "02-05-2000"
  Label7.Caption = "No"
 ' CurrentKey = "11111-11111-11111-11111-11111"
 Case 4
  List3.Clear
  List3.AddItem "11111-11111-11111-11111-11111"
  Text2.Text = "n/a"
  Text3.Text = "03-07-2001"
  Label7.Caption = "No"
Case 5
 List3.Clear
 List3.AddItem "11111-11111-11111-11111-11111"
 Text2.Text = "n/a"
 Text3.Text = "03-29-2001"
 Label7.Caption = "No"
Case 6
 List3.Clear
 List3.AddItem "11111-11111-11111-11111-11111"
 Text2.Text = "n/a"
 Text3.Text = "03-29-2003"
 Label7.Caption = "No"
Case 8
 List3.Clear
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "Key works for both 32 and 64-bit editions of this build"
 Text2.Text = "n/a"
 Text3.Text = "12-15-2005"
 Label7.Caption = "No"
Case 9
 List3.Clear
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 
 Text2.Text = "n/a"
 Text3.Text = "03-22-2009"
 Label7.Caption = "No"

End Select

Case 6
Select Case List1.ListIndex
 Case 3
  List3.Clear
  List3.AddItem "11111-11111-11111-11111-11111"
  Text2.Text = "n/a"
  Text3.Text = "02-05-2000"
  Label7.Caption = "No"
 ' CurrentKey = "11111-11111-11111-11111-11111"
 Case 4
  List3.Clear
  List3.AddItem "11111-11111-11111-11111-11111"
  Text2.Text = "n/a"
  Text3.Text = "03-10-2001"
  Label7.Caption = "No"
 ' CurrentKey = "11111-11111-11111-11111-11111"
Case 5
 List3.Clear
 List3.AddItem "11111-11111-11111-11111-11111"
 Text2.Text = "n/a"
 Text3.Text = "04-06-2001"
 Label7.Caption = "No"
Case 6
 List3.Clear
 List3.AddItem "11111-11111-11111-11111-11111"
 Text2.Text = "n/a"
 Text3.Text = "06-20-2003"
 Label7.Caption = "No"
Case 8
 List3.Clear
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "Key works for both 32 and 64-bit editions of this build"
 Text2.Text = "n/a"
 Text3.Text = "02-18-2006"
 Label7.Caption = "No"
Case 9
 List3.Clear
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "-----------------------------"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 Text2.Text = "n/a"
 Text3.Text = "04-05-2009"
 Label7.Caption = "No"

End Select

Case 7
Select Case List1.ListIndex
 Case 3
  List3.Clear
  List3.AddItem "11111-11111-11111-11111-11111"
  Text2.Text = "n/a"
  Text3.Text = "02-20-2000"
  Label7.Caption = "No"
  'CurrentKey = "11111-11111-11111-11111-11111"
 Case 4
  List3.Clear
  List3.AddItem "11111-11111-11111-11111-11111"
  Text2.Text = "n/a"
  Text3.Text = "03-11-2001"
  Label7.Caption = "No"
 ' CurrentKey = "11111-11111-11111-11111-11111"
Case 5
 List3.Clear
 List3.AddItem "11111-11111-11111-11111-11111"
 Text2.Text = "n/a"
 Text3.Text = "04-11-2001"
 Label7.Caption = "No"
Case 6
 List3.Clear
 List3.AddItem "11111-11111-11111-11111-11111"
 Text2.Text = "n/a"
 Text3.Text = "07-18-2003"
 Label7.Caption = "No"
Case 8
 List3.Clear
 List3.AddItem "11111-11111-11111-11111-11111"
 Text2.Text = "n/a"
 Text3.Text = "02-24-2006"
 Label7.Caption = "No"
Case 9
 List3.Clear
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "-----------------------------"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 Text2.Text = "n/a"
 Text3.Text = "04-09-2009"
 Label7.Caption = "No"

End Select

Case 8
Select Case List1.ListIndex
 Case 3
  List3.Clear
  List3.AddItem "11111-11111-11111-11111-11111"
  Text2.Text = "n/a"
  Text3.Text = "03-01-2000"
  Label7.Caption = "No"
 ' CurrentKey = "11111-11111-11111-11111-11111"
 Case 4
  List3.Clear
  List3.AddItem "11111-11111-11111-11111-11111"
  Text2.Text = "n/a"
  Text3.Text = "03-17-2001"
  Label7.Caption = "No"
'  CurrentKey = "11111-11111-11111-11111-11111"
Case 5
 List3.Clear
 List3.AddItem "11111-11111-11111-11111-11111"
 Text2.Text = "n/a"
 Text3.Text = "04-19-2001"
 Label7.Caption = "No"
Case 6
 List3.Clear
 List3.AddItem "11111-11111-11111-11111-11111"
 Text2.Text = "n/a"
 Text3.Text = "08-28-2003"
 Label7.Caption = "No"
Case 8
 List3.Clear
 List3.AddItem "11111-11111-11111-11111-11111"
 Text2.Text = "n/a"
 Text3.Text = "03-22-2006"
 Label7.Caption = "No"
Case 9
 List3.Clear
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "-----------------------------"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 Text2.Text = "n/a"
 Text3.Text = "04-22-2009"
 Label7.Caption = "No"

End Select

Case 9
Select Case List1.ListIndex
Case 3
  List3.Clear
  List3.AddItem "11111-11111-11111-11111-11111"
  Text2.Text = "n/a"
  Text3.Text = "03-04-2000"
  Label7.Caption = "No"
  'CurrentKey = "11111-11111-11111-11111-11111"
Case 4
  List3.Clear
  List3.AddItem "11111-11111-11111-11111-11111 (Home)"
  List3.AddItem "11111-11111-11111-11111-11111 (Pro)"
  Text2.Text = "n/a"
  Text3.Text = "04-13-2001"
  Label7.Caption = "No"
 ' CurrentKey = "RBDC9-VTRC8-D7972-J97JY-PRVMG"
Case 5
 List3.Clear
 List3.AddItem "DTWB2-VX8WY-FG8R3-X696T-66Y46"
 Text2.Text = "n/a"
 Text3.Text = "06-13-2001"
 Label7.Caption = "No"
Case 6
 List3.Clear
 List3.AddItem "11111-11111-11111-11111-11111 (32-bit)"
 List3.AddItem "11111-11111-11111-11111-11111 (64-bit)"
 Text2.Text = "n/a"
 Text3.Text = "10-02-2003"
 Label7.Caption = "No"
Case 8
 List3.Clear
 List3.AddItem "11111-11111-11111-11111-11111"
 Text2.Text = "n/a"
 Text3.Text = "04-20-2006"
 Label7.Caption = "No"
Case 9
 List3.Clear
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "-----------------------------"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 Text2.Text = "n/a"
 Text3.Text = "05-08-2009"
 Label7.Caption = "No"

End Select

Case 10
Select Case List1.ListIndex
 Case 3
  List3.Clear
  List3.AddItem "11111-11111-11111-11111-11111"
  Text2.Text = "n/a"
  Text3.Text = "04-01-2000"
  Label7.Caption = "No"
 ' CurrentKey = "11111-11111-11111-11111-11111"
 Case 4
  List3.Clear
  List3.AddItem "11111-11111-11111-11111-11111"
  Text2.Text = "n/a"
  Text3.Text = "05-02-2001"
  Label7.Caption = "No"
  'CurrentKey = "11111-11111-11111-11111-11111"
Case 5
 List3.Clear
 List3.AddItem "11111-11111-11111-11111-11111"
 Text2.Text = "n/a"
 Text3.Text = "06-28-2001"
 Label7.Caption = "No"
Case 6
 List3.Clear
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 Text2.Text = "n/a"
 Text3.Text = "10-23-2003"
 Label7.Caption = "No"
Case 8
 List3.Clear
 List3.AddItem "11111-11111-11111-11111-11111"
 Text2.Text = "n/a"
 Text3.Text = "05-02-2006"
 Label7.Caption = "No"
Case 9
 List3.Clear
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 Text2.Text = "n/a"
 Text3.Text = "05-22-2009"
 Label7.Caption = "No"

End Select

Case 11
Select Case List1.ListIndex
 Case 3
  List3.Clear
  List3.AddItem "11111-11111-11111-11111-11111"
  Text2.Text = "n/a"
  Text3.Text = "04-10-2000"
  Label7.Caption = "No"
 'CurrentKey = "11111-11111-11111-11111-11111"
 Case 4
  List3.Clear
  List3.AddItem "11111-11111-11111-11111-11111"
  Text2.Text = "n/a"
  Text3.Text = "05-09-2001"
  Label7.Caption = "No"
  'CurrentKey = "11111-11111-11111-11111-11111"
Case 5
 List3.Clear
 List3.AddItem "11111-11111-11111-11111-11111"
 Text2.Text = "n/a"
 Text3.Text = "07-31-2001"
 Label7.Caption = "No"
Case 6
 List3.Clear
 List3.AddItem "11111-11111-11111-11111-11111"
 Text2.Text = "n/a"
 Text3.Text = "04-25-2004"
 Label7.Caption = "No"
Case 8
 List3.Clear
 List3.AddItem "Same keys for 32 and 64-bit editions of this build"
 List3.AddItem "11111-11111-11111-11111-11111 (Home Basic)"
 List3.AddItem "11111-11111-11111-11111-11111 (Home Premium)"
 List3.AddItem "11111-11111-11111-11111-11111 (Business)"
 List3.AddItem "11111-11111-11111-11111-11111 (Ultimate)"
 Text2.Text = "n/a"
 Text3.Text = "05-19-2006"
 Label7.Caption = "No"
Case 9
 List3.Clear
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 Text2.Text = "n/a"
 Text3.Text = "06-02-2009"
 Label7.Caption = "No"
End Select

Case 12
Select Case List1.ListIndex
 Case 3
  List3.Clear
  List3.AddItem "11111-11111-11111-11111-11111"
  Text2.Text = "n/a"
  Text3.Text = "04-20-2000"
  Label7.Caption = "No"
 ' CurrentKey = "11111-11111-11111-11111-11111"
 Case 4
  List3.Clear
  List3.AddItem "11111-11111-11111-11111-11111"
  Text2.Text = "n/a"
  Text3.Text = "05-15-2001"
  Label7.Caption = "No"
'  CurrentKey = "11111-11111-11111-11111-11111"
Case 5
 List3.Clear
 List3.AddItem "11111-11111-11111-11111-11111"
 Text2.Text = "n/a"
 Text3.Text = "08-11-2001"
 Label7.Caption = "No"
Case 6
 List3.Clear
 List3.AddItem "11111-11111-11111-11111-11111"
 Text2.Text = "n/a"
 Text3.Text = "05-17-2004"
 Label7.Caption = "No"
Case 8
 List3.Clear
 List3.AddItem "Same keys for 32 and 64-bit editions of this build"
 List3.AddItem "11111-11111-11111-11111-11111 (Home Basic)"
 List3.AddItem "11111-11111-11111-11111-11111 (Home Premium)"
 List3.AddItem "11111-11111-11111-11111-11111 (Business)"
 List3.AddItem "11111-11111-11111-11111-11111 (Ultimate)"
 Text2.Text = "n/a"
 Text3.Text = "06-21-2006"
 Label7.Caption = "No"
Case 9
 List3.Clear
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 Text2.Text = "n/a"
 Text3.Text = "06-03-2009"
 Label7.Caption = "No"
End Select

Case 13
Select Case List1.ListIndex
 Case 4
  List3.Clear
  List3.AddItem "11111-11111-11111-11111-11111"
  Text2.Text = "n/a"
  Text3.Text = "05-24-2001"
  Label7.Caption = "No"
  ' CurrentKey = "11111-11111-11111-11111-11111"
Case 5
 List3.Clear
 List3.AddItem "11111-11111-11111-11111-11111"
 Text2.Text = "n/a"
 Text3.Text = "04-03-2002"
 Label7.Caption = "No"
Case 6
 List3.Clear
 List3.AddItem "11111-11111-11111-11111-11111"
 Text2.Text = "n/a"
 Text3.Text = "08-20-2004"
 Label7.Caption = "No"
Case 8
 List3.Clear
 List3.AddItem "11111-11111-11111-11111-11111 (Home Basic)"
 List3.AddItem "11111-11111-11111-11111-11111 (Home Premium)"
 List3.AddItem "11111-11111-11111-11111-11111 (Business)"
 List3.AddItem "11111-11111-11111-11111-11111 (Ultimate)"
 Text2.Text = "n/a"
 Text3.Text = "07-14-2006"
 Label7.Caption = "No"
Case 9
 List3.Clear
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 Text2.Text = "n/a"
 Text3.Text = "06-06-2009"
 Label7.Caption = "No"
End Select

Case 14
Select Case List1.ListIndex
 Case 4
  List3.Clear
  List3.AddItem "11111-11111-11111-11111-11111"
  Text2.Text = "n/a"
  Text3.Text = "06-01-2001"
  Label7.Caption = "No"
 ' CurrentKey = "11111-11111-11111-11111-11111"
Case 5
 List3.Clear
 List3.AddItem "11111-11111-11111-11111-11111"
 Text2.Text = "n/a"
 Text3.Text = "11-15-2002"
 Label7.Caption = "No"
Case 8
 List3.Clear
 List3.AddItem "No keys for Starter or the N editions"
 List3.AddItem "11111-11111-11111-11111-11111 (Home Basic)"
 List3.AddItem "11111-11111-11111-11111-11111 (Home Premium)"
 List3.AddItem "11111-11111-11111-11111-11111 (Business)"
 List3.AddItem "11111-11111-11111-11111-11111 (Ultimate)"
 Text2.Text = "n/a"
 Text3.Text = "07-14-2006"
 Label7.Caption = "No"
Case 9
 List3.Clear
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 Text2.Text = "n/a"
 Text3.Text = "06-09-2009"
 Label7.Caption = "No"
End Select

Case 15
Select Case List1.ListIndex
 Case 4
  List3.Clear
  List3.AddItem "11111-11111-11111-11111-11111"
  Text2.Text = "n/a"
  Text3.Text = "06-03-2001"
  Label7.Caption = "No"
 ' CurrentKey = "11111-11111-11111-11111-11111"
Case 8
 List3.Clear
 List3.AddItem "No keys for Starter or the N editions"
 List3.AddItem "11111-11111-11111-11111-11111 (Home Basic)"
 List3.AddItem "11111-11111-11111-11111-11111 (Home Premium)"
 List3.AddItem "11111-11111-11111-11111-11111 (Business)"
 List3.AddItem "11111-11111-11111-11111-11111 (Ultimate)"
 Text2.Text = "n/a"
 Text3.Text = "08-23-2006"
 Label7.Caption = "No"
Case 9
 List3.Clear
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 Text2.Text = "n/a"
 Text3.Text = "06-11-2009"
 Label7.Caption = "No"
End Select

Case 16
Select Case List1.ListIndex
 Case 4
  List3.Clear
  List3.AddItem "11111-11111-11111-11111-11111"
  Text2.Text = "n/a"
  Text3.Text = "06-14-2001"
  Label7.Caption = "No"
 ' CurrentKey = "11111-11111-11111-11111-11111"
Case 8
 List3.Clear
 List3.AddItem "No keys for Starter or the N editions"
 List3.AddItem "Keys are the same for 32 and 64-bit editions of this build"
 List3.AddItem "11111-11111-11111-11111-11111 (Home Basic)"
 List3.AddItem "11111-11111-11111-11111-11111 (Home Premium)"
 List3.AddItem "11111-11111-11111-11111-11111 (Business)"
 List3.AddItem "11111-11111-11111-11111-11111 (Ultimate)"
 Text2.Text = "n/a"
 Text3.Text = "08-30-2006"
 Label7.Caption = "No"
Case 9
 List3.Clear
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111"
 Text2.Text = "n/a"
 Text3.Text = "06-13-2009"
 Label7.Caption = "No"
End Select

Case 17
Select Case List1.ListIndex
 Case 4
  List3.Clear
  List3.AddItem "11111-11111-11111-11111-11111"
  Text2.Text = "n/a"
  Text3.Text = "06-15-2001"
  Label7.Caption = "No"
 ' CurrentKey = "11111-11111-11111-11111-11111"
Case 8
 List3.Clear
 List3.AddItem "No key for Starter edition"
 List3.AddItem "11111-11111-11111-11111-11111 (Home Basic)"
 List3.AddItem "11111-11111-11111-11111-11111 (Home Premium)"
 List3.AddItem "11111-11111-11111-11111-11111 (Business)"
 List3.AddItem "11111-11111-11111-11111-11111 (Ultimate)"
 Text2.Text = "n/a"
 Text3.Text = "08-30-2006"
 Label7.Caption = "No"
End Select

Case 18
Select Case List1.ListIndex
 Case 4
  List3.Clear
  List3.AddItem "11111-11111-11111-11111-11111"
  Text2.Text = "n/a"
  Text3.Text = "06-23-2001"
  Label7.Caption = "No"
'  CurrentKey = "11111-11111-11111-11111-11111"
Case 8
 List3.Clear
 List3.AddItem "No keys for Starter or the N editions"
 List3.AddItem "Keys are the same for 32 and 64-bit versions of this build"
 List3.AddItem "11111-11111-11111-11111-11111 (Home Basic)"
 List3.AddItem "11111-11111-11111-11111-11111 (Home Premium)"
 List3.AddItem "11111-11111-11111-11111-11111 (Business)"
 List3.AddItem "11111-11111-11111-11111-11111 (Ultimate)"
 Text2.Text = "n/a"
 Text3.Text = "09-18-2006"
 Label7.Caption = "No"
End Select

Case 19
Select Case List1.ListIndex
 Case 4
  List3.Clear
  List3.AddItem "11111-11111-11111-11111-11111"
  Text2.Text = "n/a"
  Text3.Text = "06-26-2001"
  Label7.Caption = "No"
 ' CurrentKey = "11111-11111-11111-11111-11111"
Case 8
 List3.Clear
 List3.AddItem "No keys for Starter or the N editions"
 List3.AddItem "Keys are the same for 32 and 64-bit editions of this build"
 List3.AddItem "11111-11111-11111-11111-11111 (Home Basic)"
 List3.AddItem "11111-11111-11111-11111-11111 (Home Premium)"
 List3.AddItem "11111-11111-11111-11111-11111 (Business)"
 List3.AddItem "11111-11111-11111-11111-11111 (Ultimate)"
 Text2.Text = "n/a"
 Text3.Text = "10-04-2006"
 Label7.Caption = "No"
End Select

Case 20
Select Case List1.ListIndex
 Case 4
  List3.Clear
  List3.AddItem "11111-11111-11111-11111-11111"
  Text2.Text = "n/a"
  Text3.Text = "06-27-2001"
  Label7.Caption = "No"
 ' CurrentKey = "11111-11111-11111-11111-11111"
Case 8
 List3.Clear
 List3.AddItem "No keys for Starter or the N editions"
 'List3.AddItem "Keys are the same for 32 and 64-bit editions of this build"
 List3.AddItem "11111-11111-11111-11111-11111 (Home Basic)"
 List3.AddItem "11111-11111-11111-11111-11111 (Home Premium)"
 List3.AddItem "11111-11111-11111-11111-11111 (Business)"
 List3.AddItem "11111-11111-11111-11111-11111 (Ultimate)"
 Text2.Text = "n/a"
 Text3.Text = "10-07-2006"
 Label7.Caption = "No"
End Select

Case 21
Select Case List1.ListIndex
 Case 4
  List3.Clear
  List3.AddItem "11111-11111-11111-11111-11111"
  Text2.Text = "n/a"
  Text3.Text = "07-14-2001"
  Label7.Caption = "No"
 ' CurrentKey = "11111-11111-11111-11111-11111"
Case 8
 List3.Clear
 'List3.AddItem "No keys for Starter or the N editions"
 'List3.AddItem "Keys are the same for 32 and 64-bit editions of this build"
 List3.AddItem "11111-11111-11111-11111-11111 (Home Basic)"
 List3.AddItem "11111-11111-11111-11111-11111 (Home Premium)"
 List3.AddItem "11111-11111-11111-11111-11111 (Business)"
 List3.AddItem "11111-11111-11111-11111-11111 (Ultimate)"
 Text2.Text = "n/a"
 Text3.Text = "10-19-2006"
 Label7.Caption = "No"
End Select

Case 22
Select Case List1.ListIndex
 Case 4
  List3.Clear
  List3.AddItem "11111-11111-11111-11111-11111"
  Text2.Text = "n/a"
  Text3.Text = "07-18-2001"
  Label7.Caption = "No"
Case 8
 List3.Clear
 'List3.AddItem "No keys for Starter or the N editions"
 'List3.AddItem "Keys are the same for 32 and 64-bit editions of this build"
' List3.AddItem "11111-11111-11111-11111-11111"
' List3.AddItem "711111-11111-11111-11111-11111"
 'List3.AddItem "11111-11111-11111-11111-11111"
 List3.AddItem "11111-11111-11111-11111-11111 (Ultimate)"
 Text2.Text = "n/a"
 Text3.Text = "10-31-2006"
 Label7.Caption = "No"
End Select

Case 23
Select Case List1.ListIndex
 Case 4
  List3.Clear
  List3.AddItem "11111-11111-11111-11111-11111 (Home)"
  List3.AddItem "11111-11111-11111-11111-11111(Pro)"
  Text2.Text = "n/a"
  Text3.Text = "07-25-2001"
  Label7.Caption = "No"
End Select


Case 24
Select Case List1.ListIndex
 Case 4
  List3.Clear
  List3.AddItem "11111-11111-11111-11111-11111"
  Text2.Text = "n/a"
  Text3.Text = "07-31-2001"
  Label7.Caption = "No"
End Select

Case 25
Select Case List1.ListIndex
 Case 4
  List3.Clear
  List3.AddItem "11111-11111-11111-11111-11111"
  Text2.Text = "n/a"
  Text3.Text = "08-01-2001"
  Label7.Caption = "No"
End Select

Case 26
Select Case List1.ListIndex
 Case 4
  List3.Clear
  List3.AddItem "11111-11111-11111-11111-11111"
  Text2.Text = "n/a"
  Text3.Text = "08-04-2001"
  Label7.Caption = "No"
End Select

Case 27
Select Case List1.ListIndex
 Case 4
  List3.Clear
  List3.AddItem "11111-11111-11111-11111-11111"
  Text2.Text = "n/a"
  Text3.Text = "08-12-2001"
  Label7.Caption = "No"
End Select
End Select

' Now do you see why there was never a Key Keeper 2.0?

End Sub

Private Sub ViewClipBoard_Click()
Clipboardviewer.Show
End Sub



