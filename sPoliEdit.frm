VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fPoliEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   4785
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   180
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sPoliEdit.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sPoliEdit.frx":015C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sPoliEdit.frx":02B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sPoliEdit.frx":2A6C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   7230
      Width           =   1515
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1380
      TabIndex        =   1
      Top             =   7230
      Width           =   1515
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   6975
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   12303
      _Version        =   393217
      Indentation     =   353
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
End
Attribute VB_Name = "fPoliEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bDirtyBit As Boolean
Dim asRegName(1 To 36) As String
Dim nExplorer As Long, nNetwork As Long, nSystem As Long, nDos As Long

Private Sub Command1_Click()
   Dim n As Node
   For Each n In TreeView1.Nodes
       If n.Image = 1 Then Call DisableCPL(n.Parent.Text, n.Key)
       If n.Image = 2 Then Call EnableCPL(n.Parent.Text, n.Key)
   Next
   Command1.Enabled = False
   If SHRestartSystem(hWnd, "(Note: Some changes don't require reboot)" & vbCrLf & vbCrLf & Chr$(0), Restart_Reboot) = vbYes Then
      Unload Me
      End
   End If
End Sub

Private Sub Command2_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Dim Root As Node
  Dim s1 As String, s2 As String, n As Long
  InitRegArray
  With TreeView1.Nodes
     For i = 1 To UBound(asRegName)
         If i = nExplorer Then Set Root = .Add(, , , "Explorer", 3)
         If i = nNetwork Then Set Root = .Add(, , , "Network", 3)
         If i = nSystem Then Set Root = .Add(, , , "System", 3)
         If i = nDos Then Set Root = .Add(, , , "Dos", 3)
         Root.ExpandedImage = 4
         n = InStr(1, asRegName(i), ",")
         s1 = Left$(asRegName(i), n - 1)
         s2 = Mid$(asRegName(i), n + 1)
         .Add Root, tvwChild, s1, s2, Abs(IsCPLEnable(Root.Text, s1)) + 1
       Next i
    End With
    Command1.Caption = "&Apply"
    Command1.Enabled = False
    Command2.Caption = "&Close"
    Caption = "Program by Registry Editor"
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
  Command1.Enabled = True
  bDirtyBit = True
  If Node.Image = 1 Then Node.Image = 2 Else Node.Image = 1
End Sub

Private Sub InitRegArray()
  nExplorer = 1
' Menu staff
  asRegName(1) = "NoRun,Hide Run Menu"
  asRegName(2) = "NoFind,Hide Find Menu"
  asRegName(3) = "NoSetFolders,Hide Folders in Settings Menu"
  asRegName(4) = "NoSetTaskbar,Hide Taskbar in Settings Menu"
  asRegName(5) = "NoRecentDocsHistory,Hide Resent Documents Menu"
  asRegName(6) = "ClearRecentDocsHistory,Clear Resent Documents Menu"
'desktop staff
  asRegName(7) = "NoDeletePrinter,Disable printer deletion"
  asRegName(8) = "NoAddPrinter,Disable printer addition"
  asRegName(9) = "NoClose,Remove Shutdown Command"
  asRegName(10) = "NoSaveSettings,Don't save Settings"
  asRegName(11) = "NoDesktop,Hide all Items on the Desktop"
  asRegName(12) = "NoDrives,Hide Drives from  Desktop"
  asRegName(13) = "NoNetHood,Hide Network icon from Desktop"
  asRegName(14) = "NoInternetIcon,Hide Internet icon from Desktop"
  asRegName(15) = "NoPrinterTabs,Disable Details and General Pages"
'Network
  nNetwork = 16
  asRegName(16) = "NoNetSetup,Disable Network Control Panel"
  asRegName(17) = "NoNetSetupIDPage,Disable Identification Page"
  asRegName(18) = "NoNetSetupSecurityPage,Disable Access Control Page"
  asRegName(19) = "NoFileSharingControl,Disable File and Print Sharing Controls"
  asRegName(20) = "NoFileSharing,Disable File Sharing"
  asRegName(21) = "NoPrintSharing,Disable Print Sharing"
'System
  nSystem = 22
  asRegName(22) = "NoDispCPL,Disable Display System Control Panel"
  asRegName(23) = "NoDispBackgroundPage,Disable Display Control Background Page"
  asRegName(24) = "NoDispScrSavPage,Disable Display Control Screen Saver Page"
  asRegName(25) = "NoDispApperancePage,Disable Display Control Apperance Page"
  asRegName(26) = "NoDispSettingPage,Disable Display Control Setting Page"
  asRegName(27) = "NoConfigPage,Disable Configuration Page"
  asRegName(28) = "NoDevMgrPage,Disable Device Manager Page"
  asRegName(29) = "NoFileSysPage,Disable File System Page"
  asRegName(30) = "NoVirtMemPage,Disable Virtual Memory Page"
  asRegName(31) = "NoAdminPage,Disable Remote Administration Page"
  asRegName(32) = "NoProfilePage,Disable User Profiles Page"
  asRegName(33) = "NoPwdPage,Disable Change Passwords Page"
  asRegName(34) = "NoSecCPL,Disable Password Control Panel"
' Dos
  nDos = 35
  asRegName(35) = "Disabled,Disable MS-DOS prompt"
  asRegName(36) = "NoRealMode,Disable MS-DOS app at real mode"
End Sub
