VERSION 5.00
Object = "{2210EC79-A724-4033-AAF4-790E2467C0E8}#1.0#0"; "vbalCmdBar6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Demo membuat menu dengan vbAccelerator CommandBar Control"
   ClientHeight    =   5100
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   ScaleHeight     =   5100
   ScaleWidth      =   7740
   StartUpPosition =   2  'CenterScreen
   Begin vbalCmdBar6.vbalCommandBar cmdBar 
      Align           =   1  'Align Top
      Height          =   375
      Index           =   0
      Left            =   0
      Negotiate       =   -1  'True
      Top             =   0
      Width           =   7740
      _ExtentX        =   13653
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2160
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   "new"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":059A
            Key             =   "open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0B34
            Key             =   "close"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":10CE
            Key             =   "save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1668
            Key             =   "preview"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1C02
            Key             =   "print"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":219C
            Key             =   "mail"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2736
            Key             =   "fax"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2CD0
            Key             =   "powerpoint"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'***************************************************************************
' MMMM  MMMMM  OMMM   MMMO    OMMM    OMMM    OMMMMO     OMMMMO    OMMMMO  '
'  MM    MM   MM MM    MMMO  OMMM    MM MM    MM   MO   OM    MO  OM    MO '
'  MM  MM    MM  MM    MM  OO  MM   MM  MM    MM   MO   OM    MO       OMO '
'  MMMM     MMMMMMMM   MM  MM  MM  MMMMMMMM   MMMMMO     OMMMMO      OMO   '
'  MM  MM        MM    MM      MM       MM    MM   MO   OM    MO   OMO     '
'  MM    MM      MM    MM      MM       MM    MM    MO  OM    MO  OM   MM  '
' MMMM  MMMM    MMMM  MMMM    MMMM     MMMM  MMMM  MMMM  OMMMMO   MMMMMMM  '
'                                                                          '
' K4m4r82's Laboratory                                                     '
' http://coding4ever.wordpress.com
'***************************************************************************

Private Function getIconIndex(ByVal key As String) As Long
    getIconIndex = ImageList1.ListImages.Item(key).index - 1
End Function

Private Sub addMenu(ByVal cmdBar As vbalCommandBar, ByVal objMenu As cCommandBarButtons, ByVal menuName As String, ByVal menuCaption As String, Optional showCaptionInToolbar As Boolean = True)
                    
    Dim btn     As cButton
    
    Set btn = cmdBar.Buttons.Add(menuName, , menuCaption)
    btn.showCaptionInToolbar = showCaptionInToolbar
    objMenu.Add btn
End Sub

Private Sub addMenuItem(ByVal cmdBar As vbalCommandBar, ByVal objMenuItem As cCommandBarButtons, ByVal menuName As String, ByVal menuCaption As String, _
                        Optional iconIndex As Long = -1, Optional buttonStyle As EButtonStyle = eNormal, Optional shortcutKey As KeyCodeConstants, Optional enabled As Boolean = True, Optional selected As Boolean = False)
                        
    Dim btn     As cButton
    
    Dim i       As Long
    Dim n       As Long
    
    If buttonStyle = eSeparator Then
        n = 1
        For i = 1 To cmdBar.Buttons.Count
            If InStr(1, cmdBar.Buttons(i).key, "mnuSpr", vbTextCompare) > 0 Then n = n + 1
        Next i
        
        menuName = "mnuSpr" & n
        menuCaption = ""
    End If
        
    Set btn = cmdBar.Buttons.Add(menuName, iconIndex, menuCaption, buttonStyle, , shortcutKey)
    btn.enabled = enabled
    btn.Checked = selected
    
    objMenuItem.Add btn
End Sub

Public Sub createCommandBars()
    Dim objMenuBar          As cCommandBar
    Dim objMenuBarItem      As cCommandBar
    
    Dim objMenuBarSendTo    As cCommandBar
    
    Dim objMenu             As cCommandBarButtons
    Dim objMenuItem         As cCommandBarButtons
    Dim objMenuItemSendTo   As cCommandBarButtons
    
    Dim menuName            As String
    
    With cmdBar(0)
        '1. MEMBUAT MENU BAR
        Set objMenuBar = .CommandBars.Add("MenuBar") 'MENU BAR UNTUK MENAMPUNG MENU UTAMA. EX : MENU FILE DAN MENU STYLE MENU
        Set objMenu = objMenuBar.Buttons
        
        '2. MEMBUAT MENU/MAIN MENU (EX. FILE, STYLE MENU)
        menuName = "mnuFile"
        Call addMenu(cmdBar(0), objMenu, menuName, "File")
        Set objMenuBarItem = .CommandBars.Add(menuName)
        .Buttons(menuName).Bar = objMenuBarItem 'MENGAITKAN MENU FILE KE MENU BAR
        
        '3. MEMBUAT MENU ITEM/SUB MENU (EX. NEW, OPEN, CLOSE de el el)
        Set objMenuItem = objMenuBarItem.Buttons
        Call addMenuItem(cmdBar(0), objMenuItem, "mnuNew", "New", getIconIndex("new"), , vbKeyN)
        Call addMenuItem(cmdBar(0), objMenuItem, "mnuOpen", "Open", getIconIndex("open"), , vbKeyO)
        Call addMenuItem(cmdBar(0), objMenuItem, "mnuClose", "Close", getIconIndex("close"))
        Call addMenuItem(cmdBar(0), objMenuItem, "", "", , eSeparator)
        Call addMenuItem(cmdBar(0), objMenuItem, "mnuSave", "Save", getIconIndex("save"), , vbKeyS)
        Call addMenuItem(cmdBar(0), objMenuItem, "mnuSaveAs", "Save As...")
        Call addMenuItem(cmdBar(0), objMenuItem, "", "", , eSeparator)
        Call addMenuItem(cmdBar(0), objMenuItem, "mnuPrintPreview", "Print Preview", getIconIndex("preview"))
        Call addMenuItem(cmdBar(0), objMenuItem, "mnuPrint", "Print", getIconIndex("print"), , vbKeyP)
        Call addMenuItem(cmdBar(0), objMenuItem, "", "", , eSeparator)
        
        Call addMenuItem(cmdBar(0), objMenuItem, "mnuSendTo", "Send To")
        '>>>> SUB MENU SEND TO
            Set objMenuBarSendTo = .CommandBars.Add(menuName & ":mnuSendTo")
            .Buttons("mnuSendTo").Bar = objMenuBarSendTo
            
            Set objMenuItemSendTo = objMenuBarSendTo.Buttons
            Call addMenuItem(cmdBar(0), objMenuItemSendTo, "mnuMailRecipient", "Mail Recipient", getIconIndex("mail"))
            Call addMenuItem(cmdBar(0), objMenuItemSendTo, "mnuMailRecipientReview", "Mail Recipient (for Review)")
            Call addMenuItem(cmdBar(0), objMenuItemSendTo, "mnuOnlineMeetingParticipant", "Online Meeting Participant")
            Call addMenuItem(cmdBar(0), objMenuItemSendTo, "mnuFaxRecipient", "Fax Recipient...", getIconIndex("fax"))
            Call addMenuItem(cmdBar(0), objMenuItemSendTo, "", "", , eSeparator)
            Call addMenuItem(cmdBar(0), objMenuItemSendTo, "mnuMicrosoftPowerPoint", "Microsoft PowerPoint", getIconIndex("powerpoint"))
        '<<<<

        Call addMenuItem(cmdBar(0), objMenuItem, "", "", , eSeparator)
        Call addMenuItem(cmdBar(0), objMenuItem, "mnuExit", "Exit", , , vbKeyX)
        
        'MENU : STYLE MENU
        menuName = "mnuStyleMenu"
        Call addMenu(cmdBar(0), objMenu, menuName, "Style Menu")
        Set objMenuBarItem = .CommandBars.Add(menuName)
        .Buttons(menuName).Bar = objMenuBarItem 'MENGAITKAN MENU STYLE MENU KE MENU BAR
        
        '>> SUB MENU STYLE MENU
        Set objMenuItem = objMenuBarItem.Buttons
        Call addMenuItem(cmdBar(0), objMenuItem, "mnuStyle1", "Office XP", , eRadio)
        Call addMenuItem(cmdBar(0), objMenuItem, "mnuStyle2", "Office 2003", , eRadio, , , True)
        Call addMenuItem(cmdBar(0), objMenuItem, "mnuStyle3", "Ms Money", , eRadio)
        Call addMenuItem(cmdBar(0), objMenuItem, "mnuStyle4", "Standar", , eRadio)
        '>>
        
        .MenuImageList = ImageList1
        .Toolbar = .CommandBars("MenuBar")
    End With
End Sub

Private Sub cmdBar_ButtonClick(index As Integer, btn As vbalCmdBar6.cButton)
    Select Case btn.key
        Case "mnuNew": 'TODO : something here
        Case "mnuOpen": 'TODO : something here
        Case "mnuClose": 'TODO : something here
        Case "mnuSave": 'TODO : something here
        Case "mnuSaveAs": 'TODO : something here
        Case "mnuPrintPreview": 'TODO : something here
        Case "mnuPrint": 'TODO : something here
        Case "mnuMailRecipient": 'TODO : something here
        Case "mnuMailRecipientReview": 'TODO : something here
        Case "mnuOnlineMeetingParticipant": 'TODO : something here
        Case "mnuFaxRecipient": 'TODO : something here
        Case "mnuMicrosoftPowerPoint": 'TODO : something here
        Case "mnuExit": End
        
        Case "mnuStyle1": cmdBar(0).Style = eOfficeXP
        Case "mnuStyle2": cmdBar(0).Style = eOffice2003
        Case "mnuStyle3": cmdBar(0).Style = eMoney
        Case "mnuStyle4": cmdBar(0).Style = eComCtl32
    End Select
End Sub

Private Sub cmdBar_RequestNewInstance(index As Integer, ctl As Object)
    Dim lNewIndex As Long

    lNewIndex = cmdBar.UBound + 1
    Load cmdBar(lNewIndex)

    cmdBar(lNewIndex).Align = 0
    Set ctl = cmdBar(lNewIndex)
End Sub

Private Sub Form_Load()
    Call createCommandBars
End Sub
