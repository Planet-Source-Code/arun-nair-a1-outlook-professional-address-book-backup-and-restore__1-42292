VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Outlook Address Book Salvager v 1.1"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7050
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   7050
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtComments 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1845
      TabIndex        =   11
      Tag             =   "_bak"
      Top             =   990
      Width           =   5055
   End
   Begin VB.CommandButton cmdRestore 
      BackColor       =   &H00FF8080&
      Caption         =   "Restore the Selected Contacts Profile"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1530
      Style           =   1  'Graphical
      TabIndex        =   9
      Tag             =   "_restore"
      Top             =   5715
      Width           =   4020
   End
   Begin VB.ComboBox cbUsers 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmMain.frx":0442
      Left            =   1305
      List            =   "frmMain.frx":0444
      TabIndex        =   8
      Tag             =   "_restore"
      Text            =   "Select the user from the list ..."
      Top             =   1125
      Width           =   4605
   End
   Begin MSFlexGridLib.MSFlexGrid wabGridRestore 
      Height          =   3255
      Left            =   90
      TabIndex        =   7
      Tag             =   "_restore"
      Top             =   1530
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   5741
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdSaveProfile 
      BackColor       =   &H00008000&
      Caption         =   "Save Current Contacts Profile"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3735
      Style           =   1  'Graphical
      TabIndex        =   4
      Tag             =   "_bak"
      Top             =   5715
      Width           =   3075
   End
   Begin VB.CommandButton cmdRetrContacts 
      BackColor       =   &H002C2C2C&
      Caption         =   "Retrieve Outlook Contacts"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   135
      TabIndex        =   2
      Tag             =   "_bak"
      Top             =   5715
      Width           =   2895
   End
   Begin MSComctlLib.ProgressBar prgWAB 
      Height          =   285
      Left            =   90
      TabIndex        =   1
      Top             =   5085
      Visible         =   0   'False
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSFlexGridLib.MSFlexGrid wabFlexGrid 
      Height          =   3255
      Left            =   90
      TabIndex        =   0
      Tag             =   "_bak"
      Top             =   1530
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   5741
      _Version        =   393216
      Cols            =   22
      FixedCols       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "Comments on this Contacts Profile"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Tag             =   "_bak"
      Top             =   960
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   555
      Left            =   0
      Picture         =   "frmMain.frx":0446
      Top             =   0
      Width           =   7260
   End
   Begin VB.Label lblPrgStatus 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1080
      TabIndex        =   6
      Top             =   4815
      Width           =   5055
   End
   Begin VB.Label lblUser 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   90
      TabIndex        =   5
      Top             =   675
      Width           =   6900
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      Height          =   780
      Left            =   -45
      TabIndex        =   3
      Top             =   5535
      Width           =   7080
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Menu"
      Begin VB.Menu mnuBackup 
         Caption         =   "&Backup Mode"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuRestore 
         Caption         =   "&Restore Mode"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbt 
         Caption         =   "About Outlook Contacts Salvager"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ** Array which maps the fields to be imported/exported
Dim arrMapFields() As String



Dim cnObj As New adodb.Connection
 Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
                  (ByVal lpBuffer As String, nSize As Long) As Long
    Public Property Get UserName() As Variant
          Dim sBuffer As String
          Dim lSize As Long
          sBuffer = Space$(255)
          lSize = Len(sBuffer)
          Call GetUserName(sBuffer, lSize)
          UserName = Left$(sBuffer, lSize)
     End Property

Public Function qzSizeGrid()
'    Business2TelephoneNumber = rsObj(2)
'                    .BusinessAddress = rsObj(3)
'                    .BusinessAddressCity = rsObj(4)
'                    .BusinessAddressCountry = rsObj(5)
'                    .BusinessAddressPostalCode = rsObj(6)
'                    .BusinessAddressPostOfficeBox = rsObj(7)
'                    .BusinessAddressState = rsObj(8)
'                    .BusinessAddressStreet = rsObj(9)
'                    .BusinessFaxNumber = rsObj(10)
'                    .BusinessHomePage = rsObj(11)
'                    .BusinessTelephoneNumber = rsObj(12)
'                    .Email1DisplayName = rsObj(13)
'                    .Email2Address = rsObj(14)
'                    .Email2DisplayName = rsObj(15)
'                    .Email3Address = rsObj(16)
'                    .Email3DisplayName = rsObj(17)
'                    .MobileTelephoneNumber = rsObj(18)

                    
    With Me.wabFlexGrid
        .ColWidth(0) = 0.55 * wabFlexGrid.Width
        .TextMatrix(0, 0) = "Full Name"
        .ColWidth(1) = 0.438 * wabFlexGrid.Width
        .TextMatrix(0, 1) = "E-Mail Address"
        
    End With
    
    Dim intHiddenHeaders As Integer
    
    For intHiddenHeaders = 2 To wabFlexGrid.Cols - 1
        wabFlexGrid.ColWidth(intHiddenHeaders) = 1000
    Next
    
    With Me.wabGridRestore
        .ColWidth(0) = 0.55 * wabFlexGrid.Width
        .TextMatrix(0, 0) = "Comments"
        .ColWidth(1) = 0.438 * wabFlexGrid.Width
        .TextMatrix(0, 1) = "Date of Backup"
        .ColWidth(2) = 0
        .TextMatrix(0, 2) = "Profile Id"
    End With
    
    
End Function

Private Sub cmdRefresh_Click()
    Me.prgWAB.Value = 1
End Sub



Private Sub cbUsers_Click()
   ' MsgBox TypeName(Me.wabFlexGrid)
    Call clearGrid(Me.wabGridRestore)
    Dim cmdObj As New adodb.Command
    Dim rsObjCnt As New adodb.Recordset
    
    With cmdObj
        .ActiveConnection = cnObj
    End With

    cmdObj.CommandText = "ProfileView"
    cmdObj.CommandType = adCmdStoredProc


            ' *** Append the Parameters for the Stored Procedure call
    cmdObj.Parameters.Append _
    cmdObj.CreateParameter("@username", adVarChar, _
                        adParamInput, 50, Trim(cbUsers.Text))
                        
    Set rsObj = cmdObj.Execute
    
    Dim intRowGrid As Integer
    Do While Not rsObj.EOF
        intRowGrid = intRowGrid + 1
        wabGridRestore.Rows = wabGridRestore.Rows + 1
        wabGridRestore.TextMatrix(intRowGrid, 0) = rsObj(1)
        wabGridRestore.TextMatrix(intRowGrid, 1) = rsObj(2)
        wabGridRestore.TextMatrix(intRowGrid, 2) = rsObj(0)
        'Me.cbUsers.AddItem rsObj(0)
        rsObj.MoveNext
    Loop
    
End Sub

Private Sub cmdRestore_Click()
 '   On Error GoTo chkErr
'    MsgBox Me.wabGridRestore.RowSel
    
    
    Dim cmdObj As New adodb.Command
    Dim rsExec As adodb.Recordset
    Dim rsObj As adodb.Recordset
    Dim rsObj2 As adodb.Recordset
    Dim ol As New Outlook.Application
    Dim ns As Outlook.NameSpace
            
    ' get MAPI reference
    Set ns = ol.GetNamespace("MAPI")

        
     
    Dim paramProfileId As adodb.Parameter
 
 
   
    
    cmdObj.ActiveConnection = cnObj

        
    cmdObj.CommandText = "abView"
    cmdObj.CommandType = adCmdStoredProc


  ' *** Append the Parameters for the Stored Procedure call
            
            
            cmdObj.Parameters.Append _
            cmdObj.CreateParameter("@ProfileId", adInteger, _
                        adParamInput, , _
                        wabGridRestore.TextMatrix(wabGridRestore.RowSel, 2))
                        
            
            Set rsExec = cmdObj.Execute
            


    ' *** Now add these contacts to the address book
        Me.prgWAB.Visible = True
        Me.lblPrgStatus.Caption = "Restoring Outlook Contacts ..."
        Dim intPos As Integer
        Dim intTotalRecs As Integer
        Set rsObj2 = rsExec
        intTotalRecs = rsObj2(0)
        Set rsObj = rsExec.NextRecordset
        
        
        
        Do While Not rsObj.EOF
                
                Dim itmContact As Outlook.ContactItem
                DoEvents
                ' Create new Contact item
                Set itmContact = ol.CreateItem(olContactItem)
            
                ' Setup Contact information...
'                MsgBox "x"
                With itmContact
                    
                    .FullName = Trim(rsObj(0))
                    .Email1Address = Trim(rsObj(1))
                    
                    .Business2TelephoneNumber = Trim(rsObj(2))
                    .BusinessAddress = Trim(rsObj(3))
                    .BusinessAddressCity = Trim(rsObj(4))
                    .BusinessAddressCountry = Trim(rsObj(5))
                    .BusinessAddressPostalCode = Trim(rsObj(6))
                    .BusinessAddressPostOfficeBox = Trim(rsObj(7))
                    .BusinessAddressState = Trim(rsObj(8))
                    .BusinessAddressStreet = Trim(rsObj(9))
                    .BusinessFaxNumber = Trim(rsObj(10))
                    .BusinessHomePage = Trim(rsObj(11))
                    .BusinessTelephoneNumber = Trim(rsObj(12))
                    '.Email1DisplayName = Trim(rsObj(13))
                    .Email2Address = Trim(rsObj(14))
                    '.Email2DisplayName = Trim(rsObj(15))
                    .Email3Address = Trim(rsObj(16))
                    '.Email3DisplayName = Trim(rsObj(17))
                    .MobileTelephoneNumber = Trim(rsObj(18))
                    .Home2TelephoneNumber = Trim(rsObj(19))
                    .HomeTelephoneNumber = Trim(rsObj(20))
                    .JobTitle = Trim(rsObj(21))
                    
                End With
            
                ' Save Contact...
                itmContact.Save
                
                
                intPos = intPos + 1
                prgWAB.Value = intPos / intTotalRecs * 100
                
                Set itmContact = Nothing
                
            rsObj.MoveNext
        Loop
       
        prgWAB.Visible = False
        lblPrgStatus.Caption = ""
        
        MsgBox "All Outlook Contacts in the selected profile has been restored", vbInformation
        
       Set ol = Nothing
       Set ns = Nothing

            
    Exit Sub
    
chkErr:
    
                        
End Sub

Private Sub cmdRetrContacts_Click()
    Call clearGrid(wabFlexGrid)
    Call PopulateContacts
    DoEvents
    Me.cmdSaveProfile.Enabled = True
End Sub
Public Function SaveProfile() As Boolean
 Dim cmdObj As New adodb.Command
 Dim paramProfileId As adodb.Parameter
 
 Dim getCurrentUser As String
 Dim intProfileId As Integer
 
 getCurrentUser = UserName

    
 On Error GoTo abortTransaction
   
    
    cmdObj.ActiveConnection = cnObj
'    MsgBox "start"
  '  cnObj.BeginTrans
        
            cmdObj.CommandText = "AddProfile"
            cmdObj.CommandType = adCmdStoredProc


            ' *** Append the Parameters for the Stored Procedure call
            cmdObj.Parameters.Append _
            cmdObj.CreateParameter("@username", adVarChar, _
                        adParamInput, 50, Left(getCurrentUser, Len(getCurrentUser) - 1))
                        
                        
                
            cmdObj.Parameters.Append _
            cmdObj.CreateParameter("@comment", adVarChar, _
            adParamInput, 250, Trim(txtComments))

            cmdObj.Parameters.Append _
            cmdObj.CreateParameter("@dateOfProfile", adDate, _
            adParamInput, , Now())

            Set paramProfileId = cmdObj.CreateParameter("@ProfileId", adInteger, _
            adParamOutput)
            cmdObj.Parameters.Append paramProfileId
            
            cmdObj.Execute
        
            intProfileId = paramProfileId.Value
            
            Dim intArrParse As Integer
            Dim intArrParse2 As Integer
            Dim strInitQuery As String
            Dim strInitQuery2 As String
            Dim strDelim As String
            
            strDelim = ","
            
            'MsgBox strInitQuery
            
                Me.prgWAB.Visible = True
                Me.lblPrgStatus.Caption = "Back-up of Outlook Contacts in progress ..."
             
'            MsgBox arrMapFields(UBound(arrMapFields, 1) - 1, 0)
            
            For intArrParse = 1 To Me.wabFlexGrid.Rows - 1
            DoEvents
                 strInitQuery = "insert into AddressBookMaster " & _
                    "(ProfileId,"
                
                For intArrParse2 = 0 To UBound(arrMapFields, 2) - 1
                    strInitQuery = strInitQuery & _
                    Trim(Mid(strDelim, 1, intArrParse2)) & _
                    arrMapFields(0, intArrParse2)
                Next
            
                strInitQuery = strInitQuery & ") values (" & intProfileId & ","
 '               MsgBox strInitQuery
                For intArrParse2 = 0 To Me.wabFlexGrid.Cols - 1
'                    MsgBox wabFlexGrid.TextMatrix(intArrParse, intArrParse2)
                    strInitQuery = strInitQuery & _
                    Trim(Mid(strDelim, 1, intArrParse2)) & _
                    "'" & _
                    Trim(charReplacer(Me.wabFlexGrid.TextMatrix(intArrParse, intArrParse2), "'", "''")) _
                    & "'"
                Next
'                MsgBox strInitQuery
                strInitQuery = strInitQuery & ")"
                
                Set cmdObj = Nothing
                Set cmdObj.ActiveConnection = cnObj
                
                cmdObj.CommandText = strInitQuery
                Clipboard.Clear
                Clipboard.SetText (strInitQuery)
                
                cmdObj.Execute
                
                prgWAB.Value = intArrParse / Me.wabFlexGrid.Rows * 100
                
            Next
            
            prgWAB.Visible = False
            Me.lblPrgStatus.Caption = ""
            
        
  '  cnObj.CommitTrans
    SaveProfile = True
Exit Function
    
abortTransaction:
   ' cnObj.RollbackTrans
    SaveProfile = False
    MsgBox Err.Description & wabFlexGrid.TextMatrix(intArrParse, 0) & wabFlexGrid.TextMatrix(intArrParse, 1) & wabFlexGrid.TextMatrix(intArrParse, 2)
    
End Function

Private Sub cmdSaveProfile_Click()
    ' *** If the comments are left blank
    If Trim(txtComments.Text) = "" Then
        MsgBox "You cannot leave the 'Comments' empty. Please fill in the Comments for this contacts profile before back-up", vbExclamation
    Else
    
        If SaveProfile = True Then
            MsgBox "Outlook Contacts back-up successful for the current profile", vbInformation
            populateUserList
        Else
            MsgBox "The Backup operation failed", vbCritical
        End If
        
    End If
End Sub

Private Sub Form_Load()
    Me.MousePointer = vbHourglass
        With cnObj
        .Open "Provider=sqloledb;" & "Data Source=203.199.102.153,1433;" & _
                   "Network Library=DBMSSOCN;" & _
                   "Initial Catalog=bqinAB;" & _
                   "User ID=dbadmin;" & _
                   "Password=inwabpwd;"


        End With
    Me.MousePointer = vbNormal
    
    
    
    Dim getCurrentUser As String
    getCurrentUser = UserName
    Me.lblUser.Caption = "Profile for this user is registered as '" & _
    Left(getCurrentUser, Len(getCurrentUser) - 1) & "'"
    DoEvents
    Call qzSizeGrid
    DoEvents
    BackupOption
    populateUserList
End Sub

   
Public Function PopulateContacts()
   On Error Resume Next
      
      
      Dim intContactsCtr As Integer
      Dim intInitCtr As Integer
      
      Dim ol As New Outlook.Application
      Dim olns As Outlook.NameSpace
      Dim objFolder As Outlook.MAPIFolder
      Dim objAllContacts As Object
      Dim Contact As Object
     
      'Dim objAllContacts As Object
      'Dim Contact As ContactItem
      
'      Contact.Business2TelephoneNumber
'      Contact.BusinessAddress
'      Contact.BusinessAddressCity
'      Contact.BusinessAddressCountry
'      Contact.BusinessAddressPostalCode
'      Contact.BusinessAddressPostOfficeBox
'      Contact.BusinessAddressState
'      Contact.BusinessAddressStreet
'      Contact.BusinessFaxNumber
'      Contact.BusinessHomePage
'      Contact.BusinessTelephoneNumber
'      Contact.Email1DisplayName
'      Contact.Email2Address
'      Contact.Email2DisplayName
'      Contact.Email3Address
'      Contact.Email3DisplayName
'      Contact.MobileTelephoneNumber
       'Contact.Home2TelephoneNumber
       'Contact.HomeTelephoneNumber
      
      ' Set the namespace object
      Set olns = ol.GetNamespace("MAPI")
      
      ' Set the default Contacts folder
      Set objFolder = olns.GetDefaultFolder(olFolderContacts)
      
      ' Set objAllContacts = the collection of all contacts
      Set objAllContacts = objFolder.Items
      intContactsCtr = objAllContacts.Count
      
      ReDim Preserve arrMapFields(1, 22)
      
        arrMapFields(0, 0) = "[Full Name]"
        arrMapFields(0, 1) = "[E-mail Address]"
        arrMapFields(0, 2) = "Business2TelephoneNumber"
        arrMapFields(0, 3) = "BusinessAddress"
        arrMapFields(0, 4) = "BusinessAddressCity"
        arrMapFields(0, 5) = "BusinessAddressCountry"
        arrMapFields(0, 6) = "BusinessAddressPostalCode"
        arrMapFields(0, 7) = "BusinessAddressPostOfficeBox"
        arrMapFields(0, 8) = "BusinessAddressState"
        arrMapFields(0, 9) = "BusinessAddressStreet"
        arrMapFields(0, 10) = "BusinessFaxNumber"
        arrMapFields(0, 11) = "BusinessHomePage"
        arrMapFields(0, 12) = "BusinessTelephoneNumber"
        arrMapFields(0, 13) = "Email1DisplayName"
        arrMapFields(0, 14) = "Email2Address"
        arrMapFields(0, 15) = "Email2DisplayName"
        arrMapFields(0, 16) = "Email3Address"
        arrMapFields(0, 17) = "Email3DisplayName"
        arrMapFields(0, 18) = "MobileTelephoneNumber"
        arrMapFields(0, 19) = "Home2TelephoneNumber"
        arrMapFields(0, 20) = "HomeTelephoneNumber"
        arrMapFields(0, 21) = "JobTitle"

      DoEvents
      
      Me.prgWAB.Visible = True
      For Each Contact In objAllContacts
         Me.lblPrgStatus.Caption = "Reading..."
         DoEvents
         'MsgBox " >> --" & Contact.FullName
         If intInitCtr = 0 Then
         
            Dim intPplGrid As Integer
                'MsgBox Contact.Email1Address
                wabFlexGrid.TextMatrix(1, 0) = Contact.Email1Address
                wabFlexGrid.TextMatrix(1, 1) = Contact.FullName
                wabFlexGrid.TextMatrix(1, 2) = Contact.Business2TelephoneNumber

                wabFlexGrid.TextMatrix(1, 3) = Contact.BusinessAddress
                wabFlexGrid.TextMatrix(1, 4) = Contact.BusinessAddressCity
                wabFlexGrid.TextMatrix(1, 5) = Contact.BusinessAddressCountry
                wabFlexGrid.TextMatrix(1, 6) = Contact.BusinessAddressPostalCode
                wabFlexGrid.TextMatrix(1, 7) = Contact.BusinessAddressPostOfficeBox
                wabFlexGrid.TextMatrix(1, 8) = Contact.BusinessAddressState
                wabFlexGrid.TextMatrix(1, 9) = Contact.BusinessAddressStreet
                wabFlexGrid.TextMatrix(1, 10) = Contact.BusinessFaxNumber
                wabFlexGrid.TextMatrix(1, 11) = Contact.BusinessHomePage
                wabFlexGrid.TextMatrix(1, 12) = Contact.BusinessTelephoneNumber
                wabFlexGrid.TextMatrix(1, 13) = Contact.Email1DisplayName
                wabFlexGrid.TextMatrix(1, 14) = Contact.Email2Address
                wabFlexGrid.TextMatrix(1, 15) = Contact.Email2DisplayName
                wabFlexGrid.TextMatrix(1, 16) = Contact.Email3Address
                wabFlexGrid.TextMatrix(1, 17) = Contact.Email3DisplayName
                wabFlexGrid.TextMatrix(1, 18) = Contact.MobileTelephoneNumber
                wabFlexGrid.TextMatrix(1, 19) = Contact.Home2TelephoneNumber
                wabFlexGrid.TextMatrix(1, 20) = Contact.HomeTelephoneNumber
                wabFlexGrid.TextMatrix(1, 21) = Contact.JobTitle

         Else
         
                wabFlexGrid.TextMatrix(wabFlexGrid.Rows - 1, 0) = Contact.FullName
                wabFlexGrid.TextMatrix(wabFlexGrid.Rows - 1, 1) = Contact.Email1Address
                wabFlexGrid.TextMatrix(wabFlexGrid.Rows - 1, 2) = Contact.Business2TelephoneNumber
                wabFlexGrid.TextMatrix(wabFlexGrid.Rows - 1, 3) = Contact.BusinessAddress
                wabFlexGrid.TextMatrix(wabFlexGrid.Rows - 1, 4) = Contact.BusinessAddressCity
                wabFlexGrid.TextMatrix(wabFlexGrid.Rows - 1, 5) = Contact.BusinessAddressCountry
                wabFlexGrid.TextMatrix(wabFlexGrid.Rows - 1, 6) = Contact.BusinessAddressPostalCode
                wabFlexGrid.TextMatrix(wabFlexGrid.Rows - 1, 7) = Contact.BusinessAddressPostOfficeBox
                wabFlexGrid.TextMatrix(wabFlexGrid.Rows - 1, 8) = Contact.BusinessAddressState
                wabFlexGrid.TextMatrix(wabFlexGrid.Rows - 1, 9) = Contact.BusinessAddressStreet
                wabFlexGrid.TextMatrix(wabFlexGrid.Rows - 1, 10) = Contact.BusinessFaxNumber
                wabFlexGrid.TextMatrix(wabFlexGrid.Rows - 1, 11) = Contact.BusinessHomePage
                wabFlexGrid.TextMatrix(wabFlexGrid.Rows - 1, 12) = Contact.BusinessTelephoneNumber
                wabFlexGrid.TextMatrix(wabFlexGrid.Rows - 1, 13) = Contact.Email1DisplayName
                wabFlexGrid.TextMatrix(wabFlexGrid.Rows - 1, 14) = Contact.Email2Address
                wabFlexGrid.TextMatrix(wabFlexGrid.Rows - 1, 15) = Contact.Email2DisplayName
                wabFlexGrid.TextMatrix(wabFlexGrid.Rows - 1, 16) = Contact.Email3Address
                wabFlexGrid.TextMatrix(wabFlexGrid.Rows - 1, 17) = Contact.Email3DisplayName
                wabFlexGrid.TextMatrix(wabFlexGrid.Rows - 1, 18) = Contact.MobileTelephoneNumber
                wabFlexGrid.TextMatrix(wabFlexGrid.Rows - 1, 19) = Contact.Home2TelephoneNumber
                wabFlexGrid.TextMatrix(wabFlexGrid.Rows - 1, 20) = Contact.HomeTelephoneNumber
                wabFlexGrid.TextMatrix(wabFlexGrid.Rows - 1, 21) = Contact.JobTitle
                
         End If
         
         intInitCtr = intInitCtr + 1
         prgWAB.Value = (intInitCtr / intContactsCtr) * 100
         
         DoEvents
         
         If Err.Number = 0 And intInitCtr < intContactsCtr Then
            wabFlexGrid.Rows = Me.wabFlexGrid.Rows + 1
         End If
         Err.Number = 0

      Next
      Me.lblPrgStatus.Caption = ""
      Me.prgWAB.Visible = False
      
      Set ol = Nothing
      Set olns = Nothing
      Set objFolder = Nothing
      Set objAllContacts = Nothing
      Set Contact = Nothing
      
      

End Function

Function charReplacer(originalString, charToReplace As String, replaceWith As String) As String

Dim tmpStr As String

Dim ptr
Dim qPos As String

ptr = 1
If IsNull(originalString) Then originalString = ""


    Do

        qPos = InStr(1, Mid(originalString, ptr), charToReplace)
    
        If qPos = 0 Then
            
                tmpStr = Trim(tmpStr) & _
                        Mid(originalString, ptr)
                        Exit Do
                          
        Else
            
                tmpStr = Trim(tmpStr) & _
                        Mid(originalString, ptr, qPos - 1) & replaceWith
    
        End If
    
        ptr = ptr + qPos
    
    Loop

charReplacer = tmpStr

End Function

Private Sub mnuAbt_Click()
    frmAbout.Show
End Sub

Private Sub mnuBackup_Click()
    BackupOption
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuRestore_Click()
    RestoreOption
End Sub

Public Function RestoreOption()
    Dim ctlParse As Control
    For Each ctlParse In Me.Controls
        If Trim(ctlParse.Tag) = "_bak" Then
            ctlParse.Visible = False
        ElseIf Trim(ctlParse.Tag) = "_restore" Then
            ctlParse.Visible = True
        End If
    Next
    
    mnuRestore.Checked = True
    mnuBackup.Checked = False
End Function

Public Function BackupOption()
    Dim ctlParse As Control
    For Each ctlParse In Me.Controls
        If Trim(ctlParse.Tag) = "_bak" Then
            ctlParse.Visible = True
        ElseIf Trim(ctlParse.Tag) = "_restore" Then
            ctlParse.Visible = False
        End If
    Next
    
    mnuBackup.Checked = True
    mnuRestore.Checked = False
    
End Function

Public Function populateUserList()
    Dim cmdObj As New adodb.Command
    Dim rsObj As adodb.Recordset
    
    
    
    With cmdObj
        .ActiveConnection = cnObj
    End With

    cmdObj.CommandText = "ProfileView"
    cmdObj.CommandType = adCmdStoredProc


            ' *** Append the Parameters for the Stored Procedure call
    cmdObj.Parameters.Append _
    cmdObj.CreateParameter("@username", adVarChar, _
                        adParamInput, 50, "")
                        
    Set rsObj = cmdObj.Execute
    
    cbUsers.Clear
    cbUsers.Text = "Select the user from the list ..."
    
    Do While Not rsObj.EOF
        Me.cbUsers.AddItem rsObj(0)
        rsObj.MoveNext
    Loop

End Function


Public Function clearGrid(ByRef objFlexGrid As MSFlexGrid)


On Error GoTo chkErr
    Dim intParseFlxG As Integer
    
    'MsgBox qzFlexGrid.Rows
    'Me.qzFlexGrid.AddItem ("test")
    If objFlexGrid.Rows > 1 Then
        
        For intParseFlxG = 1 To objFlexGrid.Rows
            objFlexGrid.RemoveItem 1
        Next
    End If
    
    Exit Function
chkErr:
    If Err.Number = 30015 Then
        ReDim arrGridHeaders(objFlexGrid.Cols) As String
        Dim intParseGridHeader As Integer
        For intParseGridHeader = 0 To objFlexGrid.Cols - 1
            arrGridHeaders(intParseGridHeader) = objFlexGrid.TextMatrix(0, intParseGridHeader)
        Next
        
        objFlexGrid.Clear
        
        
        
        
        For intParseGridHeader = 0 To objFlexGrid.Cols - 1
            objFlexGrid.TextMatrix(0, intParseGridHeader) = arrGridHeaders(intParseGridHeader)
        Next
        
        
    End If
End Function

