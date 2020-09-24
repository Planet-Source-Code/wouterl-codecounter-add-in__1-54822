VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmAddIn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Code counter"
   ClientHeight    =   3195
   ClientLeft      =   2175
   ClientTop       =   1935
   ClientWidth     =   6015
   ControlBox      =   0   'False
   Icon            =   "frmAddIn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6015
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin MSComctlLib.ImageList imlMain 
      Left            =   5100
      Top             =   1860
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddIn.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddIn.frx":035E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddIn.frx":06B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddIn.frx":0A02
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddIn.frx":0D54
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddIn.frx":10A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddIn.frx":13F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddIn.frx":174A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddIn.frx":1A9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddIn.frx":1DEE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView trvSturcture 
      Height          =   2715
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   4789
      _Version        =   393217
      Indentation     =   529
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      Appearance      =   1
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.Frame fraBusy 
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4455
      Begin MSComctlLib.ProgressBar pgbMain 
         Height          =   375
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Please wait while counting lines of code..."
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   660
         Width           =   4455
      End
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblTotal"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   2940
      Width           =   510
   End
End
Attribute VB_Name = "frmAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VBInstance As VBIDE.VBE
Public Connect As Connect

Private Type typMember
    Name As String
    CodeLocation As Long
    Type As Long
End Type

Option Explicit

Private Sub cmdAbout_Click()
    MsgBox App.Title & " by W.O. van der Logt", vbInformation
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    On Error GoTo Error_Handler
    
    Dim lngLines As Long                'Lines in all projects
    Dim lngProjectLines As Long         'Lines in project
    Dim lngMemberLines As Integer       'Lines in member (Method or property)
    
    Dim objVBProject As VBProject       'VB Project
    Dim objVBComponent As VBComponent   'VB Component
    Dim objMember As Member             'Member of the component (Method or property)
    Dim objNode As Node                 'Project
    Dim objSubNode As Node              'Component
    Dim intCounter As Integer           'Counter
    Dim intIcon As Integer              'Icon to add with the node
    
    Dim arrMembers() As typMember       'Temp members array
    
    'Make sure progressbar frame is on top
    fraBusy.ZOrder
    
    'Show dialog. So the user can see the progressbar
    Me.Show
    
    'Set image list
    Set trvSturcture.ImageList = imlMain
    
    'Set progressbar max value
    Dim lngPBCount As Long
    For Each objVBProject In VBInstance.VBProjects
        lngPBCount = lngPBCount + objVBProject.VBComponents.Count
    Next
    pgbMain.Max = lngPBCount
    
    'Loop through all projects
    For Each objVBProject In VBInstance.VBProjects
        'Add project to treeview
        Set objNode = trvSturcture.Nodes.Add(, , objVBProject.Name, objVBProject.Name, 1)
        lngProjectLines = 0
        'Loop trrough all components in project
        For Each objVBComponent In objVBProject.VBComponents
            'Update progressbar
            pgbMain.Value = pgbMain.Value + 1
            DoEvents
            'Only components with a name.. This excludes .RES files
            If objVBComponent.Name <> "" Then
                'Determine icon for the object
                Select Case objVBComponent.Type
                    Case vbext_ct_ActiveXDesigner
                        intIcon = 7
                    Case vbext_ct_ClassModule
                        intIcon = 4
                    Case vbext_ct_DocObject
                        intIcon = 7
                    Case vbext_ct_MSForm
                        intIcon = 2
                    Case vbext_ct_PropPage
                        intIcon = 6
                    Case vbext_ct_RelatedDocument
                        intIcon = 3
                    Case vbext_ct_ResFile
                        intIcon = 3
                    Case vbext_ct_StdModule
                        intIcon = 3
                    Case vbext_ct_VBForm
                        intIcon = 2
                    Case vbext_ct_VBMDIForm
                        intIcon = 8
                    Case vbext_ct_UserControl
                        intIcon = 5
                End Select
                
                'Add the component to project node
                Set objSubNode = trvSturcture.Nodes.Add(objNode.Key, tvwChild, objVBProject.Name & "_" & objVBComponent.Name, objVBComponent.Name & ": " & objVBComponent.CodeModule.CountOfLines & " lines", intIcon)
                'Loop all component codemodule members
                ReDim arrMembers(objVBComponent.CodeModule.Members.Count)
                For intCounter = 1 To objVBComponent.CodeModule.Members.Count
                    'Add all members to array
                    Set objMember = objVBComponent.CodeModule.Members(intCounter)
                    arrMembers(intCounter).Name = objMember.Name
                    arrMembers(intCounter).CodeLocation = objMember.CodeLocation
                    arrMembers(intCounter).Type = objMember.Type
                Next
                    
                'Sort members array on codelocation
                'Based on Philippe Lord's Array-handling/sorting v3 functions
                'http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=24546&lngWId=1
                '====================================================================================
                Dim i          As Long   ' Loop Counter
                Dim j          As Long
                Dim iLBound    As Long
                Dim iUBound    As Long
                Dim iMax       As Long
                Dim iTemp      As Long
                Dim iVal1      As Long
                Dim sVal2      As String
                Dim distance   As Long
                
                iLBound = LBound(arrMembers)
                iUBound = UBound(arrMembers)
                
                iMax = iUBound - iLBound + 1
                
                Do
                    distance = distance * 3 + 1
                Loop Until distance > iMax
                
                Do
                    distance = distance \ 3
                    For i = distance + iLBound To iUBound
                        iTemp = arrMembers(i).CodeLocation
                        iVal1 = arrMembers(i).Type
                        sVal2 = arrMembers(i).Name
                        j = i
                        Do While (arrMembers(j - distance).CodeLocation > iTemp)
                            arrMembers(j).CodeLocation = arrMembers(j - distance).CodeLocation
                            arrMembers(j).Type = arrMembers(j - distance).Type
                            arrMembers(j).Name = arrMembers(j - distance).Name
                            j = j - distance
                            If j - distance < iLBound Then Exit Do
                        Loop
                        arrMembers(j).CodeLocation = iTemp
                        arrMembers(j).Type = iVal1
                        arrMembers(j).Name = sVal2
                    Next i
                Loop Until distance = 1
                '====================================================================================
                
                'Add members to component node
                For intCounter = 1 To UBound(arrMembers)
                    If arrMembers(intCounter).Type = vbext_mt_Method Or arrMembers(intCounter).Type = vbext_mt_Property Then
                        If intCounter = UBound(arrMembers) Then
                            'last member
                            lngMemberLines = (objVBComponent.CodeModule.CountOfLines + 1) - arrMembers(intCounter).CodeLocation
                        Else
                            lngMemberLines = arrMembers(intCounter + 1).CodeLocation - arrMembers(intCounter).CodeLocation
                        End If
                        
                        'Icon
                        Select Case arrMembers(intCounter).Type
                            Case vbext_mt_Method
                                'Method icon
                                intIcon = 9
                            Case vbext_mt_Property
                                'Property icon
                                intIcon = 10
                        End Select
                        
                        trvSturcture.Nodes.Add objSubNode.Key, tvwChild, objVBProject.Name & "_" & objVBComponent.Name & "_" & arrMembers(intCounter).Name, arrMembers(intCounter).Name & ": " & lngMemberLines & " lines", intIcon
                    End If
                Next
                
                'Add the total number of lines in the codemodule to the overall counter
                lngLines = lngLines + objVBComponent.CodeModule.CountOfLines
                'Add the total number of lines in the codemodule to the projectlines counter
                lngProjectLines = lngProjectLines + objVBComponent.CodeModule.CountOfLines
            End If
        Next
        'Update project node with the linecount
        objNode.Text = objNode.Text & ": " & lngProjectLines
    Next
        
    'Show total count over all projects
    lblTotal.Caption = "Total number of lines in all projects: " & lngLines & " lines"
    
Error_Exit:
    'Hide progressbar frame
    fraBusy.Visible = False
    'Enable buttons
    cmdCancel.Enabled = True
    cmdAbout.Enabled = True
    
    Exit Sub
Error_Handler:
    MsgBox "There was an error while counting your code: " & Err.Description, vbExclamation

    Resume Error_Exit
End Sub

