VERSION 5.00
Begin {17016CEE-E118-11D0-94B8-00A0C91110ED} WCStart 
   ClientHeight    =   7920
   ClientLeft      =   750
   ClientTop       =   1425
   ClientWidth     =   8940
   _ExtentX        =   15769
   _ExtentY        =   13970
   MajorVersion    =   0
   MinorVersion    =   8
   StateManagementType=   1
   ASPFileName     =   ""
   DIID_WebClass   =   "{12CBA1F6-9056-11D1-8544-00A024A55AB0}"
   DIID_WebClassEvents=   "{12CBA1F5-9056-11D1-8544-00A024A55AB0}"
   TypeInfoCookie  =   22
   BeginProperty WebItems {193556CD-4486-11D1-9C70-00C04FB987DF} 
      WebItemCount    =   2
      BeginProperty WebItem1 {FA6A55FE-458A-11D1-9C71-00C04FB987DF} 
         MajorVersion    =   0
         MinorVersion    =   8
         Name            =   "tmp_Home"
         DISPID          =   1280
         Template        =   "Home_Template1.htm"
         Token           =   "WC@"
         DIID_WebItemEvents=   "{4C60625B-BBA1-11D4-B2F8-000102A90980}"
         ParseReplacements=   0   'False
         AppendedParams  =   ""
         HasTempTemplate =   0   'False
         UsesRelativePath=   -1  'True
         OriginalTemplate=   "C:\VbExample\HTML Templates\Home_Template.htm"
         TagPrefixInfo   =   2
         BeginProperty Events {193556D1-4486-11D1-9C70-00C04FB987DF} 
            EventCount      =   0
         EndProperty
         BeginProperty BoundTags {FA6A55FA-458A-11D1-9C71-00C04FB987DF} 
            AttribCount     =   0
         EndProperty
      EndProperty
      BeginProperty WebItem2 {FA6A55FE-458A-11D1-9C71-00C04FB987DF} 
         MajorVersion    =   0
         MinorVersion    =   8
         Name            =   "WC_frmSubmit"
         DISPID          =   1281
         Template        =   ""
         Token           =   "WC@"
         DIID_WebItemEvents=   "{4C606201-BBA1-11D4-B2F8-000102A90980}"
         ParseReplacements=   0   'False
         AppendedParams  =   ""
         HasTempTemplate =   0   'False
         UsesRelativePath=   0   'False
         OriginalTemplate=   ""
         TagPrefixInfo   =   2
         BeginProperty Events {193556D1-4486-11D1-9C70-00C04FB987DF} 
            EventCount      =   0
         EndProperty
         BeginProperty BoundTags {FA6A55FA-458A-11D1-9C71-00C04FB987DF} 
            AttribCount     =   1
            BeginProperty Attrib0 {FA6A55FC-458A-11D1-9C71-00C04FB987DF} 
               TagType         =   1
               Attribute       =   "ACTION"
               State           =   2
               TagName         =   "frmsimExamp"
               OriginalURL     =   ""
               Parent          =   ""
               Template        =   "tmp_Home"
               BoundEvent      =   ""
               BoundItem       =   "WC_frmSubmit"
               Suffix          =   ""
               UsesAnonymousName=   0
               TagNumber       =   0
            EndProperty
         EndProperty
      EndProperty
   EndProperty
   NameInURL       =   "Start"
End
Attribute VB_Name = "WCStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text
Private Sub WebClass_Start()
' When you start your project this code automatically runs by default.

    tmp_Home.WriteTemplate ' Display this Template on screen
End Sub
Private Sub WC_frmSubmit_Respond()
Dim usrName As String
    'When You Press Submit Do This
        usrName = ChangeCase(Request.Form("textfield")) ' Request & Send Value You Entered in The Text Box With This Name To Function ChangeCase
        If usrName = "" Then    ' If User Entered Nothing
            Response.Write ("<HTML>" & Chr(13))
            Response.Write ("You Entered No Value<BR>" & Chr(13))
            Response.Write ("<a href=""" & "Start.ASP""" & ">click here to try again</a><BR>" & Chr(13))
            Response.Write ("</HTML>")
        Else                    ' If You Entered A Value
            Response.Write ("<HTML>" & Chr(13))
            Response.Write ("Value You Entered Was " & usrName)
            Response.Write ("</HTML>")
        End If
End Sub
