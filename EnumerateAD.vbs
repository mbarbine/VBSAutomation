
Option Explicit

On Error Resume Next

Dim strGroupDN
Dim objRootDSE, strDNSDomain, strBase
Set objRootDSE = GetObject("LDAP://RootDSE")
strDNSDomain = objRootDSE.Get("defaultNamingContext")
' Specify the base of the ADSI search.
strBase = "<LDAP://" & strDNSDomain & ">"

' Use ADO to search Active Directory.
Dim objADSICommand, objADSIConnection, objADSIRecordset
Set objADSICommand = CreateObject("ADODB.Command")
Set objADSIConnection = CreateObject("ADODB.Connection")
Set objADSIRecordset = CreateObject("ADODB.Recordset")
objADSIConnection.Provider = "ADsDSOObject"
objADSIConnection.Open "Active Directory Provider"
objADSICommand.ActiveConnection = objADSIConnection

Dim strInput
strInput = InputBox("Please enter the name of the group.", "Group Name")

If strInput = "" Then
	MsgBox "The field was left blank or the function was cancelled!", 48, "Warning:"
  WScript.Quit
End If

strGroupDN = FindObject("group", strInput, "distinguishedName")
MsgBox strGroupDN

Dim dicSeenGroupMember, objExcel
Set dicSeenGroupMember = CreateObject("Scripting.Dictionary")
'Creates Excel Object
Set objExcel = WScript.CreateObject("Excel.Application")
objExcel.Visible = True
objExcel.Workbooks.Add

objExcel.ActiveSheet.Name = strInput
objExcel.ActiveSheet.Range("A1").Activate
'Writes a title for each DL enumerated (optional)
objExcel.ActiveCell.Value = "Name"					'col header
objExcel.ActiveCell.Offset(0,1).Value = "Login ID"			'col header 1
objExcel.ActiveCell.Offset(0,2).Value = "Display Name"			'col header 2
objExcel.ActiveCell.Offset(0,3).Value = "Email Address"			'col header 3
objExcel.ActiveCell.Offset(0,4).Value = "Phone Number"			'col header 4
objExcel.ActiveCell.Offset(0,5).Value = "Company"			'col header 5
objExcel.ActiveCell.Offset(0,6).Value = "Office"	       		'col header 6
objExcel.ActiveCell.Offset(0,7).Value = "department"	       		'col header 7
objExcel.ActiveCell.Offset(0,8).Value = "title"	         	 	'col header 8
objExcel.ActiveCell.Offset(0,9).Value = "manager"	       		'col header 9
objExcel.ActiveCell.Offset(0,10).Value = "Description"	      		'col header 10
objExcel.ActiveCell.Offset(0,11).Value = "postalAddress"       		'col header 11
objExcel.ActiveCell.Offset(0,12).Value = "Street"	       		'col header 12
objExcel.ActiveCell.Offset(0,13).Value = "City"	        		'col header 13
objExcel.ActiveCell.Offset(0,14).Value = "State" 	       		'col header 14
objExcel.ActiveCell.Offset(0,15).Value = "postalCode"	       		'col header 15
objExcel.ActiveCell.Offset(0,16).Value = "scriptPath"	       		'col header 16
objExcel.ActiveCell.Offset(0,17).Value = "facsimileTelephoneNumber"    	'col header 17
objExcel.ActiveCell.Offset(0,18).Value = "employeeID"	       		'col header 18
objExcel.ActiveCell.Offset(0,19).Value = "mobile"	       		'col header 10
objExcel.ActiveCell.Offset(0,20).Value = "memberOf"	       		'col header 20



objExcel.ActiveCell.Offset(1,0).Activate				'move 1 down

DisplayMembers "LDAP://" & strGroupDN, " ", dicSeenGroupMember

WScript.Echo vbCrLf & "The script is complete."
WScript.Quit

Function DisplayMembers (strGroupADsPath, strSpaces, dicSeenGroupMember)
	 On Error Resume Next
   Dim objGroup, objMember
   Set objGroup = GetObject(strGroupADsPath)
   For Each objMember In objGroup.Members

      If objMember.Class = "group" Then

         If dicSeenGroupMember.Exists(objMember.ADsPath) Then
            Wscript.Echo strSpaces & "   ^ already seen group member " & _
                                     "(stopping to avoid loop)"
         Else
            dicSeenGroupMember.Add objMember.ADsPath, 1
            DisplayMembers objMember.ADsPath, strSpaces & "  ", _
                           dicSeenGroupMember
         End If
         
      Else
      
      	objExcel.ActiveCell.Value = objMember.Get("name")
				objExcel.ActiveCell.Offset(0,1).Value = objMember.Get("sAMAccountName")
				objExcel.ActiveCell.Offset(0,2).Value = objMember.Get("displayName")
				objExcel.ActiveCell.Offset(0,3).Value = objMember.Get("mail")
				objExcel.ActiveCell.Offset(0,4).Value = objMember.Get("telephoneNumber")
				objExcel.ActiveCell.Offset(0,5).Value = objMember.Get("company")
                                objExcel.ActiveCell.Offset(0,6).Value = objMember.Get("physicalDeliveryOfficeName")
				objExcel.ActiveCell.Offset(0,7).Value = objMember.Get("department")
                                objExcel.ActiveCell.Offset(0,8).Value = objMember.Get("title")
                                objExcel.ActiveCell.Offset(0,9).Value = objMember.Get("manager")
				objExcel.ActiveCell.Offset(1,0).Activate			'move 1 down

      End If

   Next
End Function

Function FindObject(strClass, strObjectName, strAttrib)
	Dim strFilter, strQuery, strValue
	strFilter = "(&(objectClass=" & strClass & ")(name=" & strObjectName & "))"
			
	' Specify the query. "subtree" means to search all child containers/OU's.
	strQuery = strBase & ";" & strFilter & ";" & strAttrib & ";subtree"
		
	objADSICommand.CommandText = strQuery
	Set objADSIRecordSet = objADSICommand.Execute
	
	Do Until objADSIRecordSet.EOF
		strValue = objADSIRecordSet.Fields(strAttrib)
		objADSIRecordSet.MoveNext
	Loop
	
	Set objADSIRecordSet = Nothing
	
	FindObject = strValue
End Function
