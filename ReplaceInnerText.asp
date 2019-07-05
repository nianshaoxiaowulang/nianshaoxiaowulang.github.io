<% Option Explicit %>
<!--#include file="Inc/Const.asp" -->
<!--#include file="Inc/Cls_DB.asp" -->
<!--#include file="Inc/Function.asp" -->
<%
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing

Dim strSQL,HelpID
strSQL = "select ID,HelpContent,HelpSingleContent From [FS_Help]"
Dim rs
set Rs = server.createobject("Adodb.recordset")
Rs.open strSQL,conn,1,3
Dim tData,tBool,tData2
do while not Rs.eof
	tBool = false
	HelpID = Rs("ID")
	tData = Rs("HelpContent")
	tData2 = Rs("HelpSingleContent")
	tData = ReplaceLink(tData)
	tData2 = ReplaceLink(tData2)
	Rs("HelpContent") = tData
	Rs("HelpSingleContent") = tData
	If tBool Then
		Rs.update
	End If
Rs.movenext
loop
Rs.close
Set Rs = Nothing

Conn.close
Set Conn = Nothing

function ReplaceLink(v)
	dim tempstr,re
	tempstr = v
	'Set re=new RegExp
	'	re.IgnoreCase =True
	'	re.Global=True
		'"<(.*)>.*<\/\1>"
		'(http:(\/\/|\\\\)[A-Za-z0-9\./=\?%\-&_~`@[\]:+!]+([^<>""|'])+)
		're.Pattern = "<a (href=[""|'](http:(\/\/|\\\\)[A-Za-z0-9\./=\?%\-&_~`@[\]:+!]+([^<>""|'])+)[""|'])>(.*)<\/a>"
		'tempstr = re.Replace(tempstr,"<a target='_self' $1>$5</a>")
	'set re = nothing
	'tempstr = replace(tempstr,"http://localhost/","../../")
	'tempstr = replace(tempstr,"ShowHelp.asp","search.asp")
	'tempstr = replace(tempstr,"http://192.168.1.11/fooSun/Search.aSp","search.asp")
	'tempstr = replace(tempstr,"search.asp?keywords=","search.asp?keyword=")
	'tempstr = replace(tempstr,"http://192.168.1.11/help/HelpImages/","../../Files/Help/")
	'tempstr = replace(tempstr,"http://192.168.1.11/fs10/Files/Help/Helpimages/","../../Files/Help/")
	'tempstr = replace(tempstr,"../../help/HelpImages/","../../Files/Help/")
	ReplaceLink = tempstr
end function
%>
