<% Option Explicit %>
<!--#include file="../../../Inc/NosqlHack.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<%
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../Inc/Cls_RefreshJs.asp" -->
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<!--#include file="Function.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P031100") then Call ReturnError1()
Dim FileID,FileIDArray,tpi,RsCrObj,ResultString,j,k,TempStrTip,TemppSql
FileID = Request("FileID")
k=0
if FileID="" then
	Set RsCrObj = Server.CreateObject(G_FS_RS)
	TemppSql = "Select FileName,FileType,FileCName from FS_SysJs"
	RsCrObj.Open TemppSql,Conn,1,1
	tpi = RsCrObj.RecordCount
	do while not RsCrObj.eof
		ResultString = CreateSysJS(RsCrObj("FileName"))
		if ResultString <> true then
			if j="" then
				j = RsCrObj("FileCName")&"&nbsp;"
			else
				j = j &RsCrObj("FileCName") &"&nbsp;"
			end if
			k=k+1
		End If
	RsCrObj.MoveNext
	loop
	RsCrObj.Close
	Set RsCrObj = Nothing
else
	FileIDArray = Split(FileID,"***")
	j = ""
	for tpi = LBound(FileIDArray) to UBound(FileIDArray)
		if FileIDArray(tpi) <> "" then
			Set RsCrObj = Conn.Execute("Select FileName,FileType,FileCName from FS_SysJs where ID="&FileIDArray(tpi)&"")
			if Not RsCrObj.eof then
				ResultString = CreateSysJS(RsCrObj("FileName"))
				if ResultString <> true then
					if j="" then
						j = RsCrObj("FileCName")&"&nbsp;"
					else
						j = j & RsCrObj("FileCName") &"&nbsp;"
					end if
					k=k+1
				End If
				RsCrObj.Close
				Set RsCrObj = Nothing
			end if
		end if
	next
end if
If j = "" then
	TempStrTip = "成功生成所选项目"
Else
	TempStrTip = "共生成项目" & tpi & "项,其中" & K & "项(" & j & ")未能查询到符合条件的新闻！"
End if
Set Conn = Nothing
%>
<html>
<head>
<link href="../../../CSS/ModeWindow.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>生成全站JS</title>
</head>
<body topmargin="0" leftmargin="0">
<table width="94%"  border="0" align="center" cellpadding="0" cellspacing="5">
  <tr>
    <td width="22%" rowspan="3"><div align="center"><img src="../../Images/Info.gif" width="34" height="33"></div></td>
    <td width="78%" height="5"></td>
  </tr>
  <tr>
    <td><%=TempStrTip%> </td>
  </tr>
  <tr>
    <td height="5"></td>
  </tr>
  <tr>
    <td colspan="2"><div align="center">
      <input type="button" name="Submit" value="关闭窗口" onclick="window.close();">
    </div></td>
  </tr>
</table>
</body>
</html>
