<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<%
  dim CodeStr,Condition,JSName,RsTemp,conn,DBC,JSTable,JsID,SQLStr,CodeConfigObj
	Set DBC = New DataBaseClass
	Set Conn = DBC.OpenConnection()
	Set DBC = Nothing
%>
<!--#include file="../../../Inc/Session.asp" -->

<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P060700") then Call ReturnError()
  	Condition=request("Condition")
	JSName=request("JSName")
	JsTable=request("JsTable")
	JsID=request("JsID")
	Set CodeConfigObj = Conn.Execute("Select DoMain from FS_Config")
	if JSName="Location" then
		CodeStr=CodeConfigObj("DoMain")&"/JS/AdsJS/"&JsID&".js"
	ElseIf InStr("Location,Ename,FileName",Trim(JSName))>0 And InStr("FS_Ads,FS_FreeJS,FS_SysJS",Trim(JsTable))>0 Then
		SQLStr="select "& JSName &" from "& JsTable &" where ID= "& Cint(JsID) &""
			Set RsTemp = Conn.Execute(SQLStr)
		CodeStr=RsTemp(JsName)&".js"
		select case JSName
			case "Ename"  CodeStr=CodeConfigObj("DoMain")&"/JS/FreeJS/"&CodeStr
			case "FileName" 
				SQLStr="select FileSavePath from "& JsTable &" where ID = "& Cint(JsID)
				set RsTemp=Conn.Execute(SQLStr)
				CodeStr=CodeConfigObj("DoMain")&RsTemp("FileSavePath") & "/" &CodeStr
			case "Location" CodeStr=CodeConfigObj("DoMain")&"/JS/AdsJS/"&CodeStr
		end select
	end if
	
	  CodeStr=server.HTMLEncode("<script src="&CodeStr&"></script>")
%><head>
<link href="../../../CSS/ModeWindow.css" rel="stylesheet" type="text/css">
</head>

<title>代码调用</title>
<body topmargin="0" leftmargin="0">
<table width="75%" border="0" align="center" cellpadding="0" cellspacing="1">
  <tr> 
    <td width="21%" rowspan="3"><div align="center"><img src="../../Images/Info.gif" width="34" height="33"></div></td>
    <td width="79%" height="15">&nbsp;</td>
  </tr>
  <tr> 
    <td>该JS调用代码为:</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="2"> <div align="center"> 
        <textarea name="textfield" cols="60" rows="4"><%=CodeStr%></textarea>
      </div></td>
  </tr>
  <tr> 
    <td colspan="2"> <div align="center"> 
        <input type="button" name="Submit" value=" 关 闭 " onclick="window.close();">
      </div></td>
  </tr>
  <tr> 
    <td height="10" colspan="2">&nbsp;</td>
  </tr>
</table>
</body>
<script>
  document.all.textfield.select();
</script>
