<% Option Explicit %>
<!--#include file="../../../Inc/NosqlHack.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<%
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
'==============================================================================
'������ƣ���Ѷ��վ��Ϣ����ϵͳ
'��ǰ�汾��Foosun Content Manager System(FoosunCMS V3.1.0930)
'���¸��£�2005.10
'==============================================================================
'Copyright (C) 2002-2004 Foosun.Net  All rights reserved.
'��ҵע����ϵ��028-85098980-601,��Ŀ������028-85098980-606��609,�ͻ�֧�֣�608
'��Ʒ��ѯQQ��394226379,159410,125114015
'����֧��QQ��315485710,66252421 
'��Ŀ����QQ��415637671��655071
'���򿪷����Ĵ���Ѷ�Ƽ���չ���޹�˾(Foosun Inc.)
'Email:service@Foosun.cn
'MSN��skoolls@hotmail.com
'��̳֧�֣���Ѷ������̳(http://bbs.foosun.net)
'�ٷ���վ��www.Foosun.cn  ��ʾվ�㣺test.cooin.com 
'��վͨϵ��(���ܿ��ٽ�վϵ��)��www.ewebs.cn
'==============================================================================
'��Ѱ汾���ڳ�����ҳ������Ȩ��Ϣ�������ϱ�վLOGO��������
'��Ѷ��˾�����˳���ķ���׷��Ȩ��
'�������2�ο��������뾭����Ѷ��˾������������׷����������
'==============================================================================

%>
<!--#include file="../../../Inc/Session.asp" -->

<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P020402") then Call ReturnError()
Dim MoveOrCopyClassPara
If Request("MoveOrCopyClassPara") = "" then
	Response.Write("<script>alert(""�������ݴ���"");window.close();</script>")
	Response.End
Else
	MoveOrCopyClassPara = Request("MoveOrCopyClassPara")
	Dim MoveTF,MoveTFText,ResultString,TempSpecialIDStr,IntentEName,TempNewsObj,IntentCName,SourceCName,SourceEName,DoNews,IntentObj,SourceObj,IntentFileSql,IntentFileObj,SourceFileObj
	ResultString = ""
	MoveTF = GetParaValue(MoveOrCopyClassPara,"MoveTF")
	If MoveTF="true" then
		MoveTFText = "�ƶ�"
	Else
		MoveTFText = "����"
	End If
	IntentEName = GetParaValue(MoveOrCopyClassPara,"ObjectClass") 'Ŀ��JS
	SourceEName = GetParaValue(MoveOrCopyClassPara,"SourceClass") 'ԴJS
	DoNews = GetParaValue(MoveOrCopyClassPara,"SourceNews") '�������� 
	Set IntentObj = Conn.Execute("Select CName from FS_Special where SpecialID='"&IntentEName&"'")
	If IntentObj.eof then
		Set IntentObj = Nothing
		Set Conn = Nothing
		Response.Write("<script>alert(""Ŀ��ר���Ѿ�������,��ˢ�º��ٽ��в���"");dialogArguments.location.reload();window.close();</script>")
		Response.End
	Else
		IntentCName = IntentObj("CName")
	End if
	Set SourceObj = Conn.Execute("Select CName from FS_Special where SpecialID='"&SourceEName&"'")
	If SourceObj.eof then
		Set SourceObj = Nothing
		Set Conn = Nothing
		Response.Write("<script>alert(""Դר���Ѿ�������,��ˢ�º��ٽ��в���"");dialogArguments.location.reload();window.close();</script>")
		Response.End
	Else
		SourceCName = SourceObj("CName")
	End if
	If Cstr(IntentEName) = Cstr(SourceEName) then
		Set IntentObj = Nothing
		Set SourceObj = Nothing
		Set Conn = Nothing
		Response.Write("<script>alert(""Դר���Ŀ��ר�ⲻ����ͬ"");window.close();</script>")
		Response.End
	End if

	If Request.Form("action") = "trues" then
		Dim DoNewsArray,SpMove_i
		DoNewsArray = Array("")
		DoNewsArray = Split(DoNews,"***")
		For SpMove_i = 0 to UBound(DoNewsArray)
			Set TempNewsObj = Conn.Execute("Select SpecialID from FS_News where NewsID='"&DoNewsArray(SpMove_i)&"'")
			If Not TempNewsObj.eof then
				If MoveTF = "true" then '�ƶ�����
					TempSpecialIDStr = Replace(TempNewsObj("SpecialID"),SourceEName,IntentEName) 
					Conn.Execute("Update FS_News set SpecialID='"&TempSpecialIDStr&"' where NewsID='"&Cstr(DoNewsArray(SpMove_i))&"'")
					ResultString = "�����ƶ��ɹ�"
				Else '��������
					TempSpecialIDStr = TempNewsObj("SpecialID")&","&IntentEName
					Conn.Execute("Update FS_News set SpecialID='"&TempSpecialIDStr&"' where NewsID='"&Cstr(DoNewsArray(SpMove_i))&"'")
					ResultString = "���ſ����ɹ�"
				End if
			End If
			TempNewsObj.Close
			Set TempNewsObj = Nothing
		Next  
	End If
	If ResultString <> "" then
		Response.Write("<script>dialogArguments.location.reload();window.close();</script>")
		Response.End
	End If
End If
  
Function GetParaValue(ParaStr,ParaName)
	Dim BeginIndex,EndIndex
	BeginIndex = InStr(ParaStr,ParaName)+Len(ParaName)+1
	EndIndex = InStr(BeginIndex,ParaStr,",")
	GetParaValue = Mid(ParaStr,BeginIndex,EndIndex-BeginIndex)
End Function
%>

<html>
<head>
<link href="../../../CSS/ModeWindow.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�ƶ�����ר������</title>
</head>
<body topmargin="0" leftmargin="0" >
<form action="" method="post" name="MoveForm">
<table width="100%" border="0" cellspacing="3" cellpadding="0">
  <tr> 
    <td width="7%" height="10">&nbsp;</td>
    <td width="12%" rowspan="3"><div align="center"><img src="../../Images/Question.gif" width="39" height="37"></div></td>
    <td width="75%">&nbsp;</td>
    <td width="6%" height="10">&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td>��ȷ��Ҫ�������Ŵ�<font color="#FF0000"><%=SourceCName%></font><font color="#0000FF"><%=MoveTFText%></font>��<font color=red><%=IntentCName%></font>?</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td height="2">&nbsp;</td>
    <td height="2">&nbsp;</td>
    <td height="2">&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td colspan="2"><div align="center"> 
        <input type="submit" name="Submit" value=" ȷ �� ">
        <input type="hidden" name="action" value="trues">
        <input type="button" name="Submit2" value=" ȡ �� " onClick="window.close();">
      </div></td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td height="10">&nbsp;</td>
    <td height="10" colspan="2">&nbsp;</td>
    <td height="10">&nbsp;</td>
  </tr>
</table>
</form>
</body>
</html>
