<% Option Explicit %>
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
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
Dim SpecialID,Types,NewsID,NewsObj,SpecialIDStr
if Request("Types")<>"" and Request("NewsID")<>"" and Request("SpecialID")<>"" then
	if Not JudgePopedomTF(Session("Name"),"P020401") then Call ReturnError()
	Types = Request("Types")
	NewsID = Request("NewsID")
	SpecialID = Request("SpecialID")
elseif Request("SpecialID") <> "" and Request("Types") = "" then
	if Not JudgePopedomTF(Session("Name"),"P020300") then Call ReturnError()
	SpecialID = Cstr(Request("SpecialID"))
	Types = ""
else
	Response.Write("<script>alert(""�������ݴ���"");window.close();</script>")
	Response.End
end if 
%>
<html>
<head>
<link href="../../../CSS/ModeWindow.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Ƶ��/ר��ɾ��</title>
</head>
<body leftmargin="0" topmargin="0">
<%If Types = "" then %>
<form action="" name="JSDellForm" method="post">
<table width="100%" border="0" cellspacing="5" cellpadding="0">
  <tr> 
    <td width="26%"><div align="center"><img src="../../Images/Question.gif" width="39" height="37"></div></td>
    <td width="74%">��ȷ��Ҫɾ����ѡƵ��/ר��?</td>
    </tr>
  <tr> 
    <td colspan="3"><div align="center"> 
          <input type="submit" name="Submit" value=" ȷ �� ">
          <input type="hidden" name="action" value="trues">
          <input type="button" name="Submit2" value=" ȡ �� " onClick="window.close();">
      </div></td>
    </tr>
</table>
</form>
<%else%>
<form action="" name="JSDellForm" method="post">
<table width="100%" border="0" cellspacing="5" cellpadding="0">
  <tr> 
    <td colspan="2"><div align="center"><img src="../../Images/Question.gif" width="39" height="37"></div></td>
    <td width="74%" height="2">��ȷ��Ҫ��Ƶ��/ר����ɾ����ѡ����?</td>
    </tr>
  <tr> 
    <td colspan="3"><div align="center"> 
          <input type="submit" name="Submit" value=" ȷ �� ">
          <input name="newsaction" type="hidden" id="newsaction" value="trues">
          <input type="button" name="Submit2" value=" ȡ �� " onClick="window.close();">
      </div></td>
    </tr>
</table>
</form>
<%end if%>
</body>
</html>
<%
if request.Form("action")="trues" then
	if Not JudgePopedomTF(Session("Name"),"P020300") then Call ReturnError()
	Dim DspArray,Dsp_i
	DspArray = Array("")
	DspArray = Split(SpecialID,",")
	For Dsp_i = 0 to UBound(DspArray)
		Conn.Execute("delete from FS_Special where SpecialID='"&DspArray(Dsp_i)&"'")
		'----------�޸����ű��SpcialID�ֶ�--------
		Dim ModifyNewsObj,ModSpecialIDStr
		Set ModifyNewsObj = Conn.Execute("Select SpecialID,NewsID from FS_News where SpecialID like '%"&DspArray(Dsp_i)&"%'")
		while not ModifyNewsObj.eof
			ModSpecialIDStr = ModifyNewsObj("SpecialID")
			ModSpecialIDStr = Replace(ModSpecialIDStr,DspArray(Dsp_i),"")
			ModSpecialIDStr = Replace(ModSpecialIDStr,",,",",")
			If left(ModSpecialIDStr,1)="," then
				ModSpecialIDStr = right(ModSpecialIDStr,len(ModSpecialIDStr)-1)
			end if
			If right(ModSpecialIDStr,1)="," then
				ModSpecialIDStr = left(ModSpecialIDStr,len(ModSpecialIDStr)-1)
			end if
			Conn.Execute("Update FS_News set SpecialID='"&ModSpecialIDStr&"' where NewsID='"&ModifyNewsObj("NewsID")&"'")
			ModifyNewsObj.MoveNext
		wend
		ModifyNewsObj.Close
		Set ModifyNewsObj = Nothing
	Next
	Response.write("<script>dialogArguments.location.reload();window.close();</script>")
	Response.end
end if
	
if Request.Form("newsaction")="trues" then
	if Not JudgePopedomTF(Session("Name"),"P020401") then Call ReturnError()
	Dim NewsArray,DSN_i
	NewsArray = Array("")
	NewsArray = Split(NewsID,"***")
	For DSN_i = 0 to UBound(NewsArray)
		Set NewsObj = Conn.Execute("Select SpecialID from FS_News where NewsID='" & NewsArray(DSN_i) & "'")
		SpecialIDStr = NewsObj("SpecialID")
		SpecialIDStr = Replace(SpecialIDStr,SpecialID,"")
		SpecialIDStr = Replace(SpecialIDStr,",,",",")
		If left(SpecialIDStr,1)="," then
			SpecialIDStr = right(SpecialIDStr,len(SpecialIDStr)-1)
		end if
		If right(SpecialIDStr,1)="," then
			SpecialIDStr = left(SpecialIDStr,len(SpecialIDStr)-1)
		end if
		Conn.Execute("Update FS_News set SpecialID='"&SpecialIDStr&"' where NewsID='"&NewsArray(DSN_i)&"'")
		NewsObj.close
		Set NewsObj = Nothing
	Next
	Response.write("<script>dialogArguments.location.reload();window.close();</script>")
	Response.end
end if
Set Conn = Nothing
%>