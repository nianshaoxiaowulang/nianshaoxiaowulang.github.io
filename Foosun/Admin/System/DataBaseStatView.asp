<% Option Explicit %>
<!--#include file="../../../Inc/NosqlHack.asp" -->
<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<%
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
Dim DBC,Conn
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P0406001") then Call ReturnError1()
	Dim Types,I,DBSVSql,RsDBSVObj,TempNums,RsTempObj
	Types = Request("Types")
%>
<html>
<head>
<link href="../../../CSS/FS_css.css" rel="stylesheet">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>����ͳ����ʾ</title>
</head>
<body topmargin="2" leftmargin="2" ondragstart="return false;"  oncontextmenu="return false;">
<%
	Select Case Types
		Case "Administrator"
			Set RsDBSVObj = Server.CreateObject(G_FS_RS)
			DBSVSql = "Select GroupName,ID from FS_AdminGroup order by ID asc"
			RsDBSVObj.Open DBSVSql,Conn,1,1
			Set RsTempObj = Conn.Execute("Select count(ID) from FS_Admin")
			TempNums = RsTempObj(0)+1
			RsTempObj.Close
			Set RsTempObj = Nothing
%>
<table border=0 align="center" cellpadding=2>
  <tr>
    <td colspan="4" align=center>����Ա����ͳ��</td>
  </tr>
  <tr>
    <td colspan="4" align=center><table border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td valign="top">
		  <table border=0 align=center cellpadding=0 cellspacing=0>
              <tr> 
                <td height="25" valign=top align="right" nowrap>100%</td>
              </tr>
              <tr> 
                <td  height="25" valign=top align="right"  nowrap>75%</td>
              </tr>
              <tr> 
                <td  height="25" valign=top  align="right" nowrap>50%</td>
              </tr>
              <tr>
                <td  height="25" valign=top  align="right" nowrap>25%</td>
              </tr>
              <tr> 
                <td  align="right" valign=top nowrap>0</td>
              </tr>
          </table>		  </td>
          <td valign="bottom">
            <table align=center>
              <tr valign=bottom >
                <td width="15" align=center nowrap background=../../Images/Visit/tu_back.gif><img src=../../Images/Visit/bar.gif width="15" height="100" id=htav><br>
                </td>
                <% 
				do while not RsDBSVObj.eof 
					Set RsTempObj = Conn.Execute("Select count(ID) from FS_Admin where GroupID="&RsDBSVObj("ID")&"")
				%>
                <td width="15" align=center nowrap background=../../Images/Visit/tu_back.gif><img src="../../Images/Visit/bar.gif" width="15" height="<% =100*RsTempObj(0)/TempNums%>" id=htav><br>
                </td>
                <% 
					RsTempObj.Close
					Set RsTempObj = Nothing
				RsDBSVObj.MoveNext
				Loop
				RsDBSVObj.Close
				%>
                <td></td>
              </tr>
			  <tr>
				<td width="15" rowspan="2" align=center valign="top">��<br>(<%=TempNums-1%>)</td>
				<% 
				RsDBSVObj.Open DBSVSql,Conn,1,1
				do while not RsDBSVObj.eof 
					Set RsTempObj = Conn.Execute("Select count(ID) from FS_Admin where GroupID="&RsDBSVObj("ID")&"")
				%>
				<td width="15" rowspan="2" align=center valign="top"><% = RsDBSVObj("GroupName")%><br>(<%=RsTempObj(0)%>)</td>
				<%
						RsTempObj.Close
						Set RsTempObj = Nothing
					RsDBSVObj.MoveNext
					Loop
					RsDBSVObj.Close
					Set RsDBSVObj = Nothing
				%>
				<td width="76" align=center valign="top">��λ���飩</td>
			  </tr>
			  <tr>
			    <td align=center valign="bottom">��λ���ˣ�</td>
		      </tr>
          </table>
		  </td>
        </tr>
    </table></td>
  </tr>
</table>
<%		
		Case "Members"
			Set RsDBSVObj = Server.CreateObject(G_FS_RS)
			DBSVSql = "Select Name,GroupID from FS_MemGroup order by ID asc"
			RsDBSVObj.Open DBSVSql,Conn,1,1
			Set RsTempObj = Conn.Execute("Select count(ID) from FS_Members")
			TempNums = RsTempObj(0)+1
			RsTempObj.Close
			Set RsTempObj = Nothing
%>
<table border=0 align="center" cellpadding=2>
  <tr>
    <td colspan="4" align=center>��Ա����ͳ��</td>
  </tr>
  <tr>
    <td colspan="4" align=center><table border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td valign="top">
		  <table border=0 align=center cellpadding=0 cellspacing=0>
              <tr> 
                <td height="25" valign=top align="right" nowrap>100%</td>
              </tr>
              <tr> 
                <td  height="25" valign=top align="right"  nowrap>75%</td>
              </tr>
              <tr> 
                <td  height="25" valign=top  align="right" nowrap>50%</td>
              </tr>
              <tr>
                <td  height="25" valign=top  align="right" nowrap>25%</td>
              </tr>
              <tr> 
                <td  align="right" valign=top nowrap>0</td>
              </tr>
          </table></td>
          <td valign="bottom">
            <table align=center>
              <tr valign=bottom >
                <td width="15" align=center nowrap background=../../Images/Visit/tu_back.gif><img src=../../Images/Visit/bar.gif width="15" height="100" id=htav><br>
                </td>
                <% 
				do while not RsDBSVObj.eof 
					Set RsTempObj = Conn.Execute("Select count(ID) from FS_Members where GroupID='"&RsDBSVObj("GroupID")&"'")
				%>
                <td width="15" align=center nowrap background=../../Images/Visit/tu_back.gif><img src=../../Images/Visit/bar.gif width="15" height="<% =100*RsTempObj(0)/TempNums%>" id=htav><br>
                </td>
                <% 
					RsTempObj.Close
					Set RsTempObj = Nothing
				RsDBSVObj.MoveNext
				Loop
				RsDBSVObj.Close
				%>
                <td></td>
              </tr>
			  <tr>
				<td width="15" rowspan="2" align=center valign="top">��<br>(<%=TempNums-1%>)</td>
				<% 
				RsDBSVObj.Open DBSVSql,Conn,1,1
				do while not RsDBSVObj.eof 
					Set RsTempObj = Conn.Execute("Select count(ID) from FS_Members where GroupID='"&RsDBSVObj("GroupID")&"'")
				%>
				<td width="15" rowspan="2" align=center valign="top"><% = RsDBSVObj("Name")%><br>(<%=RsTempObj(0)%>)</td>
				<%
						RsTempObj.Close
						Set RsTempObj = Nothing
					RsDBSVObj.MoveNext
					Loop
					RsDBSVObj.Close
					Set RsDBSVObj = Nothing
				%>
				<td width="76" align=center valign="top">��λ���飩</td>
			  </tr>
			  <tr>
			    <td align=center valign="bottom">��λ���ˣ�</td>
		      </tr>
          </table>
		  </td>
        </tr>
    </table></td>
  </tr>
</table>
<%
		Case "News_Month"
		Dim ChooseYear,Temp_i
			ChooseYear = Request("ChooseYear")
			If ChooseYear = "" then
				ChooseYear = Year(Now())
			End If
			Set RsTempObj = Conn.Execute("Select count(ID) from FS_News where year(AddDate)='"&ChooseYear&"'")
			TempNums = RsTempObj(0)+1
			RsTempObj.Close
			Set RsTempObj = Nothing
%>
<table border=0 align="center" cellpadding=2>
  <tr>
    <td colspan="4" align=center>����(�·�)����ͳ��</td>
  </tr>
  <tr>
    <td colspan="4" align=center><table border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td valign="top">
		  <table border=0 align=center cellpadding=0 cellspacing=0>
              <tr> 
                <td height="25" valign=top align="right" nowrap>100%</td>
              </tr>
              <tr> 
                <td  height="25" valign=top align="right"  nowrap>75%</td>
              </tr>
              <tr> 
                <td  height="25" valign=top  align="right" nowrap>50%</td>
              </tr>
              <tr>
                <td  height="25" valign=top  align="right" nowrap>25%</td>
              </tr>
              <tr> 
                <td  align="right" valign=top nowrap>0</td>
              </tr>
          </table></td>
          <td valign="bottom">
            <table align=center>
              <tr valign=bottom >
                <td width="15" align=center nowrap background=../../Images/Visit/tu_back.gif><img src=../../Images/Visit/bar.gif width="15" height="100" id=htav><br>
                </td>
                <% 
				For Temp_i=1 to 12
					Set RsTempObj = Conn.Execute("Select count(ID) from FS_News where year(AddDate)='"&ChooseYear&"' and month(AddDate)='"&Temp_i&"'")
				%>
                <td width="15" align=center nowrap background=../../Images/Visit/tu_back.gif><img src=../../Images/Visit/bar.gif width="15" height="<% =100*RsTempObj(0)/TempNums%>" id=htav><br>
                </td>
                <% 
					RsTempObj.Close
					Set RsTempObj = Nothing
				Next
				%>
                <td></td>
              </tr>
			  <tr>
				<td width="15" rowspan="2" align=center valign="top"><%=ChooseYear%>����<br>(<%=TempNums-1%>)</td>
                <% 
				For Temp_i=1 to 12
					Set RsTempObj = Conn.Execute("Select count(ID) from FS_News where year(AddDate)='"&ChooseYear&"' and month(AddDate)='"&Temp_i&"'")
				%>
				<td width="15" rowspan="2" align=center valign="top"><% = Temp_i%>�·�<br>(<%=RsTempObj(0)%>)</td>
				<%
					RsTempObj.Close
					Set RsTempObj = Nothing
				Next
				%>
				<td width="76" align=center valign="top">��λ���£�</td>
			  </tr>
			  <tr>
			    <td align=center valign="bottom">��λ������</td>
		      </tr>
          </table>
		  </td>
        </tr>
    </table></td>
  </tr>
</table>
<%		
		Case "News_Class"
			Set RsDBSVObj = Server.CreateObject(G_FS_RS)
			DBSVSql = "Select ClassCName,ClassID from FS_NewsClass order by AddTime asc"
			RsDBSVObj.Open DBSVSql,Conn,1,1
			Set RsTempObj = Conn.Execute("Select count(ID) from FS_News")
			TempNums = RsTempObj(0)+1
			RsTempObj.Close
			Set RsTempObj = Nothing
%>
<table border=0 align="center" cellpadding=2>
  <tr>
    <td colspan="4" align=center>����(��Ŀ)����ͳ��</td>
  </tr>
  <tr>
    <td colspan="4" align=center><table border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td valign="top">
		  <table border=0 align=center cellpadding=0 cellspacing=0>
              <tr> 
                <td height="25" valign=top align="right" nowrap>100%</td>
              </tr>
              <tr> 
                <td  height="25" valign=top align="right"  nowrap>75%</td>
              </tr>
              <tr> 
                <td  height="25" valign=top  align="right" nowrap>50%</td>
              </tr>
              <tr>
                <td  height="25" valign=top  align="right" nowrap>25%</td>
              </tr>
              <tr> 
                <td  align="right" valign=top nowrap>0</td>
              </tr>
          </table></td>
          <td valign="bottom">
            <table align=center>
              <tr valign=bottom >
                <td width="15" align=center nowrap background=../../Images/Visit/tu_back.gif><img src=../../Images/Visit/bar.gif width="15" height="100" id=htav><br>
                </td>
                <% 
				do while not RsDBSVObj.eof 
					Set RsTempObj = Conn.Execute("Select count(ID) from FS_News where DelTF=0 and ClassID='"&RsDBSVObj("ClassID")&"'")
				%>
                <td width="15" align=center nowrap background=../../Images/Visit/tu_back.gif><img src=../../Images/Visit/bar.gif width="15" height="<% =100*RsTempObj(0)/TempNums%>" id=htav><br>
                </td>
                <% 
					RsTempObj.Close
					Set RsTempObj = Nothing
				RsDBSVObj.MoveNext
				Loop
				RsDBSVObj.Close
				%>
                <td></td>
              </tr>
			  <tr>
				<td width="15" rowspan="2" align=center valign="top">��<br>(<%=TempNums-1%>)</td>
				<% 
				RsDBSVObj.Open DBSVSql,Conn,1,1
				do while not RsDBSVObj.eof 
					Set RsTempObj = Conn.Execute("Select count(ID) from FS_News where DelTF=0 and ClassID='"&RsDBSVObj("ClassID")&"'")
				%>
				<td width="15" rowspan="2" align=center valign="top"><% = RsDBSVObj("ClassCName")%><br>(<%=RsTempObj(0)%>)</td>
				<%
						RsTempObj.Close
						Set RsTempObj = Nothing
					RsDBSVObj.MoveNext
					Loop
					RsDBSVObj.Close
					Set RsDBSVObj = Nothing
				%>
				<td width="76" align=center valign="top">��λ������</td>
			  </tr>
			  <tr>
			    <td align=center valign="bottom">��λ������</td>
		      </tr>
          </table>
		  </td>
        </tr>
    </table></td>
  </tr>
</table>
<%		
		Case "Contribution"
			Set RsDBSVObj = Server.CreateObject(G_FS_RS)
			DBSVSql = "Select ClassCName,ClassID from FS_NewsClass where Contribution=1 order by AddTime asc"
			RsDBSVObj.Open DBSVSql,Conn,1,1
			Set RsTempObj = Conn.Execute("Select count(ContID) from FS_Contribution")
			TempNums = RsTempObj(0)+1
			RsTempObj.Close
			Set RsTempObj = Nothing
%>
<table border=0 align="center" cellpadding=2>
  <tr>
    <td colspan="4" align=center>�������ͳ��</td>
  </tr>
  <tr>
    <td colspan="4" align=center><table border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td valign="top">
		  <table border=0 align=center cellpadding=0 cellspacing=0>
              <tr> 
                <td height="25" valign=top align="right" nowrap>100%</td>
              </tr>
              <tr> 
                <td  height="25" valign=top align="right"  nowrap>75%</td>
              </tr>
              <tr> 
                <td  height="25" valign=top  align="right" nowrap>50%</td>
              </tr>
              <tr>
                <td  height="25" valign=top  align="right" nowrap>25%</td>
              </tr>
              <tr> 
                <td  align="right" valign=top nowrap>0</td>
              </tr>
          </table></td>
          <td valign="bottom">
            <table align=center>
              <tr valign=bottom >
                <td width="15" align=center nowrap background=../../Images/Visit/tu_back.gif><img src=../../Images/Visit/bar.gif width="15" height="100" id=htav><br>
                </td>
                <% 
				do while not RsDBSVObj.eof 
					dim TemNum
					Set RsTempObj = Conn.Execute("Select count(ContID) from FS_Contribution where ClassID='"&RsDBSVObj("ClassID")&"'")
					TemNum = RsTempObj(0)
				%>
                <td width="15" align=center nowrap background=../../Images/Visit/tu_back.gif><img src=../../Images/Visit/bar.gif width="15" height="<% =100*TemNum/TempNums%>" id=htav><br>
                </td>
                <% 
					RsTempObj.Close
					Set RsTempObj = Nothing
				RsDBSVObj.MoveNext
				Loop
				RsDBSVObj.Close
				%>
                <td></td>
              </tr>
			  <tr>
				<td width="15" rowspan="2" align=center valign="top">��<br>(<%=TempNums-1%>)</td>
				<% 
				RsDBSVObj.Open DBSVSql,Conn,1,1
				do while not RsDBSVObj.eof 
					Set RsTempObj = Conn.Execute("Select count(ContID) from FS_Contribution where  ClassID='"&RsDBSVObj("ClassID")&"'")
				%>
				<td width="15" rowspan="2" align=center valign="top"><% = RsDBSVObj("ClassCName")%><br>(<%=RsTempObj(0)%>)</td>
				<%
						RsTempObj.Close
						Set RsTempObj = Nothing
					RsDBSVObj.MoveNext
					Loop
					RsDBSVObj.Close
					Set RsDBSVObj = Nothing
				%>
				<td width="76" align=center valign="top">��λ������</td>
			  </tr>
			  <tr>
			    <td align=center valign="bottom">��λ������</td>
		      </tr>
          </table>
		  </td>
        </tr>
    </table></td>
  </tr>
</table>
<%	
	Case "NewsClass"
			ChooseYear = Request("ChooseYear")
			If ChooseYear = "" then
				ChooseYear = Year(Now())
			End If
			Set RsTempObj = Conn.Execute("Select count(ID) from FS_NewsClass where year(AddTime)='"&ChooseYear&"'")
			TempNums = RsTempObj(0)+1
			RsTempObj.Close
			Set RsTempObj = Nothing
%>
<table border=0 align="center" cellpadding=2>
  <tr>
    <td colspan="4" align=center>������Ŀ����ͳ��</td>
  </tr>
  <tr>
    <td colspan="4" align=center><table border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td valign="top">
		  <table border=0 align=center cellpadding=0 cellspacing=0>
              <tr> 
                <td height="25" valign=top align="right" nowrap>100%</td>
              </tr>
              <tr> 
                <td  height="25" valign=top align="right"  nowrap>75%</td>
              </tr>
              <tr> 
                <td  height="25" valign=top  align="right" nowrap>50%</td>
              </tr>
              <tr>
                <td  height="25" valign=top  align="right" nowrap>25%</td>
              </tr>
              <tr> 
                <td  align="right" valign=top nowrap>0</td>
              </tr>
          </table></td>
          <td valign="bottom">
            <table align=center>
              <tr valign=bottom >
                <td width="15" align=center nowrap background=../../Images/Visit/tu_back.gif><img src=../../Images/Visit/bar.gif width="15" height="100" id=htav><br>
                </td>
                <% 
				For Temp_i=1 to 12
					Set RsTempObj = Conn.Execute("Select count(ID) from FS_NewsClass where year(AddTime)='"&ChooseYear&"' and month(AddTime)='"&Temp_i&"'")
				%>
                <td width="15" align=center nowrap background=../../Images/Visit/tu_back.gif><img src=../../Images/Visit/bar.gif width="15" height="<% =100*RsTempObj(0)/TempNums%>" id=htav><br>
                </td>
                <% 
					RsTempObj.Close
					Set RsTempObj = Nothing
				Next
				%>
                <td></td>
              </tr>
			  <tr>
				<td width="15" rowspan="2" align=center valign="top"><%=ChooseYear%>����<br>(<%=TempNums-1%>)</td>
                <% 
				For Temp_i=1 to 12
					Set RsTempObj = Conn.Execute("Select count(ID) from FS_NewsClass where year(AddTime)='"&ChooseYear&"' and month(AddTime)='"&Temp_i&"'")
				%>
				<td width="15" rowspan="2" align=center valign="top"><% = Temp_i%>�·�<br>(<%=RsTempObj(0)%>)</td>
				<%
					RsTempObj.Close
					Set RsTempObj = Nothing
				Next
				%>
				<td width="76" align=center valign="top">��λ���£�</td>
			  </tr>
			  <tr>
			    <td align=center valign="bottom">��λ������</td>
		      </tr>
          </table>
		  </td>
        </tr>
    </table></td>
  </tr>
</table>
<%		
	End Select
%>

</body>
</html>
