<!--#include file="../../../Inc/Const.asp" -->
<!--#include file="../../../Inc/Cls_DB.asp" -->
<!--#include file="../../../Inc/Function.asp" -->
<%
'==============================================================================
'������ƣ�FoosunShop System Form FoosunCMS
'��ǰ�汾��Foosun Content Manager System 3.0 ϵ��
'���¸��£�2004.12
'==============================================================================
'��ҵע����ϵ��028-85098980-601,602 ����֧�֣�028-85098980-605��607,�ͻ�֧�֣�608
'��Ʒ��ѯQQ��159410,394226379,125114015,655071
'����֧��:���г���ʹ�����⣬�����ʵ�bbs.foosun.net���ǽ���ʱ�ش���
'���򿪷�����Ѷ������ & ��Ѷ���������
'��̳֧�֣���Ѷ������̳(http://bbs.foosun.net)
'�ٷ���վ��www.Foosun.net  ��ʾվ�㣺test.cooin.com    
'��վ����ר����www.cooin.com
'==============================================================================
'��Ѱ汾����������ҳ������Ȩ��Ϣ�������ϱ�վLOGO��������
'==============================================================================
Dim DBC,Conn,sRootDir
Set DBC = New DataBaseClass
Set Conn = DBC.OpenConnection()
Set DBC = Nothing
Dim RsAdminConfigObj
Set RsAdminConfigObj = Conn.Execute("Select SiteName,UserConfer,Copyright,isEmail,isChange,IsShop,MaxContent,QPoint from FS_Config")
%>
<!--#include file="../../../Inc/Session.asp" -->
<!--#include file="../../../Inc/CheckPopedom.asp" -->
<%
if Not JudgePopedomTF(Session("Name"),"P070704") then Call ReturnError1()
If request("Action")="Del" Then
	If trim(Request("GID"))<>"" Then
		Dim BookStr,ParaArray,i,NumStr,NumParaArray
		BookStr = Request("GID")
		  if Right(BookStr,1)="," then
			BookStr = Left(BookStr,Len(BookStr)-1)
		  end if
		  if Left(BookStr,1)="," then
			BookStr = Right(BookStr,Len(BookStr)-1)
		  end if
		  ParaArray = Split(BookStr,",")
		For i = LBound(ParaArray) to UBound(ParaArray)
			Dim GBListObj
			Set GBListObj = Conn.execute("Select ID,UserID From FS_GBook where id="&Clng(ParaArray(i)))
		    Conn.execute("Update FS_Members Set Point = Point-"&RsAdminConfigObj("QPoint")&"  where ID="&GBListObj("UserID"))
			Conn.execute("Delete From FS_GBook Where id="&Clng(ParaArray(i)))
		Next
		Response.Write("<script>alert(""ɾ���ɹ���"&CopyRight&""");location=""Sysbook.asp"";</script>")  
		Response.End
		'�۳�����
	Else
		Response.Write("<script>alert(""��ѡ��ɾ�������ӣ�"&CopyRight&""");location=""Sysbook.asp"";</script>")  
		Response.End
	End if
End If
%>