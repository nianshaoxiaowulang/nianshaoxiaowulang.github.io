<% Option Explicit %>
<!--#include file="../../Inc/Cls_DB.asp" -->
<!--#include file="../../Inc/Const.asp" -->
<%
Dim RsConfigLoginobj,HelpTF,DBC,Conn
Set DBC = New databaseclass
Set Conn = DBC.openconnection()
Set DBC = Nothing
Set RsConfigLoginobj = Conn.execute("Select HelpTF from FS_Config")
if Not RsConfigLoginobj.Eof then
	HelpTF = RsConfigLoginobj("HelpTF")
else
	HelpTF = 1
end if
Set RsConfigLoginobj = Nothing
Set Conn = Nothing
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<script language="JavaScript">
function StartShrink(e)
{
	var myTarget;
	if (document.all) myTarget = e.toElement;
	else myTarget = e.relatedTarget;
	if (myTarget == null) BeginShrink();
	return;
}

function StartEnlarge(e)
{
	var myTarget;
	if (!CanShowNavTree())	return;
	if (document.all) myTarget = e.fromElement;
	else myTarget = e.relatedTarget;
	if (myTarget == null) setTimeout("BeginEnlarge();", 1);
	return;
}

function BeginShrink ()
{
	bEnlargePending = false;
	if (bPinned == false)
	{
		if (iFrameSize == 0)
		{
			if (bBusy == true)
			{
				bShrinkPending = true;
				setTimeout("PollShrink();", 100);
				return;
			}
		}
		else
		{
			if (bBusy == false)
			{
				iCurrentShrinkCount = 0;
				if (bShrinkPending == false)
				{
					bShrinkPending = true;
					setTimeout("DelayShrink()", 10);
				}
			}
		}
	}
}

function PollShrink()
{
	if (bShrinkPending == true)
	{
		if (bBusy == true)
		{
			setTimeout("PollShrink();", 100);
			return;
		}
		else
		{
			iCurrentShrinkCount = 0;
			setTimeout("DelayShrink()", 10);
		}
	}
}

function DelayShrink ()
{
	if (bShrinkPending == true)
	{
		iCurrentShrinkCount += 10;
		if (iCurrentShrinkCount > constMaxShrinkCount)
		{
			iCurrentShrinkCount = 0;
			ResizeFrame();
			return;
		}
		setTimeout("DelayShrink()", 10);
	}
}

function CanShowNavTree()
{
	if (typeof(top.CanShowNavTree) == "function") return (top.CanShowNavTree());
	return (true);
}
</script>
</head>
	<frameset id="nav_divider" name="nav_divider" rows="30,*<% if HelpTF = 1 then %>,10,140<% end if %>" frameborder="yes" framespacing="0" border="2" bordercolor="#ff0000">
		<frame noresize id="nav_toolbar" name="nav_toolbar" marginheight="0" marginwidth="0" frameborder="no"
			scrolling="no" src="Menu_Buttons.asp">
		<frame noresize id="nav_folder_area" name="nav_folder_area" marginwidth="0" frameborder="no"
			scrolling="auto" src="ShortCutPage.asp"><% if HelpTF = 1 then %>
		<frame noresize id="ResizeBar" name="ResizeBar" marginwidth="0" frameborder="no" scrolling="no" src="../Help/ResizeBar.htm">
		<frame noresize id="FSHelp" name="FSHelp" marginwidth="0" frameborder="yes" scrolling="auto" src="../Help/ShowHelp.asp"><% end if %></frameset><noframes></noframes>
</html>
