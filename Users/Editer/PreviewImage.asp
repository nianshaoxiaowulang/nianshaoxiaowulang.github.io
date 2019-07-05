<% Option Explicit %>
<%
Dim PreviewImagePath
PreviewImagePath = Request("FilePath")
if PreviewImagePath = "" then
	PreviewImagePath = "../Images/DefaultPreview.gif"
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>图片预览</title>
</head>
<body topmargin="0" leftmargin="0">
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td align="center" valign="middle"><img src="<% = PreviewImagePath %>"></td>
  </tr>
</table>
</body>
</html>
