<% Option Explicit %>
<%
Dim Url,FileName
FileName = Request("FileName")
Url = FileName
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><% = Request("PageTitle") %></title>
</head>
<link rel="stylesheet" type="text/css" href="../../CSS/FS_css.css">
<body>
<iframe src=<% = Url %> style="width:100%;height:100%;"  frameborder=0 scrolling="no"></iframe>
</body>
</html>
