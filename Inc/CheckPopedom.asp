<%
Function GetPopedomList(AdminName)
	GetPopedomList = Session("PopedomList")
End Function 

Function JudgePopedomTF(AdminName,PopedomName)
	Dim PopList
	JudgePopedomTF = False
	if Session("GroupID") = "0" then
		JudgePopedomTF = True
	else
		PopList = GetPopedomList(AdminName)
		if (PopList <> "") and (PopedomName <> "") then
			if InStr(PopList,PopedomName) <> 0 then
				JudgePopedomTF = True
			else
				JudgePopedomTF = False
			end if
		else
			JudgePopedomTF = False
		end if
	end if
End Function  
Sub ReturnError()
	Response.Write("<script>alert(""[ϵͳ����]\n\n����Ȩ�޲���!����ϵͳ����Ա��ϵ\n"&Copyright&""");window.close();</script>")
	Response.End 
end Sub
Sub ReturnError1()
	Response.Write("<script>alert(""[ϵͳ����]\n\n����Ȩ�޲���!����ϵͳ����Ա��ϵ\n"&Copyright&""");location.href=""javascript:history.back()"";</script>")
	Response.End 
end Sub
Sub ReturnError2()
	Response.Write("loading...")
	Response.End 
end Sub
%>