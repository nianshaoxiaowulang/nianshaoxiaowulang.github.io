<%
Dim NewsFieldName,NewsFieldEName,NewsFieldType,NewsClassFieldName,NewsClassFieldEName,NewsClassFieldType
Dim DownloadFieldName,DownloadFieldEName,DownloadFieldType
NewsFieldName = Array("���","���ű��","����","������","������ʽ","������Ŀ���","�Ƿ��������","��������·��","����","�Ƿ��û�Ͷ��","�����ļ���","�ļ���չ��","����·��","�������","�ؼ���","������Դ","��������","���α༭","�������","�Ƿ��Ƽ�","�Ƿ�ͼƬ����","ͼƬ·��","ͼƬ�������ֵ���","�Ƿ����","�Ƿ�ɾ��","ɾ��ʱ��","�������Ȩ��","�Ƿ���ʾ����","�Ƿ���������","�Ƿ�������","�Ƿ��������","�Ƿ񹫸�����","�Ƿ���ʾ�ڲ�����","�Ƿ�õ�Ƭ����","����ģ���ļ���","�Ƿ񽹵�ͼƬ����","�Ƿ񾫲�ͼƬ����","�Ƿ����ͷ��","����ר����","�������ʾ����")
NewsFieldEName = Array("ID","NewsID","Title","SubTitle","TitleStyle","ClassID","HeadNewsTF","HeadNewsPath","Content","ManuTF","FileName","FileExtName","Path","AddDate","KeyWords","TxtSource","Author","Editer","ClickNum","RecTF","PicNewsTF","PicPath","NaviWords","AuditTF","DelTF","DelTime","BrowPop","ShowReviewTF","ReviewTF","SBSNews","MarqueeNews","ProclaimNews","LinkTF","FilterNews","NewsTemplet","FocusNewsTF","ClassicalNewsTF","TodayNewsTF","SpecialID","TitleShowReview")
NewsFieldType = Array("16","100","100","100","100","100","116","100","230","116","100","100","100","7","100","100","100","100","116","116","116","100","100","116","116","7","116","116","116","116","116","116","116","116","100","116","116","116","230","116")
NewsClassFieldName = Array("���","��Ŀ���","��ĿӢ������","��Ŀ��������","����Ŀ���","����Ŀ����","��Ŀ����ģ��","��������ģ��","��������ģ��","�Ƿ�����Ͷ��","�Ƿ���ʾ","��Ŀ���ʱ��","ɾ�����","ɾ��ʱ��","�����ļ�����·��","�����ļ���չ��","��Ŀ���Ȩ��","�����������","���Ź鵵ʱ��","��������","�Ƿ�Ϊ�ⲿ��Ŀ","�ⲿ��Ŀ��ַ","��Ʒ����ģ��","��Ŀת��")
NewsClassFieldEName = Array("ID","ClassID","ClassEName","ClassCName","ParentID","ChildNum","ClassTemp","NewsTemp","DownLoadTemp","Contribution","ShowTF","AddTime","DelFlag","DelTime","SaveFilePath","FileExtName","BrowPop","DoMain","FileTime","Orders","IsOutClass","ClassLink","ProductTemp","RedirectList")
NewsClassFieldType = Array("16","100","100","100","100","116","100","100","100","116","116","7","116","7","100","100","116","100","7","116","116","100","100","100")
DownloadFieldName = Array("���","���ر��","��������","������Ŀ���","�汾��","��������","��������","����","��Ȩ","�ļ���С","����","ϵͳƽ̨","��ϵ��EMAIL","�ṩ��Url��ַ","������","��ʾͼƬ","����Ȩ��","���","��ѹ����","���ʱ��","�޸�ʱ��","�Ƿ��Ƽ�","�Ƿ����","����ҳ���ļ���","ҳ���ļ���չ��","���ش���","����ģ����","�Ƿ���������","�Ƿ���ʾ����")
DownloadFieldEName = Array("ID","DownLoadID","Name","ClassID","Version","Types","Property","Language","Accredit","FileSize","Appraise","SystemType","EMail","ProviderUrl","Provider","Pic","BrowPop","Description","PassWord","AddTime","EditTime","RecTF","AuditTF","FileName","FileExtName","ClickNum","NewsTemplet","ReviewTF","ShowReviewTF")
DownloadFieldType = Array("16","100","100","100","100","116","116","100","116","100","116","100","100","100","100","100","116","230","100","7","7","116","116","100","100","116","100","116","116")

'�����ֶ�Ӣ�����������е����
Function GetIndexOfField(FieldEName,FieldENameArray)
	Dim i,FiledName
	FiledName = Lcase(FieldEName)
	For i = 0 to UBound(FieldENameArray)
		if lcase(FieldENameArray(i)) = FiledName Then
			GetIndexOfField = i
			Exit Function
		End if
	Next
	GetIndexOfField = -1
End Function
%>