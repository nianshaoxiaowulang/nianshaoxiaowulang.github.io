////�жϵ�1��//////////////////////////Foosun Inc. ����� /////////////////////////////////////////////////////////////////
function checkdata() 
{
	if(UserForm.Username.value =="" ) {
	alert("�������û�����")
	UserForm.Username.focus()
	return false;
}
if(UserForm.Username.value.length<3 ) {
	alert("�û�������3���ַ���")
	UserForm.Username.focus()
	return false;
	}
if(UserForm.Username.value.length>18 ) {
	alert("�û������ܳ���18���ַ���")
	UserForm.Username.focus()
	return false;
	}
if(UserForm.sPassword.value =="" ) {
	alert("���������룡")
	UserForm.sPassword.focus()
	return false;
}
if(UserForm.sPassword.value.length<6 ) {
	alert("����Ӧ�ô���6С��18���ַ���")
	UserForm.sPassword.focus()
	return false;
}
if(UserForm.sPassword.value !==UserForm.Confimpass.value) {
	alert("2�����벻һ�£�")
	UserForm.sPassword.focus()
	return false;
	}
if(UserForm.email.value =="" ) {
	alert("����������ʼ���")
	UserForm.email.focus()
	return false;
}
if(UserForm.email.value.length<8 || UserForm.email.value.length>64 || !validateEmail(UserForm.email.value) ) {
		alert("����������ȷ�������ַ!")
		UserForm.email.focus()
		return false;
	}
if(UserForm.PassQuestion.value =="" ) {
	alert("�������������⣡")
	UserForm.PassQuestion.focus()
	return false;
}
if(UserForm.PassAnswer.value =="" ) {
	alert("����������𰸣�")
	UserForm.PassAnswer.focus()
	return false;
}
if(UserForm.PassQuestion.value ==UserForm.PassAnswer.value) {
	alert("������ʾ����ʹ𰸲�����ͬ����")
	UserForm.PassQuestion.focus()
	return false;
	}
}
////�жϵ�2��//////////////////////////Foosun Inc. ����� /////////////////////////////////////////////////////////////////
function checkdata1() 
{
	if(UserForm1.sName.value =="" ) {
	alert("������������ʵ������")
	UserForm1.sName.focus()
	return false;
}
if(UserForm1.sName.value.length<2 ) {
	alert("����ȷ��д������")
	UserForm1.sName.focus()
	return false;
}
if( !(UserForm1.Sex[0].checked || UserForm1.Sex[1].checked)) {
alert("��ѡ���Ա� !!")
return false;
}
	if(UserForm1.Ver.value =="" ) {
	alert("������Ч���룡")
	UserForm1.Ver.focus()
	return false;
}
if(UserForm1.tel.value =="" ) {
	alert("��������ϵ�绰���룡")
	UserForm1.tel.focus()
	return false;
}
if(UserForm1.tel.value.length<7 ) {
	alert("����ȷ��д�绰���룡")
	UserForm1.tel.focus()
	return false;
}
if(UserForm1.address.value =="" ) {
	alert("��������ϵ��ַ��")
	UserForm1.address.focus()
	return false;
}
if(UserForm1.address.value.length<5 ) {
	alert("����ȷ��д��ַ��")
	UserForm1.address.focus()
	return false;
}
if(UserForm1.email.value =="" ) {
	alert("����������ʼ���")
	UserForm1.email.focus()
	return false;
}
if(UserForm1.email.value.length<8 || UserForm1.email.value.length>64 || !validateEmail(UserForm1.email.value) ) {
		alert("����������ȷ�������ַ!")
		UserForm1.email.focus()
		return false;
	}
}
function CheckLogindata() 
{
	if(LoginForm.MemName.value =="" ) {
	alert("�������û�����")
	LoginForm.MemName.focus()
	return false;
	}
		if(LoginForm.Password.value =="" ) {
		alert("���������룡")
		LoginForm.Password.focus()
		return false;
	}
}
function CheckLoginNamedata() 
{
	if(LoginForm.MemName.value =="" ) 
	{
		alert("�������û�����");
		LoginForm.MemName.focus();
		return false;
	}
}
function checkPay() 
{
	if(PayForm.RealName.value =="" ) {
	alert("�������ջ��ˣ�")
	PayForm.RealName.focus()
	return false;
}
	if(PayForm.Address.value =="" ) {
	alert("�������ջ���ַ��")
	PayForm.Address.focus()
	return false;
}
	if(PayForm.UserTel.value =="" ) {
	alert("��������ϵ�绰��")
	PayForm.UserTel.focus()
	return false;
}
	if(PayForm.City.value =="" ) {
	alert("�������ջ����У�")
	PayForm.City.focus()
	return false;
}
	if(PayForm.Postcode.value =="" ) {
	alert("�������������룡")
	PayForm.Postcode.focus()
	return false;
}
	if(PayForm.UserEmail.value =="" ) {
	alert("����������ʼ���")
	PayForm.UserEmail.focus()
	return false;
}
}
function validateEmail(emailStr){
	var re=/^[\w-]+(\.*[\w-]+)*@([0-9a-z]+(([0-9a-z]*)|([0-9a-z-]*[0-9a-z]))+\.)+[a-z]{2,3}$/i;
	if(re.test(emailStr))
		return true;
	else
		return false;
}