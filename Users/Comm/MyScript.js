////判断第1次//////////////////////////Foosun Inc. 轻风云 /////////////////////////////////////////////////////////////////
function checkdata() 
{
	if(UserForm.Username.value =="" ) {
	alert("请输入用户名！")
	UserForm.Username.focus()
	return false;
}
if(UserForm.Username.value.length<3 ) {
	alert("用户名至少3个字符！")
	UserForm.Username.focus()
	return false;
	}
if(UserForm.Username.value.length>18 ) {
	alert("用户名不能超过18个字符！")
	UserForm.Username.focus()
	return false;
	}
if(UserForm.sPassword.value =="" ) {
	alert("请输入密码！")
	UserForm.sPassword.focus()
	return false;
}
if(UserForm.sPassword.value.length<6 ) {
	alert("密码应该大于6小于18个字符！")
	UserForm.sPassword.focus()
	return false;
}
if(UserForm.sPassword.value !==UserForm.Confimpass.value) {
	alert("2次密码不一致！")
	UserForm.sPassword.focus()
	return false;
	}
if(UserForm.email.value =="" ) {
	alert("请输入电子邮件！")
	UserForm.email.focus()
	return false;
}
if(UserForm.email.value.length<8 || UserForm.email.value.length>64 || !validateEmail(UserForm.email.value) ) {
		alert("请您输入正确的邮箱地址!")
		UserForm.email.focus()
		return false;
	}
if(UserForm.PassQuestion.value =="" ) {
	alert("请输入密码问题！")
	UserForm.PassQuestion.focus()
	return false;
}
if(UserForm.PassAnswer.value =="" ) {
	alert("请输入密码答案！")
	UserForm.PassAnswer.focus()
	return false;
}
if(UserForm.PassQuestion.value ==UserForm.PassAnswer.value) {
	alert("密码提示问题和答案不能相同！！")
	UserForm.PassQuestion.focus()
	return false;
	}
}
////判断第2次//////////////////////////Foosun Inc. 轻风云 /////////////////////////////////////////////////////////////////
function checkdata1() 
{
	if(UserForm1.sName.value =="" ) {
	alert("请输入您的真实姓名！")
	UserForm1.sName.focus()
	return false;
}
if(UserForm1.sName.value.length<2 ) {
	alert("请正确填写姓名！")
	UserForm1.sName.focus()
	return false;
}
if( !(UserForm1.Sex[0].checked || UserForm1.Sex[1].checked)) {
alert("请选择性别 !!")
return false;
}
	if(UserForm1.Ver.value =="" ) {
	alert("请输入效验码！")
	UserForm1.Ver.focus()
	return false;
}
if(UserForm1.tel.value =="" ) {
	alert("请输入联系电话号码！")
	UserForm1.tel.focus()
	return false;
}
if(UserForm1.tel.value.length<7 ) {
	alert("请正确填写电话号码！")
	UserForm1.tel.focus()
	return false;
}
if(UserForm1.address.value =="" ) {
	alert("请输入联系地址！")
	UserForm1.address.focus()
	return false;
}
if(UserForm1.address.value.length<5 ) {
	alert("请正确填写地址！")
	UserForm1.address.focus()
	return false;
}
if(UserForm1.email.value =="" ) {
	alert("请输入电子邮件！")
	UserForm1.email.focus()
	return false;
}
if(UserForm1.email.value.length<8 || UserForm1.email.value.length>64 || !validateEmail(UserForm1.email.value) ) {
		alert("请您输入正确的邮箱地址!")
		UserForm1.email.focus()
		return false;
	}
}
function CheckLogindata() 
{
	if(LoginForm.MemName.value =="" ) {
	alert("请输入用户名！")
	LoginForm.MemName.focus()
	return false;
	}
		if(LoginForm.Password.value =="" ) {
		alert("请输入密码！")
		LoginForm.Password.focus()
		return false;
	}
}
function CheckLoginNamedata() 
{
	if(LoginForm.MemName.value =="" ) 
	{
		alert("请输入用户名！");
		LoginForm.MemName.focus();
		return false;
	}
}
function checkPay() 
{
	if(PayForm.RealName.value =="" ) {
	alert("请输入收货人！")
	PayForm.RealName.focus()
	return false;
}
	if(PayForm.Address.value =="" ) {
	alert("请输入收货地址！")
	PayForm.Address.focus()
	return false;
}
	if(PayForm.UserTel.value =="" ) {
	alert("请输入联系电话！")
	PayForm.UserTel.focus()
	return false;
}
	if(PayForm.City.value =="" ) {
	alert("请输入收货城市！")
	PayForm.City.focus()
	return false;
}
	if(PayForm.Postcode.value =="" ) {
	alert("请输入邮政编码！")
	PayForm.Postcode.focus()
	return false;
}
	if(PayForm.UserEmail.value =="" ) {
	alert("请输入电子邮件！")
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