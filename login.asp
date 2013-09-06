<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Include/user_conn.asp"-->
<%
t = ""
tt = ""
t1 = request.Form("Stu_ID")
t2 = request.Form("QIdentity")
If t1 = "" or t2 = "" Then
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=GB2312" />
<meta name="author" content="ETY001(http://hi.baidu.com/ety001)" />
<title>鲁东大学调查问卷系统 Power By ETY001</title>
<style type="text/css">
<!--
.login{
	width:370px;margin:20px auto 0; height:730px; background-color:#FFF;
}
.login_top{
	background:url(image/login_top.jpg) no-repeat; height:60px;
}
.login_title{
	background:url(image/login.gif) no-repeat;height:57px; position:relative; top:15px; left:-19px;
}
.text{
	padding:30px 10px 0; text-align:center; font-family:'黑体'; font-size:20px; font-weight:bold;
}
.text div{
	width:150px;
	float:left;
	text-align:right;
	position:relative;
	left:-10px;
}
.text input{
	display:block;
	float:left;
	border:#ACACB4 solid 1px;
	height:20px;
	width:180px;
	font-family:Arial, Helvetica, sans-serif;
	font-size:16px;
}
.clear{
	clear:both;
	height:0;
}
.submit{
	margin:55px auto;
	text-align:center;
}
.submit input{
	background:url(image/button2.gif) no-repeat;
	width:84px;
	height:34px;
}
.login_info{
	padding:10px 15px;
	background-color:#dbeaf7;
	text-indent:2em;
	font-size:16px;
}
.login_info p{
	margin:5px 0;
	font-weight:bold;
	line-height:20px;
}
.login_info .teshu{
	text-indent:0em;
}
.login_info .teshuTitle{
	display:block;
	position:relative;
	left:-70px;
	width:420px;
	text-indent:0em;
	font-family:"黑体";
	font-size:30px;
	font-weight:bolder;
	color:#F00;
	background-color:#0CF;
	text-align:center;
	line-height:40px;
}
.login_info .teshu1{
	text-indent:0em;
	font-size:12px;
	color:#396;
}

-->
</style>
</head>

<body style="margin:0 auto;background:url(image/bg.jpg) repeat-x #4494B4;">
<div class="login"><form action="login.asp" method="post">
    <div class="login_info">
   	  <p class="teshuTitle">请2010级本专科同学在选课前首先填写调查问卷</p>
    	<p class="teshu">亲爱的新同学：</p>
      <p class="teshu">你好！</p>
        <p>祝贺你顺利通过高考，成为鲁东大学大家庭的一员！为了进一步做好我校2011年招生宣传工作，提高我校生源质量，结合2010级新生入学资格复查工作，请各位新同学登录 <a href="http://www.lduzs.com/wj" target="_blank">http://www.lduzs.com/wj</a> ,认真、如实填写《鲁东大学2010年招生宣传调查问卷》，谢谢。 </p>
      <p class="teshu1">注意：</p>
      <p class="teshu1">1、2010级普通本专科学生必须填写，专升本、五年一贯制学生无需填写！</p>
      <p class="teshu1">2、时间：9月13日-19日（新生选课的正选阶段）。</p>
      <p class="teshu1">3、请输入本人学号和身份证号登录调查问卷系统。</p>
    </div>
    <div class="login_top"></div>
   	<div class="login_title"></div>
    <div class="text"><label><div>学号：</div><input type="text" name="Stu_ID" value="" /></label></div>
    <div class="clear"></div>
  	<div class="text"><label><div>身份证号：</div><input type="password" name="QIdentity" value="" /></label></div>
    <div class="submit"><input type="submit" onmouseover="this.style.background = 'url(image/button1.gif)'" onmouseout="this.style.background = 'url(image/button2.gif)'" value="" /></div>
    </form>
</div>
<div style="text-align:center; display:none;"><script src="http://s17.cnzz.com/stat.php?id=2438628&web_id=2438628&show=pic" language="JavaScript"></script>
</div>
</body>
</html>
<%
Else
	For i = 1 to len(t1)
		t = t & CStr(CInt(mid(t1,i,1)))
	Next
	Stu_ID = trim(t)

	For i = 1 to len(t2)
		tt = tt & mid(t2,i,1)
	Next
	QIdentity = trim(UCase(tt))

	Set ChkLog = conn.execute("select * from QUSER where Stu_ID='" & Stu_ID & "'")

	If ChkLog("QIdentity") = QIdentity and ChkLog("QIdentity") <> "" Then
		Session("User_Logined") = "yes"
		Session("Stu_ID") = Stu_ID
		Session("Stu_name") = ChkLog("name")
		response.Redirect("user.asp")
	Else
		Session("User_Logined") = "no"
		Session("Stu_ID") = ""
		response.Write("<script language='javascript'>alert('学号与身份证不匹配！');history.go(-1)</script>")
	End If
	ChkLog.close
	Set ChkLog = nothing
End If
%>