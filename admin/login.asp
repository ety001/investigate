<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
Session("logined") = "no"
If request.Form("user") = "admin" and request.Form("pass") = "wangluobu@tw" Then
	Session("logined") = "yes"
	response.Redirect("index.asp")
Else
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=GB2312" />
<title>鲁东大学调查问卷生成管理系统</title>
</head>

<body>
<center><h1>鲁东大学调查问卷生成管理系统</h1></center>
<div style="width:400px; margin:auto auto;">
<form method="post" action="login.asp">
<table width="400" border="1">
  <tr>
    <th width="114" scope="col">用户名</th>
    <th width="270" scope="col"><input type="text" name="user" width="250" height="18" /></th>
  </tr>
  <tr>
    <th scope="row">密码</th>
    <th><input type="password" name="pass" width="250" height="18" /></th>
  </tr>
  <tr>
    <th scope="row">&nbsp;</th>
    <th><input type="submit" value="登陆" /></th>
  </tr>
</table>
</form>
</div>
</body>
</html>
<%End If%>
