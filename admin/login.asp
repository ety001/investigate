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
<title>³����ѧ�����ʾ����ɹ���ϵͳ</title>
</head>

<body>
<center><h1>³����ѧ�����ʾ����ɹ���ϵͳ</h1></center>
<div style="width:400px; margin:auto auto;">
<form method="post" action="login.asp">
<table width="400" border="1">
  <tr>
    <th width="114" scope="col">�û���</th>
    <th width="270" scope="col"><input type="text" name="user" width="250" height="18" /></th>
  </tr>
  <tr>
    <th scope="row">����</th>
    <th><input type="password" name="pass" width="250" height="18" /></th>
  </tr>
  <tr>
    <th scope="row">&nbsp;</th>
    <th><input type="submit" value="��½" /></th>
  </tr>
</table>
</form>
</div>
</body>
</html>
<%End If%>
