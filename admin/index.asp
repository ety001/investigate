<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
If request.QueryString("q") = "quit" Then Session("Logined") = "no"
If Session("Logined") <> "yes" Then
	response.redirect("login.asp")
else
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=GB2312" />
<title>³����ѧ�����ʾ�����ϵͳ</title>
</head>

<body>
<div style="margin:0 auto; width:960px;">
<table width="960" border="1">
  <tr>
    <th width="337" scope="col"><a href="add.asp" target="_self">����µĵ����ʾ�</a></th>
    <th width="408" scope="col"><a href="viewList.asp" target="_self">�鿴���еĵ����ʾ�</a></th>
    <th width="193" scope="col"><a href="index.asp?q=quit">�˳�</a></th>
  </tr>
</table>


</div>
</body>
</html>
<%End If%>