<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
If Session("Logined") <> "yes" Then
	response.redirect("login.asp")
Else
	If len(Month(Date())) = 1 Then
		rndNumMonth = "0" & Month(Date())
	Else
		rndNumMonth = Month(Date())
	End If
	If len(Day(Date())) = 1 Then
		rndNumDay = "0" & Day(Date())
	Else
		rndNumDay = Day(Date())
	End If
	If len(Hour(Now())) = 1 Then
		rndNumHour = "0" & Hour(Now())
	Else
		rndNumHour = Hour(Now())
	End If
	If len(Minute(Now())) = 1 Then
		rndNumMinute = "0" & Minute(Now())
	Else
		rndNumMinute = Minute(Now())
	End If
	rndNum = Year(Date()) & rndNumMonth & rndNumDay & rndNumHour & rndNumMinute	'调查问卷的编号
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=GB2312" />
<title>添加新的调查问卷</title>
<script language="javascript" src="js.js"></script>
</head>
<body>
<div style="margin:0 auto; width:980px;">
	<span><input type="button" value="添加一个单选题" onclick="single()" /></span>
	<span><input type="button" value="添加一个多选题" onclick="multiple()" /></span>
  <span><input type="button" value="添加一个主观题" onclick="question()" /></span>
    <span><a href="index.asp" target="_self">返回</a></span>
    <span>每个选择题最多有9个选项</span>
</div>
<form method="post" action="save.asp">
<div id="content" style="margin:0 auto; width:980px;">
	<div class="top">调查问卷的标题<input type="text" name="top" value="" />&nbsp;调查问卷的题数<input type="text" name="num" value="" width="10" />&nbsp;调查问卷的编号<input type="text" name="QN_ID" value="<%=rndNum%>" /></div>
    
</div>
<div style="margin:0 auto; width:980px;"><input type="submit" value="提交" /></div>
</form>
</body>
</html>
<%End If%>