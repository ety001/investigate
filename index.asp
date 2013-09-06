<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
If Session("Stu_ID") = "" and Session("User_Logined") <> "yes" Then
	response.redirect("login.asp")
Else
%>
<!--#include file="Include/user_conn.asp"-->
<%
QN_ID = request.QueryString("QN_ID")
title = request.QueryString("title")
If QN_ID <> "" and title <> "" Then
Set rsChk = conn.execute("select * from YON where Stu_ID='" & Session("Stu_ID") & "'")
Do While Not rsChk.eof
	If rsChk("YON") = QN_ID Then
		response.Write("你已经答过此卷")
		response.End()
	End If
	rsChk.movenext
Loop
Set rs = Server.CreateObject("Adodb.RecordSet")
SqlStr = "select * from Q" & QN_ID & " order by QNum asc"
rs.Open SqlStr,conn,1,1
j = 0
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=GB2312" />
<meta name="author" content="ETY001(http://hi.baidu.com/ety001)" />
<title><%=title%> 调查问卷——鲁东大学调查问卷系统 Power By ETY001</title>
<link href="template/style1/style.css" rel="stylesheet" type="text/css" />
</head>
<body>
<div class="main">
	<div class="top"></div>
	<div class="title"><%=title%></div>
    <div class="info">
    	<p class="teshu">新同学：</p>
        <p>你好！</p>
        <p>祝贺你顺利通过高考，成为鲁东大学大家庭的一员！祝你大学生活充实快乐！为了进一步做好我校2011年招生工作，结合我校2010年新生入学资格审查，请各位新同学认真、如实填写《鲁东大学2010年招生宣传调查问卷》，谢谢！</p>
    </div>
    
    <form action="add.asp" method="post">
    <input type="hidden" name="QN_ID" value="<%=QN_ID%>" />
    <input type="hidden" name="Stu_ID" value="<%=Session("Stu_ID")%>" />
    <%Do While Not rs.eof%>
    	<%
			QType = rs("QType")
			Option_Num = rs("Option_num")
		%>
  <%Select Case Qtype%>
			<%Case "single"%>
            	<div class="question">
                <div class="Qtitle"><%=rs("QNum")%>.<%=rs("QTitle")%></div>
        		<div class="Qanswer">
            	<%For i = 1 to Option_Num %>
            		<label><input name="Q<%=rs("QNum")%>" type="radio" value="<%=chr(64+i)%>" /><%=chr(64+i)%>.<%=rs("Op" & i)%></label><br />
            	<%Next%>
            	</div>
                </div>
                
            <%Case "multiple"%>
       			<div class="question">
                <div class="Qtitle"><%=rs("QNum")%>.<%=rs("QTitle")%></div>
            	<div class="Qanswer">
                <%For i = 1 to Option_Num %>
                	<label><input name="Q<%=rs("QNum")%>" type="checkbox" value="<%=chr(64+i)%>" /><%=chr(64+i)%>.<%=rs("Op" & i)%></label><br />
                <%Next%>
                </div>
                </div>
            <%Case "question"%>
            	<div class="question">
                <div class="Qtitle"><%=rs("QNum")%>.<%=rs("Question")%></div>
                <div class="Qanswer"><label><textarea name="Q<%=rs("QNum")%>" cols="80" rows="3"></textarea></label></div>
                </div>
        <%End Select%>
      <%
		rs.movenext
		j = j + 1
		Loop
		rs.close
		Set rs = nothing
		conn.close
		Set conn = nothing
		%>
        <input type="hidden" name="QNum" value="<%=j%>" />
    <div class="submit"><label><input type="submit" value="提交" /></label><label><input type="reset" value="重置" /></label></div>
    </form>
    <div class="bottom">
    	<hr />
        <div class="bottomContent">版权所有 <a href="http://www.jwc.ldu.edu.cn" target="_blank">鲁东大学教务处</a></div>
        <div class="bottomContent">程序制作 <a href="http://www.renren.com/ety001" target="_blank">于业超ETY001</a></div>
    </div>
</div>
<div style="text-align:center;"><script src="http://s17.cnzz.com/stat.php?id=2438628&web_id=2438628&show=pic" language="JavaScript"></script>
</div>
</body>
</html>
<%
Else
	response.Redirect("user.asp")
End If
End If%>