<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Include/user_conn.asp"-->
<%
QN_ID = request.Form("QN_ID")	'调查问卷编号
Stu_ID = request.Form("Stu_ID")		'学号
QNum = request.Form("QNum")		'调查问卷题数
'以下信息为临时代码——————————————————————————————
Msg = "您填写的调查问卷不全，请按照调查问卷的要求进行完整填写，谢谢合作！"
If request.Form("Q1") = "" Then
	Response.write (Msg & " <input type='button' value='返回' onclick='history.go(-1)'>")
	Response.End()
End If

If request.Form("Q2") = "" Then
	Response.write (Msg & " <input type='button' value='返回' onclick='history.go(-1)'>")
	Response.End()
End If

If request.Form("Q3") = "" Then
	Response.write (Msg & " <input type='button' value='返回' onclick='history.go(-1)'>")
	Response.End()
End If

If request.Form("Q4") = "" Then
	Response.write (Msg & " <input type='button' value='返回' onclick='history.go(-1)'>")
	Response.End()
End If

If request.Form("Q5") = "" Then
	Response.write (Msg & " <input type='button' value='返回' onclick='history.go(-1)'>")
	Response.End()
End If

If request.Form("Q6") = "" Then
	Response.write (Msg & " <input type='button' value='返回' onclick='history.go(-1)'>")
	Response.End()
End If

If request.Form("Q7") = "" Then
	Response.write (Msg & " <input type='button' value='返回' onclick='history.go(-1)'>")
	Response.End()
End If
'以上代码临时使用——————————————————————————
Set rs = Server.CreateObject("Adodb.RecordSet")
SqlStr = "select * from A" & QN_ID
rs.Open SqlStr,conn,1,3
rs.addnew
rs("Stu_ID") = Stu_ID
For i = 1 to QNum		
	rs("Q" & i) = request.Form("Q" & i)
Next
rs.update
rs.close

Set rs1 = Server.CreateObject("Adodb.RecordSet")
SqlStr = "select * from YON"
rs1.Open SqlStr,conn,1,3
rs1.addnew
rs1("Stu_ID") = Stu_ID
rs1("YON") = QN_ID
rs1.update
rs1.close
conn.close
Set rs1 = nothing
Set conn = nothing

response.Write("成功提交数据!")
%>