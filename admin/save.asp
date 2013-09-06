<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="func.asp"-->
<%
If Session("Logined") <> "yes" Then	'检查登陆情况
	response.redirect("login.asp")
Else
%>
<!--#include file="../Include/conn.asp"-->
<%
QNum = CInt(request.Form("num"))	'调查问卷题目的数目
QTitle = request.Form("top")	'调查问卷的题目
QN_ID = request.Form("QN_ID")	'调查问卷的编号
questionnaire = split(trim(request.Form),"&")	'获取表格的信息并以“&”号分隔
Dim subjectTemp(99,10) 	'声明一个能存储100个调查题目，每个题目最多有9个选项的数组（二维值0代表题目的类型，1代表题目的主干描述）
Dim subject(99,10) 	'声明一个能存储100个调查题目，每个题目最多有9个选项的数组（二维值0代表题目的类型，1代表题目的主干描述）
subjectNum = 0		'Subject数组的第一维计数
subjectSelectNum = 2	'Subject数组的第二维计数
upSubject = Ubound(questionnaire)	'来自表格的信息数目
'该循环是超级重点
For i = 1 to QNum
	title = "Q" & i
	jj = 2
	For j = 2 to upSubject		'i用来记录来自form提交的字段数
		QTemp = Split(unescape(questionnaire(j)),"_")
		If QTemp(0) = title Then
			If left(QTemp(1),5) = "title" Then subjectTemp(i-1,1) = right(QTemp(1),len(QTemp(1))-6)
			If right(QTemp(1),6) = "single" Then subjectTemp(i-1,0) = "single"
			If right(QTemp(1),8) = "multiple" Then subjectTemp(i-1,0) = "multiple"
			If right(QTemp(1),8) = "question" Then subjectTemp(i-1,0) = "question"
			If left(QTemp(1),6) = "answer" Then
				subjectTemp(i-1,jj) = right(QTemp(1),len(QTemp(1))-8)
				jj = jj + 1
			End If
		End If
	Next
Next	

'通过下面的循环可以进行编码转换，因为Request可以试想编码的转换
For i = 0 to QNum-1
	subject(i,1) = request.Form("Q" & (i+1) & "_title")
	subject(i,0) = request.Form("Q" & (i+1) & "_type")
	For j = 2 to 10
		If subjectTemp(i,j-1) <> "" Then		
			subject(i,j) = request.Form("Q" & (i+1) & "_answer" & (j-1))
		End If
	Next
Next

'以下开始进行数据库操作
On Error Resume Next

'创建以A加当前调查问卷编号命名的表，用来存储参与调查人员提交的调查问卷的答案信息
tableAttribute = "(id autoincrement primary key, Stu_ID varchar(20)"
For i = 1 to QNum
	tableAttribute = tableAttribute & " ,Q" & i & " varchar(200)"
Next
tableAttribute = tableAttribute & ")"
CreatTable = "create table A" & QN_ID & tableAttribute		'生成创建表的SQL语句
conn.execute(CreatTable)	'生成以A加调查问卷编号命名的表，用来存储参与调查人员提交的调查结果

'创建以Q加当前调查问卷编号命名的表，用来存储调查问卷的信息
tableAttribute = "(id autoincrement primary key, QN_ID varchar(20), QNum int, QType varchar(20),Option_Num int, QTitle varchar(100), Op1 varchar(200), Op2 varchar(200), Op3 varchar(200), Op4 varchar(200), Op5 varchar(200), Op6 varchar(200), Op7 varchar(200), Op8 varchar(200), Op9 varchar(200), Question varchar(200))"
CreatTable = "create table Q" & QN_ID & tableAttribute		'生成创建表的SQL语句
conn.execute(CreatTable)	'生成以Q加调查问卷编号命名的表



If Err.Number = 0 Then
	'更新QN_Title表（该表存储调查问卷的编号，题目，题数）的内容
	Set rsQuestion = Server.CreateObject("ADODB.RecordSet")
	sqlStrQN_Title = "select * from QN_Title"
	rsQuestion.open sqlStrQN_Title,conn,1,3	'打开调查问卷表以提交调查问卷的编号，名称，题数
	rsQuestion.addnew
	rsQuestion("QNum") = QNum
	rsQuestion("Title") = QTitle
	rsQuestion("QN_ID") = QN_ID
	rsQuestion("QVisible") = "1"
	rsQuestion.update		'更新调查问卷表
	rsQuestion.close
	Set rsQuestion = nothing
	
	'打开当前新建的该问卷的表
	Set rs = Server.CreateObject("ADODB.RecordSet")
	sqlStr = "select * from Q" & QN_ID
	rs.open sqlStr,conn,1,3

	For p = 0 to QNum-1
		Select Case subject(p,0)
			Case "single"
				rs.addnew
				ttt = 0
				rs("QN_ID") = QN_ID
				rs("QNum") = p + 1	'题号
				rs("QType") = "single"
				rs("QTitle") = subject(p,1)
				For q1 = 2 to 10
					If subject(p,q1) <> "" Then
						rs("Op" & (q1-1)) = subject(p,q1)
						ttt = ttt + 1
					End If
				Next
				rs("Option_Num") = ttt	'该题的选项数
				rs.update
			Case "multiple"
				rs.addnew
				ttt = 0
				rs("QN_ID") = QN_ID
				rs("QNum") = p + 1	'题号
				rs("QType") = "multiple"
				rs("QTitle") = subject(p,1)
				For q2 = 2 to 10
					If subject(p,q2) <> "" Then
						rs("Op" & (q2-1)) = subject(p,q2)
						ttt = ttt + 1
					End If
				Next
				rs("Option_Num") = ttt
				rs.update
			Case "question"
				rs.addnew
				rs("QN_ID") = QN_ID
				rs("QNum") = p + 1	'题号
				rs("QType") = "question"
				rs("QTitle") = "主观题"
				rs("Question") = subject(p,1)
				rs.update
			Case Else
				response.Write("Error")
		End Select	
	Next
	response.Write("<a href='index.asp' target='_self'>添加成功，点击返回</a>")
Else
	response.Write("程序出错，请联系程序编写人员！")
End If
End If
rs.close
Set rs = nothing
conn.close
Set conn = nothing 
%>