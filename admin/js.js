var i = 1;
var s;
var ss;
function single(){
	var singleTarget = document.getElementById("content");
	var singleText = "<div class='question'><div class='QTitile'>��" + i + "����Ŀ����ѡ�⣩��<input type='text' name='Q" + i + "_title' value='' /><input type='hidden' name='Q" + i + "_type' value='single'/></div><div id='QOption" + i + "'><span>��" + i + "��ѡ��1���ݣ�<input type='text' name='Q" + i + "_answer1' value='' /></span></div><input type='button' value='���һ��ѡ��' onclick='addOption(0," + i + ")' /></div>";
	singleText = singleTarget.innerHTML + singleText;
	singleTarget.innerHTML = singleText;
	i++;
	s = 2;
}
function multiple(){
	var multipleTarget = document.getElementById("content");
	var multipleText = "<div class='question'><div class='QTitile'>��" + i + "����Ŀ����ѡ�⣩��<input type='text' name='Q" + i + "_title' value='' /><input type='hidden' name='Q" + i + "_type' value='multiple' /></div><div id='QOption" + i + "'><span>��" + i + "��ѡ��1���ݣ�<input type='text' name='Q" + i + "_answer1' value='' /></span></div><input type='button' value='���һ��ѡ��' onclick='addOption(1," + i + ")' /></div>";
	multipleText = multipleTarget.innerHTML + multipleText;
	multipleTarget.innerHTML = multipleText;
	i++;
    ss = 2;
}
function question(){
	var questionTarget = document.getElementById("content");
	var questionText = "<div class='question'><div class='QTitile'>��" + i + "����Ŀ������⣩��<input type='text' name='Q" + i + "_title' value='' /><input type='hidden' name='Q" + i + "_type' value='question'/></div></div>";
	questionText = questionTarget.innerHTML + questionText;
	questionTarget.innerHTML = questionText;
	i++;
}
function addOption(t,p){
	
	switch(t)
	{
		case 0:
			if(s<10){
			pp = i - 1;
			var optionTarget = document.getElementById("QOption" + p);
			var optionText = "<span>��" + pp + "��ѡ��" + s + "���ݣ�<input type='text' name='Q" + pp + "_answer" + s + "' value='' /></span>";
			optionText = optionTarget.innerHTML + optionText;
			optionTarget.innerHTML = optionText;
			s = s + 1;
			break;
			}
			else{
			alert("ÿ�����ѡ��ܳ���9����");
			break;
			}
		case 1:
			if(ss<10){
			pp = i - 1;
			var optionTarget = document.getElementById("QOption" + p);
			var optionText = "<span>��" + pp + "��ѡ��" + ss + "���ݣ�<input type='text' name='Q" + pp + "_answer" + ss + "' value='' /></span>";
			optionText = optionTarget.innerHTML + optionText;
			optionTarget.innerHTML = optionText;
			ss = ss + 1;
			break;
			}
			else{
			alert("ÿ�����ѡ��ܳ���9����");
			break;
			}
	}
}