<%
function CheckEmail(emailStr) {
	if (Trim(emailStr) == "") return true;
	emailStr = String(emailStr);
	var emailPat=/^(.+)@(.+)$/
	var specialChars="\\(\\)<>@,;:\\\\\\\"\\.\\[\\]"
	var validChars="\[^\\s" + specialChars + "\]"
	var quotedUser="(\"[^\"]*\")"
	var ipDomainPat=/^\[(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})\]$/
	var atom=validChars + '+'
	var word="(" + atom + "|" + quotedUser + ")"
	var userPat=new RegExp("^" + word + "(\\." + word + ")*$")
	var domainPat=new RegExp("^" + atom + "(\\." + atom +")*$")
	var matchArray=emailStr.match(emailPat)
	if (matchArray==null) return false;
	var user=matchArray[1]
	var domain=matchArray[2]
	if (user.match(userPat)==null) return false;
	var IPArray=domain.match(ipDomainPat)
	if (IPArray!=null) {
		// this is an IP address
		for (var i=1;i<=4;i++){
			if (IPArray[i]>255) return false;
		}
		return true;
	}

	var domainArray=domain.match(domainPat)
	if (domainArray==null) return false;
	var atomPat=new RegExp(atom,"g")
	var domArr=domain.match(atomPat)
	var len=domArr.length
	if (domArr[domArr.length-1].length<2 || domArr[domArr.length-1].length>3) return false;
	if (len<2) return false;
	return true;
}

function FormatPhoneNumber(chvPhone_Type_1,chvPhone1_Arcd,chvPhone1_Num,chvPhone1_Ext,chvPhone_Type_2,chvPhone2_Arcd,chvPhone2_Num,chvPhone2_Ext,chvPhone_Type_3,chvPhone3_Arcd,chvPhone3_Num,chvPhone3_Ext){
	var output = "";
	var temp = "";
	if ((chvPhone1_Num != null) && (Trim(chvPhone1_Num) != "")) {
		temp = Trim(chvPhone_Type_1) + ":(" + Trim(chvPhone1_Arcd) + ") " + LeaveDigits(chvPhone1_Num).substring(0,3)+"-"+LeaveDigits(chvPhone1_Num).substring(3,7);
		if ((chvPhone1_Ext != null) && (Trim(chvPhone1_Ext) != "")) {
			temp = temp + " ext." + Trim(chvPhone1_Ext);
		}
	}
	output = temp;
	temp = "";
	if ((chvPhone2_Num != null) && (Trim(chvPhone2_Num) != "")) {
		temp = Trim(chvPhone_Type_2) + ":(" + Trim(chvPhone2_Arcd) + ")" + LeaveDigits(chvPhone2_Num).substring(0,3)+"-"+LeaveDigits(chvPhone2_Num).substring(3,7);
		if ((chvPhone2_Ext != null) && (Trim(chvPhone2_Ext) != "")) {
			temp = temp + " ext." + Trim(chvPhone2_Ext);
		}
	}
	if (Trim(temp) != "")	output = output + ", " + temp;
	temp = "";
	if ((chvPhone3_Num != null) && (Trim(chvPhone3_Num) != "")) {
		temp = Trim(chvPhone_Type_3) + ":(" + Trim(chvPhone3_Arcd) + ")" + LeaveDigits(chvPhone3_Num).substring(0,3)+"-"+LeaveDigits(chvPhone3_Num).substring(3,7);
		if ((chvPhone3_Ext != null) && (Trim(chvPhone3_Ext) != "")) {
			temp = temp + " ext." + Trim(chvPhone3_Ext);
		}
	}
	if (Trim(temp) != "") output = output + ", " + temp;
	return output;

}

function FormatContact(lastname, firstname, employer, jobtitle){
	var t = "";
	t = SpacePadFormat(lastname.substring(0,15), 15) +SpacePadFormat(firstname.substring(0,15), 15) + SpacePadFormat(employer.substring(0,23), 23) + SpacePadFormat(jobtitle.substring(0,20), 20);
	return t;
}

function FilterQuotes(s){
	if (s == null) return "";
	var t = "";
	for (var index = 0; index < s.length; index++) {
		if ((s.charAt(index)) != "'") {
			t = t + s.charAt(index);
		}
	}
	return t;
}

function SpacePadFormat(s, size){
    var S, L;

    S = new String(Trim(s));
    L = size - S.length;

    if (L > 0) {
		for (var i = 1; i <= L; i++) {
			S = S + "&nbsp;";
		}
	}

    return  S;
}

function ZeroPadFormat(s, size){
    var S, L;

    S = new String(s);
    L = size - S.length;

    if (L > 0) {
		for (var i = 1; i <= L; i++) {
			S = "0" + S;
		}
	}

    return  S;
}

function CurrentDateTime(){
   var d, s = "";
   d = new Date();

   s += d.getHours() + ":";
   s += d.getMinutes() + ":";
   s += d.getSeconds() + " ";
   s += (d.getMonth() + 1) + "/";
   s += d.getDate() + "/";
   s += d.getYear();
   return(s);
}

function CurrentDate(){
   var d, s = "";
   d = new Date();

   s += (d.getMonth() + 1) + "/";
   s += d.getDate() + "/";
   s += d.getYear();
   return(s);
}

function CurrentYear(){
   var d, s = "";
   d = new Date();

   s += d.getYear();
   return(s);
}

function CurrentMonth(){
   var d, s = "";
   d = new Date();

   s += (d.getMonth() + 1);
   return(s);
}

function Trim(tstr){
	var str = String(tstr);
	if (str == null) return "";
	str = str + "";
	var iStart, iEnd;
	var sTrimmed;
	var cChar;
	iEnd = str.length - 1;
	iStart = 0;
	bLoop = true;
	cChar = str.charAt(iStart);
	while ((iStart < iEnd) && ((cChar == "\n") || (cChar == "\r") || (cChar == "\t") || (cChar == " "))){
		iStart ++;
		cChar = str.charAt(iStart);
	}
	cChar = str.charAt(iEnd);
	while ((iEnd >= 0) && ((cChar == "\n") || (cChar == "\r") || (cChar == "\t") || (cChar == " "))){
		iEnd --;
		cChar = str.charAt(iEnd);
	}
	if (iStart < iEnd){
		sTrimmed = str.substring(iStart, iEnd + 1);
	} else {
		sTrimmed = "";
	}
	if (str.length==1) sTrimmed = str;
	return sTrimmed;
}


function FilterDate(d){
	temp = new Date(1900,0,1);
	if (String(d)==temp) {
		return "";
	} else {
		return d;
	}
}

function GetYear(p){
	var d = String(p);
	if ((d=="") || (d == null) || (d==0)){
		return ""
	} else return d.substring(0,4);
}

function GetCycle(p){
	var d = String(p);
	if ((d=="") || (d == null) || (d==0)){
		return ""
	} else return d.substring(4,6);
}

function LeaveDigits(TempInString){
	var TempOutString;
	TempOutString= "";
	for (var i=0; i<TempInString.length; i++) {
		if (((TempInString.charAt(i)>="0") && (TempInString.charAt(i)<="9")) || (TempInString.charAt(i)==".")) TempOutString= TempOutString+TempInString.charAt(i);
	}
	return TempOutString;
}

function FormatCurrency(num) {
	if (num==null) num = 0;
	num = String(num).replace(/\$|\,/g,'');
	if (isNaN(num)) num = "0";
	sign = (num == (num = Math.abs(num)));
	num = Math.floor(num*100+0.50000000001);
	cents = num%100;
	num = Math.floor(num/100).toString();
	if(cents<10) cents = "0" + cents;
	for (var i = 0; i < Math.floor((num.length-(1+i))/3); i++)
		num = num.substring(0,num.length-(4*i+3))+','+ num.substring(num.length-(4*i+3));
	return (((sign)?'':'-') + '$' + num + '.' + cents);
}

function FormatSIN(sin) {
	sin = LeaveDigits(Trim(sin));
	if (sin == "") return "";
	if (sin == null) return "";
	if (sin.length != 9) return sin;
	return sin.substring(0,3) + "-" + sin.substring(3,6) + "-" +sin.substring(6,9);
}

function FormatPostalCode(pc) {
	if (pc == null) return "";
	pc = Trim(pc);
	if (pc.length == 6) {
		return pc.substring(0,3) + " " + pc.substring(3,6);
	} else {
		return pc.toUpperCase();
	}
}

function FormatPhoneNumberOnly(phone){
	phone = LeaveDigits(Trim(phone));
	if (phone == "") return "";
	if (phone == null) return "";
	if (phone.length != 7) return phone;
	return phone.substring(0,3) + "-" + phone.substring(3,7);
}

function Truncate(str, size) {
	str = Trim(str);
	if (str.length <= size) return str;
	return str.substring(0,size) + " ...";
}

function ReplaceTags(str) {
	var temp = String(str).replace(/<ul>/g,"");
	temp = temp.replace("</ul>","\n");
	temp = temp.replace(/<p>/g,"\n");
	temp = temp.replace("</p>","\n");
	temp = temp.replace(/<li>/g,"\n-");
	temp = temp.replace(/<br>/g,"\n");
	return temp;
}
%>