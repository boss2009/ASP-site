function DayOfTheWeek(d){
	if (!CheckDate(d)) {
		return "Invalid Date";
	}
	var x = new Array("Sunday", "Monday", "Tuesday");
	x = x.concat("Wednesday","Thursday", "Friday");
	x = x.concat("Saturday")
	var newd = new Date(d);
	return x[newd.getUTCDay()];
}

function ForwardWeek(i){
	var s = "";
	dtWeek = new Date();
	dtWeek.setDate( dtWeek.getDate()+7*i);
	s += (dtWeek.getMonth() + 1) + "/";
	s += dtWeek.getDate() + "/";
	s += dtWeek.getYear();
	return(s);
}

function ForwardDay(i){
	var s = "";
	dtDay = new Date();
	dtDay.setDate( dtDay.getDate()+Number(i));
	s += (dtDay.getMonth() + 1) + "/";
	s += dtDay.getDate() + "/";
	s += dtDay.getYear();
	return(s);
}

function ForwardMonth(i){
	var s = "";
	currentDate = new Date();
	currentDate.setMonth(currentDate.getMonth() + Number(i));
	s += (currentDate.getMonth() + 1) + "/";
	s += currentDate.getDate() + "/";
	s += currentDate.getYear();
	return(s);
}

function ForwardYear(i){
	var s = "";
	currentDate = new Date();
	currentDate.setMonth(currentDate.getMonth() + 12*i );
	s += (currentDate.getMonth() + 1) + "/";
	s += currentDate.getDate() + "/";
	s += currentDate.getYear();
	return(s);
}

function CheckTextArea(obj, size){
	if (obj.value.length > size) {
		return false;
	} else {
		return true;
	}
}

function CheckBox(obj){
	obj.checked = true;
}

function UncheckBox(obj){
	obj.checked = false;
}

function ToDec(num){
	if ((num >= 0) && (num <= 9)) return num;
	num = num.toUpperCase();
	switch (num) {
		case "A":
			return 10;
		break;
		case "B":
			return 11;
		break;
		case "C":
			return 12;
		break;
		case "D":
			return 13;
		break;
		case "E":
			return 14;
		break;
		case "F":
			return 15;
		break;
		default:
			return 0;
		break;
	}
}

function ToHex(num){
	if ((num >= 0) && (num <= 9)) return num;
	switch (String(num)) {
		case "10":
			return "A";
		break;
		case "11":
			return "B";
		break;
		case "12":
			return "C";
		break;
		case "13":
			return "D";
		break;
		case "14":
			return "E";
		break;
		case "15":
			return "F";
		break;
		default:
			return "0";
		break;
	}
}

function DecToHex(dec){
	var quotient = 0;
	var remainder = 0;
	if (dec > 15) {
		remainder = (dec % 16);
		quotient = Math.floor(dec / 16);
		return DecToHex(quotient) + String(ToHex(remainder));
	} else {
		return String(ToHex(dec));
	}
}

function PadDecToHex(dec){
	var result = "0"
	result = String(DecToHex(dec));
	if ((result.length % 2) == 1) result = "0" + result;
	return result;
}

function HexToDec(hex){
	var sum = 0;
	for (var i = 0; i < hex.length; i++){
		sum = sum + ToDec(hex.substring(hex.length-i-1,hex.length-i))*(Math.pow(16,i));
	}
}

function IsString(str){
	var characters = "abcdefghijklmnopqrstuvwxyz.,-0123456789 ";
	for (var i=0; i<str.length; i++){
		if (characters.indexOf(str.charAt(i).toLowerCase()) == -1) return false;
	}
	return true;
}

function IsID(num){
	num = Trim(num);
	var numbers = "0123456789";
	for (var i=0; i<num.length; i++){
		if (numbers.indexOf(num.charAt(i).toLowerCase()) == -1)	return false;
	}
	return true;
}

function CheckSIN(sin) {
	sin = LeaveDigits(Trim(sin))
	if (sin.length == 0) return true;
	if (sin.length != 9) return false;
	return true;
}

function CheckDate(date){
	if (Trim(date)=="") return true;
	if (date.indexOf(" ")!=-1) return false;
	var splitdate;
	splitdate = date.split("/");
	if (splitdate.length != 3) return false;
	if ((splitdate[0]<01) || (splitdate[0]>12)) return false;
	if ((splitdate[1]<01) || (splitdate[1]>31)) return false;
	if ((splitdate[2]<1900) || (splitdate[2]>2500)) return false;
	if ((splitdate[0] == 2) || (splitdate[0] == 4) || (splitdate[0] == 6) || (splitdate[0] == 9) || (splitdate[0] == 11)) {
		if (splitdate[1] > 30) return false;
	}
	return true;
}

function CheckDateBetween(range){
	var splitdates;
	splitdates = range.split(" ");
	if (splitdates.length != 3) return false;
	if (!(splitdates[1].toLowerCase() == "and")) {
		alert("Use 'and' between dates.");
		return false;
	}

	if (!CheckDate(splitdates[0])) {
		alert("Invalid start date.  Use (mm/dd/yyyy).");
		return false;
	}
	if (!CheckDate(splitdates[2])) {
		alert("Invalid end date.  Use (mm/dd/yyyy).");
		return false;
	}
	//Check Date 2 > Date 1
	var start, end;
	start = splitdates[0].split("/");
	end = splitdates[2].split("/");
	if (end[2] < start[2]) {
//			alert(end[2]);
//			alert(start[2]);
		alert("Invalid date range (Year).");
		return false;
	}
	if ((end[2] == start[2]) && (end[1] < start[1])) {
		alert("Invalid date range (Month).");
		return false;
	}
	if ((end[2] == start[2]) && (end[1] < start[1]) && (end[0] < start[0])) {
		alert("Invalid date range (Day).");
		return false;
	}
	return true;
}

function CheckPostalCode(str){
	str = Trim(str);
	str = RemoveSpace(str);
	if ((str.length == 5) && (!isNaN(str))) return true;
	if (str.length == 0) return true;
	if (str.length != 6) return false;
	if ((str.charAt(0).toUpperCase() < 65) || (str.charAt(0).toUpperCase() > 90)) return false;
	if ((str.charAt(1) < 0 ) || (str.charAt(1).toUpperCase() > 9)) return false;
	if ((str.charAt(2).toUpperCase() < 65) || (str.charAt(2).toUpperCase() > 90)) return false;
	if ((str.charAt(3) < 0 ) || (str.charAt(3).toUpperCase() > 9)) return false;
	if ((str.charAt(4).toUpperCase() < 65) || (str.charAt(4).toUpperCase() > 90)) return false;
	if ((str.charAt(5) < 0 ) || (str.charAt(5).toUpperCase() > 9)) return false;
	return true;
}

function Trim(str){
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
	}
	else{
		sTrimmed = "";
	}
	if (str.length==1) sTrimmed = str;
	return sTrimmed;
}

function RemoveSpace(str){
	if (str == "") return "";
	var str2 = "";
	for (var i = 0; i<=str.length-1; i++){
		if (str.charAt(i) != " ") str2 = str2 + str.charAt(i);
	}
	return str2;
}

function CheckEmail(emailStr) {
	if (Trim(emailStr) == "") return true;
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

function ValidPEN(num){
	if (num.length == 0) return true;
	if ((num.length > 0) && (num.length < 9)) return false;
	if (num.length > 9) return false;
	var numbers = "0123456789";
	for (var i=0; i<num.length; i++){
		if (numbers.indexOf(num.charAt(i).toLowerCase()) == -1) return false;
	}
	return true;
}

function LeaveDigits(TempInString){
	var TempOutString;
	TempOutString= "";
	for (var i=0; i<TempInString.length; i++) {
		if ((TempInString.charAt(i)>="0") && (TempInString.charAt(i)<="9")) TempOutString= TempOutString+TempInString.charAt(i);
	}
	return TempOutString;
}

function AllowNumericOnly(){
	if (((window.event.keyCode<48) || (window.event.keyCode>57)) && (window.event.keyCode != 46)) window.event.keyCode=0;
}

function FormatCurrency(num) {
	num = num.toString().replace(/\$|\,/g,'');
	if (isNaN(num)) num = "0";
	sign = (num == (num = Math.abs(num)));
	num = Math.floor(num*100+0.50000000001);
	cents = num%100;
	num = Math.floor(num/100).toString();
	if (cents<10) cents = "0" + cents;
	for (var i = 0; i < Math.floor((num.length-(1+i))/3); i++) {
		num = num.substring(0,num.length-(4*i+3)) + ','+ num.substring(num.length-(4*i+3));
	}
	return (((sign)?'':'-') + '$' + num + '.' + cents);
}

function FormatSIN(obj) {
	sin = LeaveDigits(Trim(obj.value));
	var output = "";
	if ((sin == "") || (sin == null)) output = "";
	if (sin.length != 9) {
		output = sin;
	} else {
		output = sin.substring(0,3) + "-" + sin.substring(3,6) + "-" + sin.substring(6,9);
	}
	obj.value = output;
	return ;
}

function FormatPostalCode(obj) {
	var output = "";
	if (obj == null) output = "";
	pc = Trim(obj.value);
	if (pc.length == 6) {
		output = pc.substring(0,3) + " " + pc.substring(3,6);
	} else {
		output = pc;
	}
	obj.value = output.toUpperCase();
	return ;
}

function FormatPhoneNumberOnly(obj){
	phone = LeaveDigits(Trim(obj.value));
	var output = "";
	if ((phone == "") || (phone == null)) output = "";
	if (phone.length != 7) {
		output = phone;
	} else {
		output = phone.substring(0,3) + "-" + phone.substring(3,7);
	}
	obj.value = output;
	return ;
}

function FormatDate(obj){
	d = LeaveDigits(Trim(obj.value));
	var output = "";
	if ((d == "") || (d == null)) output = "";
	if (d.length != 8) {
		output = Trim(obj.value);
	} else {
		output = d.substring(0,2) + "/" + d.substring(2,4) + "/" + d.substring(4,8);
	}
	obj.value = output;
	return ;
}