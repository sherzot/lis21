<SCRIPT LANGUAGE="JavaScript" type="text/JavaScript">
<!--

function datecheck(target,errMsg)
{
  if (target.value == ""|target.value == null)
  {
    return true;
  }
  dateArray = target.value.split("/");
  if (dateArray[0] == ""|dateArray[1] == ""|dateArray[2] == "")
  {
    alert(errMsg);
    target.focus();
    return false;
  }
  else if (isNaN(dateArray[0])|isNaN(dateArray[1])|isNaN(dateArray[2]))
  {
    alert(errMsg);
    target.focus();
    return false;
  }

  // 1900 年以前はエラーになります
  if (dateArray[0] < 1900)
  {
    alert(errMsg);
    target.focus();
    return false;
  }
  //うるう年かをチェック
  if ((dateArray[0]%4 == 0) & (dateArray[0]%100 == 0) & (dateArray[0]%400 == 0))
  {
    leapYear = true;
  }
  else if((dateArray[0]%4 == 0) & (dateArray[0]%100 != 0))
  {
    leapYear = true;
  }
  else
  {
    leapYear = false;
  }

  // 月が 1 ～ 12 の間にあるかチェック
  if (dateArray[1] < 1|dateArray[1] > 12)
  {
    alert(errMsg);
    target.focus();
    return false;
  }
  
  if (dateArray[1] == 1|dateArray[1] == 3|dateArray[1] == 5|dateArray[1] == 7|dateArray[1] == 8|dateArray[1] == 10|dateArray[1] == 12)
  {
    if (dateArray[2] < 1|dateArray[2] >31)
    {
      alert(errMsg);
      target.focus();
      return false;
    }
  }
  else if (dateArray[1] == 2)
  {
    if (leapYear)
    {
      if (dateArray[2] < 1|dateArray[2] >29)
      {
        alert(errMsg);
        target.focus();
        return false;
      }
    }
    else
    {
      if (dateArray[2] < 1|dateArray[2] >28)
      {
        alert(errMsg);
        target.focus();
        return false;
      }

    }
  }
  else
  {
    if (dateArray[2] < 1|dateArray[2] >30)
    {
      alert(errMsg);
      target.focus();
      return false;
    }

  }
  if (dateArray[1].length == 1)
  {
    dateArray[1] = '0' + dateArray[1];
  }

  if (dateArray[2].length == 1)
  {
    dateArray[2] = '0' + dateArray[2];
  }
  if (dateArray[0].length > 4)
  {
    alert(errMsg);
    target.focus();
    return false;
  }
  if (dateArray[1].length != 2)
  {
    alert(errMsg);
    target.focus();
    return false;
  }
  if (dateArray[2].length != 2)
  {
    alert(errMsg);
    target.focus();
    return false;
  }

  target.value = dateArray[0] + "/" + dateArray[1] + "/" + dateArray[2]
  return true;

}

function timecheck(target, errMsg)
{
  if (target.value == ""|target.value == null)
  {
    return true;
  }
  timeArray = target.value.split(":");
  if (timeArray[0] == ""|timeArray[1] == "")
  {
    alert(errMsg);
    target.focus();
    return false;
  }
  else if (isNaN(timeArray[0])|isNaN(timeArray[1]))
  {
    alert(errMsg);
    target.focus();
    return false;
  }
  //20030327 TASC Uda Add St
  for (i in timeArray){
    if(i>=2){
      alert(errMsg);
      target.focus();
      return false;
    }
  }
  //20030327 TASC Uda Add Ed

  if (timeArray[0]<0|timeArray[0]>23)
  {
    alert(errMsg);
    target.focus();
    return false;
  }
  if (timeArray[1]<0|timeArray[1]>59)
  {
    alert(errMsg);
    target.focus();
    return false;
  }
  return true;
 
}

function stringCheckWithAlert(target,type,errMsg)
{
  if (stringCheck(target,type))
  {
    return true;
  }
  else
  {
    alert(errMsg);
    target.focus();
    return false;
  }
}

function stringCheck(target,type)
{
  targetString = target.value;
  if (type == 'number')
  {
    singleByteString = "1234567890";
  }
  else if (type == 'decimal')
  {
    singleByteString = "1234567890.";
  }
  else if (type == 'telephone')
  {
    singleByteString = "1234567890-";
  }
  else if (type == 'date')
  {
    singleByteString = "1234567890/";
  }
  else if (type == 'time')
  {
    singleByteString = "1234567890:";
  }
  else
  {
<%  '2003/05/27 TASC Terawaki Del Start
    'singleByteString = "1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz!\"#$%&'()=`|~{+*}<>?_-^\\@[;:],./";
    '2003/05/27 TASC Terawaki Del End %>
<%  '2003/05/27 TASC Terawaki Ins Start %>
    singleByteString = "1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz!\"#$%&'()=`|~\{+*}<>?_-^\\@[;:],./";
<%  '2003/05/27 TASC Terawaki Ins End %>
  }
  for (i = 0;i < targetString.length;i++)
  {
    flag = false

    for (j = 0;j < singleByteString.length;j++)
    {
      if (singleByteString.charAt(j) == (targetString.charAt(i)))
      {
        flag=true;
        break;
      }
    }
    if (!flag)
    {
      return false;
    }
  }
  return true;
}

function mailAddressCheck(target, errMsg)
{
  if (target.value.split == ""|target.value.split == null) return true;
  if (stringCheck(target,"alphabet"))
  {
    addressVar = target.value.split("@");
    if (addressVar[0] == ""|addressVar[0] == null|addressVar[1] == ""|addressVar[1] == null)
    {
      alert(errMsg);
      target.focus();
      return false;
    }
    else
    {
      return true;
    }
  }
  else
  {
    alert(errMsg);
    target.focus();
    return false;
  }
}

function necessaryInputCheck(target,type,errMsg)
{
	errFlag = false;
	if (type=='Text') {
		if (target.value == ''|target.value == null) {
			errFlag = true;
		}
	} else if (type=='Select') {
		if (target.selectedIndex == -1 |target.selectedIndex == 0) {
			errFlag = true;
		}
	} else if (type=='Radio'|type=='Checkbox') {
		errFlag = true;
		for(r = 0; r < target.length; r++) {
			if (target[r].checked) {
				errFlag = false;
			}
		}
	}

	if (errFlag) {
		alert(errMsg);
		// 2002/05/23 TASC Terawaki Mod Start
		if ( type == 'Radio' || type == 'Checkbox' ) {
			if ( r == 0 ) {
				target.focus();
			} else {
				target[0].focus();
			}
		} else {
			target.focus();
		}
		// 2002/05/23 TASC Terawaki Mod End
		return false
	} else {
		return true;
	}
}

function back()
{
     history.back();
}

// スタッフ登録用に改良
function datecheck2(target,errMsg)
{
	if (target.value == ""|target.value == null)
	{
		return true;
	}
	dateArray = target.value.split("/");
	if (dateArray[0] == "")
	{
		if ((dateArray[1] == "")|(String(dateArray[1]) == "undefined"))
		{
			alert(errMsg);
			target.focus();
			return false;
		}
	}

	if (isNaN(dateArray[0]))
	{
		alert(errMsg);
		target.focus();
		return false;
	}
	if (isNaN(dateArray[1]))
	{
		alert(errMsg);
		target.focus();
		return false;
	}

	if ((dateArray[2] == "")|(String(dateArray[2]) == "undefined"))
	{
	}
	else if (isNaN(dateArray[2]))
	{
		if (isNaN(dateArray[2]))
		{
			alert(errMsg);
			target.focus();
			return false;
		}
	}

	if (String(dateArray[2]) == "undefined")
	{
	}
	else
	{
		if (dateArray[2] == "")
		{
			alert(errMsg);
			target.focus();
			return false;
		}
		else if (isNaN(dateArray[2]))
		{
			if (isNaN(dateArray[2]))
			{
				alert(errMsg);
				target.focus();
				return false;
			}
		}
	}


	// 1900 年以前はエラーになります
	if (dateArray[0] < 1900)
	{
		alert(errMsg);
		target.focus();
		return false;
	}

	//うるう年かをチェック
	if ((dateArray[0]%4 == 0) & (dateArray[0]%100 == 0) & (dateArray[0]%400 == 0))
	{
		leapYear = true;
	}
	else if((dateArray[0]%4 == 0) & (dateArray[0]%100 != 0))
	{
		leapYear = true;
	}
	else
	{
		leapYear = false;
	}

	// 月が 1 ～ 12 の間にあるかチェック
	if (dateArray[1] < 1|dateArray[1] > 12)
	{
		alert(errMsg);
		target.focus();
		return false;
	}
  
	if (dateArray[1] == 1|dateArray[1] == 3|dateArray[1] == 5|dateArray[1] == 7|dateArray[1] == 8|dateArray[1] == 10|dateArray[1] == 12)
	{
		if ((dateArray[2] != "")&(String(dateArray[2]) != "undefined"))
		{
			if (dateArray[2] < 1|dateArray[2] >31)
			{
				alert(errMsg);
				target.focus();
				return false;
			}
		}
	}
	else if (dateArray[1] == 2)
	{
		if (leapYear)
		{
			if ((dateArray[2] != "")&(String(dateArray[2]) != "undefined"))
			{
				if (dateArray[2] < 1|dateArray[2] >29)
				{
					alert(errMsg);
					target.focus();
					return false;
				}
			}
		}
		else
		{
			if ((dateArray[2] != "")&(String(dateArray[2]) != "undefined"))
			{
				if (dateArray[2] < 1|dateArray[2] >28)
				{
					alert(errMsg);
					target.focus();
					return false;
				}
			}
		}
	}
	else
	{
		if ((dateArray[2] != "")&(String(dateArray[2]) != "undefined"))
		{
			if (dateArray[2] < 1|dateArray[2] >30)
			{
				alert(errMsg);
				target.focus();
				return false;
			}
		}
	}

	if (dateArray[1].length == 1)
	{
		dateArray[1] = '0' + dateArray[1];
	}

	if (String(dateArray[2])!="undefined")
	{
		if (dateArray[2].length == 1)
		{
			dateArray[2] = '0' + dateArray[2];
		}
	}

	if (dateArray[0].length > 4)
	{
		alert(errMsg);
		target.focus();
		return false;
	}

	if (dateArray[1].length != 2)
	{
		alert(errMsg);
		target.focus();
		return false;
	}

	if ((dateArray[2] != "")&(String(dateArray[2]) != "undefined"))
	{
		if (dateArray[2].length != 2)
		{
			alert(errMsg);
			target.focus();
			return false;
		}
	}

	if (String(dateArray[2])!="undefined")
	{
		target.value = dateArray[0] + "/" + dateArray[1] + "/" + dateArray[2]
		return true;
	}
	else
	{
		target.value = dateArray[0] + "/" + dateArray[1]
		return true;
	}

}

// -->
</SCRIPT>