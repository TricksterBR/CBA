<!-- #Include virtual="/inc/conn.asp" -->
<%
Response.Buffer = True
Response.ExpiresAbsolute = Now() - 1
Response.Expires = -1
Response.CacheControl = "no-cache"

Dim PostID, PostKEY, PostWORLD, PostBox1, PostBox2
Dim Account_Code, Account_Price
Dim Card_Price, Charde_Msg
Dim CheckStr, Error_Msg

PostID = GetPostField("id")
PostKEY = GetPostField("key")
PostWORLD = GetPostInt("world")
PostBox1 = GetPostField("TextBox1")
PostBox2 = GetPostField("TextBox2")

'PostID = "ntreev"
'PostKEY = "965eb72c92a549dd"
'PostWORLD = 1

If PostID = "" or PostKEY = "" or PostWORLD = 0 Then
	response.Write("FAIL46")
	response.End()
End If
If len(PostID) < 4 or len(PostID) > 12 Then
	response.Write("FAIL200")
	response.End()
End If
If len(PostKEY) <> 16 Then
	response.Write("FAIL210")
	response.End()
End If
If PostWORLD < 0 or PostWORLD > 10 Then
	response.Write("FAIL220")
	response.End()
End If
If PostBox1<>"" And len(PostBox1) <> 9 Then
	response.Write("FAIL300")
	response.End()
End If
If PostBox2<>"" And len(PostBox2) <> 16 Then
	response.Write("FAIL310")
	response.End()
End If

CheckStr = 0
If ChkInvaildWord(PostID) = True Then CheckStr = 1
If ChkInvaildWord(PostKEY) = True Then CheckStr = 1
If ChkInvaildWord(PostWORLD) = True Then CheckStr = 1
If ChkInvaildWord(PostBox1) = True Then CheckStr = 1
If ChkInvaildWord(PostBox2) = True Then CheckStr = 1

If CheckStr = 1 Then
	response.Write("FAIL230")
	response.End()
Else
	Set rsConn = server.createobject("ADODB.Recordset")
	rsSQL = "DECLARE @return_value int;EXEC @return_value = [dbo].[uspn_get_user_authority_check] '"& PostID &"','"& PostKEY &"','"& getip() &"';SELECT 'Return_Value' = @return_value"
	'response.Write(rsSQL & vbCrLf)
	rsConn.open rsSQL,Conn,3,2
	Error_Msg = rsConn("Return_Value")
	set rsSQL = NoThing
	rsConn.close
	Select Case Error_Msg
		Case 0
			Set rsConn = server.createobject("ADODB.Recordset")
			rsSQL = "SELECT ACCOUNT_CODE FROM [dbo].[tbl_account] WHERE ACCOUNT_GID = '"& PostID &"'"
			'response.Write(rsSQL & vbCrLf)
			rsConn.open rsSQL,Conn,1,3
			Account_Code = rsConn("ACCOUNT_CODE")
			set rsSQL = NoThing
			rsConn.close

			If PostBox1 <> "" And PostBox2 <> "" Then
				Set rsConn = server.createobject("ADODB.Recordset")
				rsSQL = "SELECT CARD_PRICE FROM [dbo].[_stt_sell_webCard] WHERE CARD_USE = 0 AND CARD_CODE = '"& PostBox1 &"' AND CARD_PASS = '"& PostBox2 &"'"
				'response.Write(rsSQL & vbCrLf)
				rsConn.open rsSQL,Conn,1,3
				If rsConn.EOF Then
					Card_Price = 0
					Charde_Msg = "儲值卡號或密碼錯誤, 儲值點數失敗!!"
				Else
					Card_Price = rsConn("CARD_PRICE")
				End If
				set rsSQL = NoThing
				rsConn.close
				

				If Card_Price <> 0 Then
					Set rsConn = server.createobject("ADODB.Recordset")
					rsSQL = "SELECT PRICE FROM [dbo].[_stt_sell_webshop] WHERE ACCOUNT_CODE = "& Account_Code &""
					'response.Write(rsSQL & vbCrLf)
					rsConn.open rsSQL,Conn,1,3
					If rsConn.EOF Then
						Conn.execute("INSERT INTO [dbo].[_stt_sell_webshop] ([ACCOUNT_CODE],[PRICE]) VALUES ("& Account_Code &","& Card_Price &")")
					Else
						Conn.execute("UPDATE [dbo].[_stt_sell_webshop] SET [PRICE] = [PRICE] + "& Card_Price &" WHERE ACCOUNT_CODE = "& Account_Code &"")
					End If
					Conn.execute("UPDATE [dbo].[_stt_sell_webCard] SET [CARD_USE] = 1 WHERE CARD_CODE = '"& PostBox1 &"'")
					Conn.execute("INSERT INTO [dbo].[_stt_sell_webCard_log] ([ACCOUNT_CODE],[CARD_CODE],[CARD_PASS],[CARD_PRICE]) VALUES ("& Account_Code &",'"& PostBox1 &"','"& PostBox2 &"',"& Card_Price &")")
					set rsSQL = NoThing
					rsConn.close
					Charde_Msg = "儲值點數 ["& Card_Price &"] 成功!!"
				End If
			End If

			Set rsConn = server.createobject("ADODB.Recordset")
			rsSQL = "SELECT PRICE FROM [dbo].[_stt_sell_webshop] WHERE ACCOUNT_CODE = "& Account_Code &""
			'response.Write(rsSQL & vbCrLf)
			rsConn.open rsSQL,Conn,1,3
			If rsConn.EOF Then
				Account_Price = 0
			Else
				Account_Price = rsConn("PRICE")
			End If
			set rsSQL = NoThing
			rsConn.close

			'response.Write("itemclient_charge_cli.asp" & vbCrLf)
			'response.Write(Request.Form() & vbCrLf)
		Case 103
			response.Write("FAIL40")
			response.End()
		Case Else
			response.Write("FAIL" & Error_Msg)
			response.End()
	End Select
	If Error_Msg <> 0 Then
		'response.Write(Error_Msg & vbCrLf)
	End If
End If
%>
<%If Error_Msg = 0 Then%><!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<HTML>
	<HEAD>
		<title>卡巴拉島-- 線上儲值</title>
		<meta http-equiv="Content-Type" content="text/html; charset=big5">
		<LINK href="font.css" type="text/css" rel="stylesheet">
	</HEAD>
	<body style="BACKGROUND-COLOR: #418cdd" leftmargin="0" marginwidth="0" topmargin="0" marginheight="0" scroll="no" oncontextmenu="window.event.returnValue=false">
	<form name="form1" method="post" action="itemclient_charge_cli.asp" id="form1">
			<div>
				<input type="hidden" name="id" id="id" value="<%=PostID%>" />
				<input type="hidden" name="key" id="key" value="<%=PostKEY%>" />
				<input type="hidden" name="world" id="world" value="<%=PostWORLD%>" />
			</div>
		<table class="text" height="472" cellSpacing="0" cellPadding="2" width="314" align="center" bgColor="#ffffff" border="0">
				<tr>
					<td class="title" bgColor="#1b5596"><IMG height="20" src="images/start.gif" width="20" align="absMiddle"><span class="title-p">線上儲值</span></td>
				</tr>
				<tr>
					<td class="title" bgColor="#000066">
						<div align="center"><span id="Label1"><%=PostID%></span>你好</div>
					</td>
				</tr>
				<tr>
					<td vAlign="top">
						<table class="text" cellSpacing="1" cellPadding="3" width="285" align="center" border="0">
							<tr>
								<td>目前點數有<span class="text2">&nbsp;<span id="Label2"><%=Account_Price%></span>&nbsp;</span>點</td>
							</tr>
							<tr>
								<td>
									<div align="center"></div>
								</td>
							</tr>
							<tr>
								<td>
									<br>
									<font color='red'><b><%=Charde_Msg%></b></font>
									<br>
									<p>請輸入儲值卡號密碼</p>
									<P>儲值卡號：
										<input name="TextBox1" type="text" maxlength="9" id="TextBox1" style="width:120px;" />
										(限用大寫)<br>
										儲值密碼：
										<input name="TextBox2" type="password" maxlength="16" id="TextBox2" style="width:120px;" />
										(注意大小寫)<br>
										<br>
										<span id="Label4" style="display:inline-block;color:Red;font-weight:bold;width:264px;">*My Card儲值系統固定於每星期一早上08:00~09:30維護</span>
									<P></P>
									<div align="center"><span id="Label3" style="color:Red;font-size:Small;font-weight:bold;"></span></div>
								</td>
							</tr>
							<tr>
								<td>
									<div align="center"><input type="submit" name="Button1" value="確定" onclick="this.disabled = true;this.value='資料處理中,請稍候';form1.Button3.style.display='none';form1.submit();" id="Button1" class="border3" />&nbsp;<INPUT class="border3" id="Button3" onclick="location.href='itemclient_close.html'" type="button" value="關閉" name="Button3"></div>
									<DIV align="center">&nbsp;</DIV>
									<DIV align="center"><span id="Label5"></span></DIV>
								</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</form>
	</body>
</HTML><%End If%>
<!-- #Include virtual="/inc/close.asp" -->