<!-- #Include virtual="/inc/conn.asp" -->
<%
Response.Buffer = True
Response.ExpiresAbsolute = Now() - 1
Response.Expires = -1
Response.CacheControl = "no-cache"

Dim PostID, PostKEY, PostWORLD, PostItem, PostCheck
Dim Account_Code, Game_Code, Account_Price
Dim Goods_Name, Goods_Desc, Goods_Limit_Desc, Goods_Char_Level, Goods_Char_Type, Goods_Limit
Dim Limit_Code, Goods_Limit_Price
Dim CheckStr, Tip_Goods, Tip_Image, Tip_Msg, Buy_Msg, Error_Msg
dim IsHack


IsHack=0
PostID = GetPostField("id")
PostKEY = GetPostField("key")
PostWORLD = GetPostInt("world")
PostItem = GetPostInt("item")
PostCheck = GetPostField("Check")

'PostID = "showme"
'PostKEY = "49ba59abbe56e057"
'PostWORLD = 1
'PostItem = 19966

If PostID = "" or PostKEY = "" or PostWORLD = 0 or PostItem = 0 Then
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
If PostItem < 0 or PostItem > 818415 Then
	response.Write("FAIL240")
	response.End()
End If

CheckStr = 0
If ChkInvaildWord(PostID) = True Then CheckStr = 1
If ChkInvaildWord(PostKEY) = True Then CheckStr = 1
If ChkInvaildWord(PostWORLD) = True Then CheckStr = 1
If ChkInvaildWord(PostItem) = True Then CheckStr = 1

If CheckStr = 1 Then
	response.Write("FAIL230")
	response.End()
Else

	if PostItem=33039 then PostItem=19683
	if PostItem=33038 then PostItem=19725
	if PostItem=33037 then PostItem=19724
	
	Set rsConn = server.createobject("ADODB.Recordset")
	rsSQL = "DECLARE @return_value int;EXEC @return_value = [dbo].[uspn_get_user_authority_check] '"& PostID &"','"& PostKEY &"','"& getip() &"';SELECT 'Return_Value' = @return_value"
	'response.Write(rsSQL & vbCrLf)
	rsConn.open rsSQL,Conn,3,2
	Error_Msg = rsConn("Return_Value")
	set rsSQL = NoThing
	rsConn.close
	Select Case Error_Msg
		Case 0
			Buy_Msg = 100
			If PostCheck = "" Then
				Set rsConn = server.createobject("ADODB.Recordset")
				rsSQL = "SELECT * FROM [dbo].[tbl_goods] WHERE goods_issell=1 and GOODS_CODE = "& PostItem &""
				'response.Write(rsSQL & vbCrLf)
				rsConn.open rsSQL,Conn,1,3
				If rsConn.EOF Then
					Tip_Goods = " disabled=""disabled"""
					Tip_Msg = "<span id=""Label9"" style=""color:Red;font-size:Small;font-weight:bold;"">找不到此物品資料喔!!</span>"
					IsHack=1
				Else
					Goods_Name = rsConn("GOODS_NAME")
					Goods_Desc = rsConn("GOODS_DESC")
					Goods_Limit_Desc = rsConn("GOODS_LIMIT_DESC")
					Goods_Char_Level = rsConn("GOODS_CHAR_LEVEL")
					If Goods_Char_Level = 0 Then
						Goods_Char_Level = "全部使用"
					Else
						Goods_Char_Level = "限定"& Goods_Char_Level &"級以上"
					End If
					Select Case rsConn("GOODS_CHAR_TYPE")
						Case 0 Goods_Char_Type = ""
						Case 1 Goods_Char_Type = ""
						Case 2 Goods_Char_Type = ""
						Case 3 Goods_Char_Type = ""
						Case 4 Goods_Char_Type = ""
						Case 5 Goods_Char_Type = ""
						Case 6 Goods_Char_Type = ""
						Case 8 Goods_Char_Type = ""
						Case 9 Goods_Char_Type = ""
						Case 10 Goods_Char_Type = ""
						Case 11 Goods_Char_Type = ""
						Case 12 Goods_Char_Type = ""
						Case 15 Goods_Char_Type = "消耗品"
						Case Else Goods_Char_Type = "未知品"
					End Select
					Goods_Limit = "<input id=""RadioButton1"" type=""radio"" name=""Check"" value=""RadioButton1"" checked=""checked"" /><label for=""RadioButton1"">"& Goods_Char_Type
				End If
				set rsSQL = NoThing
				rsConn.close
				
				if IsHack=1 then
				'記錄黑客POST記錄
					Set rsConn = server.createobject("ADODB.Recordset")
					rsSQL = "INSERT INTO [gmg_account].[dbo].[_stt_HackTools_Log]([PostID],[PostKEY],[PostItem],[PostFrom],[PostToFrom],[PostMessage]) VALUES ('"& PostID &"','"& PostKEY &"',"& PostItem &",'','','')"
					'response.Write(rsSQL & vbCrLf)
					rsConn.open rsSQL,Conn,1,3
					'set rsSQL = NoThing
					'rsConn.close
					
					'response.End()
				end if

				Set rsConn = server.createobject("ADODB.Recordset")
				rsSQL = "SELECT * FROM [dbo].[tbl_goods_limit] WHERE GOODS_CODE = "& PostItem &""
				'response.Write(rsSQL & vbCrLf)
				rsConn.open rsSQL,Conn,1,3
				If rsConn.EOF Then
					Tip_Goods = " disabled=""disabled"""
					Tip_Image = ""
				Else
					Tip_Image = "images/CMItems/"& PostItem &".gif"
					Goods_Limit_Price = rsConn("GOODS_LIMIT_PRICE")
					Goods_Limit = Goods_Limit &"/"& Goods_Limit_Price &"點</label>"
				End If
				set rsSQL = NoThing
				rsConn.close

				Set rsConn = server.createobject("ADODB.Recordset")
				rsSQL = "SELECT ACCOUNT_CODE FROM [dbo].[tbl_account] WHERE ACCOUNT_GID = '"& PostID &"'"
				'response.Write(rsSQL & vbCrLf)
				rsConn.open rsSQL,Conn,1,3
				Account_Code = rsConn("ACCOUNT_CODE")
				set rsSQL = NoThing
				rsConn.close

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
				

				If Tip_Msg = "" Then
					If Account_Price >= Goods_Limit_Price Then
						Tip_Goods = ""
						Tip_Msg = "您確定購買此商品嗎?"
					Else
						Tip_Goods = " disabled=""disabled"""
						Tip_Msg = "<span id=""Label9"" style=""color:Red;font-size:Small;font-weight:bold;"">目前點數不足喔!!</span>"
					End If
				End If
			Else
				Set rsConn = server.createobject("ADODB.Recordset")
				rsSQL = "SELECT ACCOUNT_CODE FROM [dbo].[tbl_account] WHERE ACCOUNT_GID = '"& PostID &"'"
				'response.Write(rsSQL & vbCrLf)
				rsConn.open rsSQL,Conn,1,3
				Account_Code = rsConn("ACCOUNT_CODE")
				set rsSQL = NoThing
				rsConn.close

				Set rsConn = server.createobject("ADODB.Recordset")
				rsSQL = "SELECT GAME_CODE FROM [dbo].[tbl_account_game] WHERE ACCOUNT_CODE = "& Account_Code &""
				'response.Write(rsSQL & vbCrLf)
				rsConn.open rsSQL,Conn,1,3
				Game_Code = rsConn("GAME_CODE")
				set rsSQL = NoThing
				rsConn.close

				Set rsConn = server.createobject("ADODB.Recordset")
				rsSQL = "SELECT * FROM [dbo].[tbl_goods_limit] WHERE GOODS_CODE = "& PostItem &""
				'response.Write(rsSQL & vbCrLf)
				rsConn.open rsSQL,Conn,1,3
				Limit_Code = rsConn("LIMIT_CODE")
				Goods_Limit_Price = rsConn("GOODS_LIMIT_PRICE")
				set rsSQL = NoThing
				rsConn.close

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

				Set rsConn = server.createobject("ADODB.Recordset")
				rsSQL = "SELECT * FROM [dbo].[tbl_goods_limit] WHERE GOODS_CODE = "& PostItem &""
				'response.Write(rsSQL & vbCrLf)
				rsConn.open rsSQL,Conn,1,3
				If rsConn.EOF Then
					Goods_Limit_Price=999999999
				Else
					Goods_Limit_Price = rsConn("GOODS_LIMIT_PRICE")
				End If
				set rsSQL = NoThing
				rsConn.close
				

				If Goods_Limit_Price > Account_Price Then
					response.Write("FAIL50")
					response.End()
				End If

				Set rsConn = server.createobject("ADODB.Recordset")
				rsSQL = "DECLARE @return_value int;EXEC @return_value = [dbo].[uspn_buy_goods_game] @game_code = "& Game_Code &",@world_code = "& PostWORLD &",@account_gid = N'"& PostID &"',@goods_code = "& PostItem &",@settle_price = "& Goods_Limit_Price &",@sell_store = 0,@limit_code = "& Limit_Code &",@user_cash = "& Account_Price &";SELECT 'Return_Value' = @return_value"
				'response.Write(rsSQL & vbCrLf)
				'response.end()
				rsConn.open rsSQL,Conn,3,2
				Buy_Msg = rsConn("Return_Value")
				set rsSQL = NoThing
				rsConn.close
				Select Case Buy_Msg
					Case 0
						Conn.execute("UPDATE [dbo].[_stt_sell_webshop] SET [PRICE] = [PRICE] - "& Goods_Limit_Price &" WHERE ACCOUNT_CODE = "& Account_Code &"")
						'寫入購買日誌
						Conn.execute("exec  [dbo].[uspn_buy_goods_game_log] @account_code="& Account_Code &", @account_gid= N'"& PostID &"', @goods_code= "& PostItem &", @goods_name = N'"&Goods_Name&"', @buy_price= "& Goods_Limit_Price &", @char_uid=0, @gift_char_uid=0")
						'統計當前消費的總額
						Conn.execute("exec  [dbo].[uspn_buy_goods_game_give_63027] @account_code="& Account_Code &", @account_gid= N'"& PostID &"', @goods_code= "& PostItem &", @goods_name = N'"&Goods_Name&"', @buy_price= "& Goods_Limit_Price &", @char_uid=0, @gift_char_uid=0")
						'查詢是否需要補償
						'dim IsFix,FixAmount
						'IsFix=0
						'FixAmount=0
						'Set rsConn = server.createobject("ADODB.Recordset")
						'rsSQL = "select fix_flag from gmg_account.dbo._stt_sell_webshop where account_code="& Account_Code &""
						'response.Write(rsSQL & vbCrLf)
						'rsConn.open rsSQL,Conn,3,2
						'	IsFix=rsConn("fix_flag")
						'set rsSQL = NoThing
						'rsConn.close
						
						'查詢到後求已經消費總數的商數
						'if IsFix=0 then
						'	Set rsConn = server.createobject("ADODB.Recordset")
						'	rsSQL = "select sum(sell_price) as sell_amount from gmg_account.dbo.tbl_sell where account_code="& Account_Code &""
						'	'response.Write(rsSQL & vbCrLf)
						'	rsConn.open rsSQL,Conn,3,2
						'	FixAmount=rsConn("sell_amount")
						'	set rsSQL = NoThing
						'	rsConn.close
						'end if
						
						'根據個數來更新補償
						'Conn.execute("exec  [dbo].[uspn_buy_goods_game_fix_63027] @account_code="& Account_Code &", @account_gid= N'"& PostID &"', @goods_code= "& PostItem &", @goods_name = N'"&Goods_Name&"', @buy_price= "& Goods_Limit_Price &", @char_uid=0, @gift_count="&int(FixAmount/400)&"")
					Case Else
						response.Write("FAIL" & Error_Msg)
				End Select
			End If
			'response.Write("itemdetail_buy_cli.asp" & vbCrLf)
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
<%If Buy_Msg = 0 Then%><HTML>
	<HEAD>
		<title>卡巴拉島-- 購買物品</title>
		<meta http-equiv="Content-Type" content="text/html; charset=big5">
		<LINK href="font.css" type="text/css" rel="stylesheet">
	</HEAD>
	<body style="BACKGROUND-COLOR: #418cdd" leftmargin="0" marginwidth="0" topmargin="0" marginheight="0" scroll="no" oncontextmenu="window.event.returnValue=false">
		<form name="form1" method="post" action="itemdetail_buy_cli.asp" id="form1">
			<table class="text" height="472" cellSpacing="0" cellPadding="2" width="314" align="center" bgColor="#ffffff" border="0">
				<tr>
					<td class="title" bgColor="#1b5596" colSpan="2"><IMG height="20" src="images/start.gif" width="20" align="absMiddle">購買物品</td>
				</tr>
				<tr>
				  <td vAlign="top" width="38%" rowSpan="4">&nbsp;</td>
					<td class="red" width="62%">&nbsp;</td>
			  </tr>
				<tr>
					<td bgColor="#ffffff">&nbsp;</td>
				</tr>
				<tr>
					<td class="text2" bgColor="#ffffff">&nbsp;</td>
				</tr>
				<tr>
					<td>&nbsp;</td>
				</tr>
				<tr>
					<td colSpan="2" align="center"><span id="Label9" style="color:Red;font-size:Small;font-weight:bold;">恭喜, 物品購買成功!!</span></td>
				</tr>
				<tr>
					<td colSpan="2">
						<p>&nbsp;</p>
					</td>
				</tr>
				<tr>
					<td colSpan="2">
						<p align="center">&nbsp;</p>
					</td>
				</tr>
				<tr>
					<td colSpan="2">
						<div align="center"><span id="Label9" style="color:Red;font-size:Small;font-weight:bold;"></span><br>
							<input class="border3" id="Button3" onClick="location.href='itemclient_close.html'" type="button" value="關閉" name="Button3">
							<br>
						</div>
					</td>
				</tr>
			</table>
		</form>
	</body>
</HTML><%End If%><%If Error_Msg = 0 And Buy_Msg = 100 Then%><HTML>
	<HEAD>
		<title>卡巴拉島-- 購買物品</title>
		<meta http-equiv="Content-Type" content="text/html; charset=big5">
		<LINK href="font.css" type="text/css" rel="stylesheet">
	</HEAD>
	<body style="BACKGROUND-COLOR: #418cdd" leftmargin="0" marginwidth="0" topmargin="0" marginheight="0" scroll="no" oncontextmenu="window.event.returnValue=false">
		<form name="form1" method="post" action="itemdetail_buy_cli.asp" id="form1">
			<div>
				<input type="hidden" name="id" id="id" value="<%=PostID%>" />
				<input type="hidden" name="key" id="key" value="<%=PostKEY%>" />
				<input type="hidden" name="world" id="world" value="<%=PostWORLD%>" />
				<input type="hidden" name="item" id="item" value="<%=PostItem%>" />
			</div>
			<table class="text" height="472" cellSpacing="0" cellPadding="2" width="314" align="center" bgColor="#ffffff" border="0">
				<tr>
					<td class="title" bgColor="#1b5596" colSpan="2"><IMG height="20" src="images/start.gif" width="20" align="absMiddle">購買物品</td>
				</tr>
				<tr>
					<td class="title" bgColor="#000066" colSpan="2">
						<div align="center">物品詳細說明</div>
					</td>
				</tr>
				<tr>
					<td vAlign="top" width="38%" rowSpan="4">
						<table cellSpacing="0" cellPadding="2" align="center" bgColor="#999999" border="0">
							<tr>
								<td width="110"><img id="Image1" src="<%=Tip_Image%>" alt="<%=Goods_Name%>" style="height:109px;width:110px;border-width:0px;" /></td>
							</tr>
						</table>
					</td>
					<td class="red" width="62%"><span id="Label1"><%=Goods_Name%></span></td>
				</tr>
				<tr>
					<td bgColor="#ffffff">職業&nbsp;<span id="Label3"><%=Goods_Limit_Desc%></span></td>
				</tr>
				<tr>
					<td class="text2" bgColor="#ffffff">等級&nbsp;<span id="Label4"><%=Goods_Char_Level%></span></td>
				</tr>
				<tr>
					<td>期限<br><span style="font-size:8pt;"><%=Goods_Limit%></span></td>
				</tr>
				<tr>
					<td colSpan="2">功能說明<br><span id="Label5"><%=Goods_Desc%></span></td>
				</tr>
				<tr>
					<td class="title" bgColor="#000066" colSpan="2">
						<div align="center"><span id="Label17">購買情報</span></div>
					</td>
				</tr>
				<tr>
					<td colSpan="2">
						<p>目前剩餘點數：<span class="text2">&nbsp;<span id="Label7"><%=Account_Price%></span>&nbsp;</span>點</p>
					</td>
				</tr>
				<tr>
					<td colSpan="2">
						<p align="center"><span id="Label2"><%=Tip_Msg%></span></p>
					</td>
				</tr>
				<tr>
					<td colSpan="2">
						<div align="center"><span id="Label9" style="color:Red;font-size:Small;font-weight:bold;"></span><br>
							<input type="submit" name="Button1" value="確定" id="Button1"<%=Tip_Goods%> class="border3" onClick="this.disabled = true;this.value='資料處理中,請稍候';form1.Button3.style.display='none';form1.submit();" /><input class="border3" id="Button3" onClick="location.href='itemclient_close.html'" type="button" value="關閉" name="Button3">
							<br>
						</div>
					</td>
				</tr>
			</table>
		</form>
	</body>
</HTML><%End If%>
<!-- #Include virtual="/inc/close.asp" -->