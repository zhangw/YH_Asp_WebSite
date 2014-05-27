<%
'*************************************************
'��������gotTopic
'��  �ã����ַ���������һ���������ַ���Ӣ����һ���ַ�
'��  ����str   ----ԭ�ַ���
'       strlen ----��ȡ����
'����ֵ����ȡ����ַ���
'*************************************************
function gotTopic(str,strlen)
	if str="" then
		gotTopic=""
		exit function
	end if
	dim l,t,c, i
	str=replace(replace(replace(replace(str,"&nbsp;"," "),"&quot;",chr(34)),"&gt;",">"),"&lt;","<")
	l=len(str)
	t=0
	for i=1 to l
		c=Abs(Asc(Mid(str,i,1)))
		if c>255 then
			t=t+2
		else
			t=t+1
		end if
		if t>=strlen then
			gotTopic=left(str,i) & "��"
			exit for
		else
			gotTopic=str
		end if
	next
	gotTopic=replace(replace(replace(replace(gotTopic," ","&nbsp;"),chr(34),"&quot;"),">","&gt;"),"<","&lt;")
end function

'��������IsValidEmail
'��  �ã����Email��ַ�Ϸ���
'��  ����email ----Ҫ����Email��ַ
'����ֵ��True  ----Email��ַ�Ϸ�
'       False ----Email��ַ���Ϸ�
'********************************************
function IsValidEmail(email)
	dim names, name, i, c
	IsValidEmail = true
	names = Split(email, "@")
	if UBound(names) <> 1 then
	   IsValidEmail = false
	   exit function
	end if
	for each name in names
		if Len(name) <= 0 then
			IsValidEmail = false
    		exit function
		end if
		for i = 1 to Len(name)
		    c = Lcase(Mid(name, i, 1))
			if InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 and not IsNumeric(c) then
		       IsValidEmail = false
		       exit function
		     end if
	   next
	   if Left(name, 1) = "." or Right(name, 1) = "." then
    	  IsValidEmail = false
	      exit function
	   end if
	next
	if InStr(names(1), ".") <= 0 then
		IsValidEmail = false
	   exit function
	end if
	i = Len(names(1)) - InStrRev(names(1), ".")
	if i <> 2 and i <> 3 then
	   IsValidEmail = false
	   exit function
	end if
	if InStr(email, "..") > 0 then
	   IsValidEmail = false
	end if
end function

'***************************************************
'��������IsObjInstalled
'��  �ã��������Ƿ��Ѿ���װ
'��  ����strClassString ----�����
'����ֵ��True  ----�Ѿ���װ
'       False ----û�а�װ
'***************************************************
Function IsObjInstalled(strClassString)
	On Error Resume Next
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If 0 = Err Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
End Function

'**************************************************
'��������strLength
'��  �ã����ַ������ȡ������������ַ���Ӣ����һ���ַ���
'��  ����str  ----Ҫ�󳤶ȵ��ַ���
'����ֵ���ַ�������
'**************************************************
function strLength(str)
	ON ERROR RESUME NEXT
	dim WINNT_CHINESE
	WINNT_CHINESE    = (len("�й�")=2)
	if WINNT_CHINESE then
        dim l,t,c
        dim i
        l=len(str)
        t=l
        for i=1 to l
        	c=asc(mid(str,i,1))
            if c<0 then c=c+65536
            if c>255 then
                t=t+1
            end if
        next
        strLength=t
    else 
        strLength=len(str)
    end if
    if err.number<>0 then err.clear
end function

'****************************************************
'��������SendMail
'��  �ã���Jmail��������ʼ�
'��  ����ServerAddress  ----��������ַ
'        AddRecipient  ----�����˵�ַ
'        Subject       ----����
'        Body          ----�ż�����
'        Sender        ----�����˵�ַ
'****************************************************
function SendMail(MailServerAddress,AddRecipient,Subject,Body,Sender,MailFrom)
	on error resume next
	Dim JMail
	Set JMail=Server.CreateObject("JMail.SMTPMail")
	if err then
		SendMail= "<br><li>û�а�װJMail���</li>"
		err.clear
		exit function
	end if
	JMail.Logging=True
	JMail.Charset="gb2312"
	JMail.ContentType = "text/html"
	JMail.ServerAddress=MailServerAddress
	JMail.AddRecipient=AddRecipient
	JMail.Subject=Subject
	JMail.Body=MailBody
	JMail.Sender=Sender
	JMail.From = MailFrom
	JMail.Priority=1
	JMail.Execute 
	Set JMail=nothing 
	if err then 
		SendMail=err.description
		err.clear
	else
		SendMail="OK"
	end if
end function
'****************************************************
'��������echo
'��  �ã�������ָ���� ANSI �ַ��������Ӧ���ַ���
'��  ����num
'****************************************************
Function echo(num)
echo=Chr(num)
End Function
'****************************************************
'��������WriteErrMsg
'��  �ã���ʾ������ʾ��Ϣ
'��  ����errmsg
'****************************************************
sub WriteErrMsg(errmsg)
	dim strErr
	strErr=strErr & "<html><head><title>������Ϣ</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbcrlf
	strErr=strErr & "<link href='style.css' rel='stylesheet' type='text/css'></head><body>" & vbcrlf
	strErr=strErr & "<table cellpadding=2 cellspacing=2 border=0 width=400 class='border' align=center>" & vbcrlf
	strErr=strErr & "  <tr align='center'><td height='20' class='title'><strong>������Ϣ</strong></td></tr>" & vbcrlf
	strErr=strErr & "  <tr><td height='100' class='tdbg' valign='top'><b>��������Ŀ���ԭ��</b><br>" & errmsg &"</td></tr>" & vbcrlf
	strErr=strErr & "  <tr align='center'><td class='title'><a href='javascript:history.go(-1)'>&lt;&lt; ������һҳ</a></td></tr>" & vbcrlf
	strErr=strErr & "</table>" & vbcrlf
	strErr=strErr & "</body></html>" & vbcrlf
	response.write strErr
end sub

'**************************************************
'
'				��ʾ��ʾ
'
'*************************************************
sub errview(id,url)
	select case id
		case 1
			call(WriteErrMsg("û�������û���..."))
		case 2
			call(WriteErrMsg("û����������..."))
		case 3
			call(WriteErrMsg("û����������ȷ��..."))
		case 4
			call(WriteErrMsg("û��������֤��..."))
		case 5
			call(WriteErrMsg("��֤����������..."))
		case 6
			call(WriteErrMsg("�û���ֻ���ɴ�Сд��ĸ���ֺ��»������..."))
		case 7
			call(WriteErrMsg("����ֻ���ɴ�Сд��ĸ���ֺ��»������..."))
		case 8
			call(WriteErrMsg("�û������������..."))
		case 9
			call(WriteErrMsg("..."))
		case else
			response.Write "<style type='text/css'>"
			response.Write "<!--"
			response.Write ".STYLE1 {font-size: 36px}"
			response.Write "-->"
			response.Write "</style>"
			response.Write "<div align='center' class='STYLE1'>������ֱ�ӷ��ʱ�ҳ��</div>"
	end select 
end sub
'**************************************************
'
'				���Ƿ��ַ�
'
'*************************************************
function reg(num,id)
	onlyCharNumAnd_ = "[^a-zA-Z0-9_]"			'ƥ��Ǵ�С��ĸ���ֺ��»���
	onlyChar = "[^a-zA-Z]"						'ƥ��Ǵ�С��ĸ
	onlyNum = "[^0-9]"							'ƥ�������
	Set re = New RegExp
	re.Global = True
	re.IgnoreCase = True
		re.MultiLine = True
	select case id
		case 1 									'ƥ��Ǵ�С��ĸ���ֺ��»���
			re.Pattern = onlyCharNumAnd_
			reg = re.Test(num)
		case 2									'ƥ��Ǵ�С��ĸ
			re.Pattern = onlyCharNumAnd_
			reg = re.Test(num)
		case 3									'ƥ�������
			re.Pattern = onlyNum
			reg = re.Test(num)
	end select
end function
%>