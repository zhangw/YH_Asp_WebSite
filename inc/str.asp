<%

Function CheckStringLength(txt)
txt=trim(txt)
x = len(txt)
y = 0
for ii = 1 to x
if asc(mid(txt,ii,1)) < 0 or asc(mid(txt,ii,1)) >255 then '如果是汉字
y = y + 2
else
y = y + 1
end if
next
CheckStringLength = y

End Function

'--************* 截取字符串 ************** 

function InterceptString(txt,length)
txt=trim(txt)
x = len(txt)
y = 0
if x >= 1 then
for ii = 1 to x
if asc(mid(txt,ii,1)) < 0 or asc(mid(txt,ii,1)) >255 then '如果是汉字
y = y + 2
else
y = y + 1
end if
if y >= length then 
txt = left(trim(txt),ii) '字符串限长
exit for
end if
next
InterceptString = txt
else
InterceptString = ""
end if

End Function 
Function delHtml(strHtml)
Dim objRegExp, strOutput
Set objRegExp = New Regexp
objRegExp.IgnoreCase = True
objRegExp.Global = True
objRegExp.Pattern = "(<[a-zA-Z].*?>)|(<[\/][a-zA-Z].*?>)"
strOutput = objRegExp.Replace(strHtml, "")
strOutput = Replace(strOutput, "<", "&lt;")
strOutput = Replace(strOutput, ">", "&gt;") 
delHtml = strOutput
Set objRegExp = Nothing
End Function

function GetImgSrc(str) '取得img 标签内容
    dim tmp
    Set objRegExp = New Regexp
     objRegExp.IgnoreCase = True    '忽略大小写
     objRegExp.Global = false        '全文搜索 !关键!
     objRegExp.Pattern = "<img (.*?)src=(.[^\[^>]*)(.*?)>"
    Set Matches =objRegExp.Execute(str)
    For Each Match in Matches
         tmp=tmp & Match.Value
    Next
     GetImgSrc=getimgs(tmp)
end function

function getimgs(str)'取得
    Set objRegExp1 = New Regexp
     objRegExp1.IgnoreCase = True    '忽略大小写
     objRegExp1.Global = True    '全文搜索
    objRegExp1.Pattern = "src\=.+?\.(gif|jpg|png|bmp)"
    set mm=objRegExp1.Execute(str)
    For Each Match1 in mm
         imgsrc=Match1.Value
        '也许存在不能过滤的字符，确保万一
         imgsrc=replace(imgsrc,"""","")
         imgsrc=replace(imgsrc,"src=","")
         imgsrc=replace(imgsrc,"<","")
         imgsrc=replace(imgsrc,">","")
         imgsrc=replace(imgsrc,"img","")
         imgsrc=replace(imgsrc," ","")
         getimgs=getimgs&imgsrc'把里面的地址串起来备用
    next
end function
%>