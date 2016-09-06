<%
Function apath()
    Dim userurl
    If Request.ServerVariables("SERVER_PORT")<>"80" Then
        userurl = "http://"&Request.ServerVariables("SERVER_NAME")& ":" & Request.ServerVariables("SERVER_PORT")& Request.ServerVariables("URL")
    Else
        userurl = "http://"&Request.ServerVariables("SERVER_NAME")& Request.ServerVariables("URL")
	End If
    userurl=left(UserUrl,InstrRev(UserUrl,"/")-1)
    userurl=mid(userurl,InstrRev(UserUrl,"/")+1)
    apath=userurl
End Function

Function currentpath()
    Dim userurl
    If Request.ServerVariables("SERVER_PORT")<>"80" Then
        userurl = "http://"&Request.ServerVariables("SERVER_NAME")& ":" & Request.ServerVariables("SERVER_PORT")& Request.ServerVariables("URL")
    Else
        userurl = "http://"&Request.ServerVariables("SERVER_NAME")& Request.ServerVariables("URL")
    End If
    userurl=mid(userurl,8)
    currentpath=userurl
End Function

Public Function CheckPostUrl()
    Dim Server_v1,Server_v2
    Server_v1 = CStr(Request.ServerVariables("HTTP_REFERER"))
    Server_v2 = CStr(Request.ServerVariables("SERVER_NAME"))
    If Mid(Server_v1,8,Len(Server_v2))<>Server_v2 Then
        alertMsg "不允许跨站提交！","../index.asp"
        Response.End
    End If
End Function

Public Function alertMsg(str,url)
    Dim urlstr
    If url<>"" then urlstr="location.href='"&url&"';"
    If not isNul(str) Then str ="alert('"&str&"');"
    echo("<script>"&str&urlstr&"</script>")
End Function

Public Function echo(str)
    response.write(str)
    response.Flush()
End Function

Public Function die(str)
    If not isNul(str) Then
        echo str
    End If
    Response.End()
End Function

Public Function isNul(str)
    If isnull(str) or str = ""  Then isNul = true Else isNul = False 
End Function

Public Function isNum(str)
    If not isNul(str) Then   isNum=isnumeric(str) else isNum=False
End Function

Function isUrl(str)
    If not isNul(str) Then 
        if left(str,7) = "http://" or left(str,8) = "https://" then isUrl = true else isUrl = False
    Else
        isUrl = False
    End If
End Function
'================================================
'函数名:delHtml
'作  用:过滤javascript及iframe
'================================================
Function deljs(byval strHtml) 
    Dim objRegExp, strOutput
    Set objRegExp = New Regexp		
    objRegExp.IgnoreCase = True 
    objRegExp.Global = True 
    objRegExp.Pattern = "(<(script|iframe).*?>)|(<[\/](script|iframe).*?>)"
    strOutput = objRegExp.Replace(strHtml, "")
    strOutput = Replace(strOutput, "<", "<") 
    strOutput = Replace(strOutput, ">", ">") 
    deljs = strOutput
    Set objRegExp = Nothing
End Function
'================================================
'函数名:delHtml
'作  用:过滤HTML
'================================================
Function delHtml(byval strHtml) 
    Dim objRegExp, strOutput
    Set objRegExp = New Regexp		
    objRegExp.IgnoreCase = True 
    objRegExp.Global = True 
    objRegExp.Pattern = "(<[a-zA-Z].*?>)|(<[\/][a-zA-Z].*?>)"
    strOutput = objRegExp.Replace(strHtml, "")
    strOutput = Replace(strOutput, "<", "<") 
    strOutput = Replace(strOutput, ">", ">") 
    delHtml = strOutput
    Set objRegExp = Nothing
End Function
'================================================
'函数名:isInteger
'作  用:判断数字是否整形
'参  数:para
'================================================
Public Function isInteger(para)
    Dim str
    Dim l,i
    If isNUll(para) then 
        isInteger=false
        exit function
    End if
    str=cstr(para)
    If trim(str)="" then
        isInteger=false
        exit function
    End if
    l=len(str)
    For i=1 to l
        If mid(str,i,1)>"9" or mid(str,i,1)<"0" then
            isInteger=false 
            exit function
        End if
    Next
    isInteger=true
    If err.number<>0 then err.clear
End Function
'================================================
'函数名:GetSystem 
'作  用:获取客户端操作系统 
'返回值:客户端操作系统 
'参  数:无
'================================================
Public Function GetOS()
    Dim info,system
    info=Request.ServerVariables("HTTP_USER_AGENT")
    if Instr(info,"NT 5.1")>0 then
        system="Windows XP"
    elseif Instr(info,"NT 6.0")>0 then
        system="Windows Vista"
    elseif Instr(info,"NT 6.1")>0 then
        system="Windows 7"
    elseif Instr(info,"NT 6.2")>0 then
        system="Windows 8"
    elseif Instr(info,"NT 6.3")>0 then
        system="Windows 8.1"
    elseif Instr(info,"NT 6.4")>0 then
        system="Windows 10"
    elseif Instr(info,"NT 10.0")>0 then
        system="Windows 10"
    elseif Instr(info,"Windows Phone")>0 then
        system="Windows Phone"
    elseif Instr(info,"Android")>0 then
        system="Android"
    elseif Instr(info,"iPhone")>0 then
        system="iPhone"
    elseif Instr(info,"iPad")>0 then
        system="iPad"
    elseif Instr(info,"Ubuntu")>0 then
        system="Ubuntu"
    elseif Instr(info,"Debian")>0 then
        system="Debian"
    elseif instr(info,"unix") or instr(info,"linux") or instr(info,"SunOS") or instr(info,"BSD") then
        system="Linux System"
    elseif Instr(info,"mac")>0 then
        system="Mac OS X"
    else
        system="Other"
    end if
    echo system
End Function
'================================================
'函数名:GetBrowser
'作  用:获取客户端浏览器信息
'返回值:客户端浏览器信息
'参  数:无
'================================================
Public Function GetBrowser()
    Dim info,browser
    info=Request.ServerVariables("HTTP_USER_AGENT")
    if Instr(info,"MSIE 6.0")>0 then
        browser="Internet Explorer 6.0"
    elseif Instr(info,"MSIE 7.0")>0 then
        browser="Internet Explorer 7.0"
    elseif Instr(info,"MSIE 8.0")>0 then
        browser="Internet Explorer 8.0"
    elseif Instr(info,"MSIE 9.0")>0 then
        browser="Internet Explorer 9.0"
    elseif Instr(info,"MSIE 10.0")>0 then
        browser="Internet Explorer 10.0"
    elseif Instr(info,"rv:11.0")>0 then
        browser="Internet Explorer 11.0"
    elseif Instr(info,"Edge")>0 then
        browser="Microsoft Edge"
    elseif Instr(info,"Firefox")>0 then
        browser="火狐浏览器"
    elseif Instr(info,"Opera")>0 then
        browser="Opera浏览器"
    elseif Instr(info,"Maxthon")>0 then
        browser="遨游浏览器"
    elseif Instr(info,"The World")>0 then
        browser="世界之窗浏览器"
    elseif Instr(info,"SE 2.X")>0 then
        browser="搜狗浏览器"
    elseif Instr(info,"360SE")>0 then
        browser="360浏览器"
    elseif Instr(info,"QQBrowser")>0 then
        browser="QQ浏览器"
    elseif Instr(info,"UCBrowser")>0 then
        browser="UC浏览器手机版"
    elseif Instr(info,"Mobile Safari")>0 then
        browser="Chrome手机版"
    elseif Instr(info,"Chrome")>0 then
        browser="Chrome浏览器"
    else
        browser="其它"
    end if
    echo browser
End Function
'================================================
'函数名:GetIP
'作  用:获取客户端IP地址
'参  数:无
'================================================
Public Function GetIP()
dim strIPAddr
    If Request.ServerVariables("HTTP_X_FORWARDED_FOR") = "" OR InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), "unknown") > 0 Then
        strIPAddr = Request.ServerVariables("REMOTE_ADDR")
    ElseIf InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ",") > 0 Then
        strIPAddr = Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), 1, InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ",")-1)
    ElseIf InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ";") > 0 Then
        strIPAddr = Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), 1, InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ";")-1)
    Else
        strIPAddr = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
    End If
    GetIP = Trim(Mid(strIPAddr, 1, 30))
End Function
'================================================
'函数名:CheckEmail
'作  用:邮箱格式检测
'参  数:str ----Email地址
'返回值:true无误，false有误
'================================================
Public Function CheckEmail(email)
    CheckEmail=true
        Dim Rep,pass
        Set Rep = new RegExp
        rep.pattern="([\.a-zA-Z0-9_-]){2,10}@([a-zA-Z0-9_-]){2,10}(\.([a-zA-Z0-9]){2,}){1,4}$"
        pass=rep.Test(email)
        Set Rep=Nothing
        If not pass Then CheckEmail=false
End Function
'================================================
'函数名:FormatTime 
'作  用:时间格式化 
'参  数:DateTime ----要格式化的时间 
'       Format   ----格式的形式 
'================================================
Public Function FormatTime(DateTime,Format)
    Dim temp
    select case Format
    case "1"
        FormatTime=""&year(DateTime)&"年"&month(DateTime)&"月"&day(DateTime)&"日"
    case "2"
        FormatTime=""&month(DateTime)&"月"&day(DateTime)&"日"
    case "3"
        FormatTime=""&year(DateTime)&"/"&month(DateTime)&"/"&day(DateTime)&""
    case "4"
        FormatTime=""&month(DateTime)&"/"&day(DateTime)&""
    case "5"
        FormatTime=""&month(DateTime)&"月"&day(DateTime)&"日"&FormatDateTime(DateTime,4)&""
    case "6"
        temp="周日,周一,周二,周三,周四,周五,周六"
        temp=split(temp,",")
        FormatTime=temp(Weekday(DateTime)-1)
    case "7"
        FormatTime=""&year(DateTime)&"-"&month(DateTime)&"-"&day(DateTime)&" "&FormatDateTime(DateTime,4)&"" 
    case Else
        FormatTime=DateTime
    end select
End Function
'================================================
'过程名:HTMLEncode
'作  用:将HTML代码转化为文本
'参  数:fString--传入代码
'返回值:转化后的文本
'================================================
Function HTMLEncode(fString)
    fString=Trim(fString)
    fString=Replace(fString,CHR(38),"&#38;")	'“&”
    fString=replace(fString,"<","&lt;")
    fString=replace(fString,">","&gt;")
    fString=Replace(fString,"\","&#92;")
    fString=Replace(fString,"--","&#45;&#45;")
    fString=Replace(fString,CHR(9),"&#9;")
    fString=Replace(fString,CHR(10),"<br>")
    fString=Replace(fString,CHR(13),"")
    fString=Replace(fString,CHR(22),"&#22;")
    fString=Replace(fString,CHR(32),"&#32;")
    fString=Replace(fString,CHR(39),"&#39;")	'“'”
    fString=Replace(fString,CHR(59),"&#59;")	'“;”
    fString=ReplaceText(fString,"([&#])([a-z0-9]*)&#59;","$1$2;")

    if IsSqlDataBase=0 then '过滤片假名(日文字符)[\u30A0-\u30FF] by wodig
        fString=escape(fString)
        fString=ReplaceText(fString,"%u30([A-F][0-F])","&#x30$1;")
        fString=unescape(fString)
    end if

    if SiteConfig("BannedText")<>"" then
        filtrate=split(SiteConfig("BannedText"),"|")
        for i = 0 to ubound(filtrate)
            fString=ReplaceText(fString,""&filtrate(i)&"",string(len(filtrate(i)),"*"))
        next
    end if
    HTMLEncode=fString
End Function
'================================================
'函数名:Chklogin
'作  用:检测是否登录
'参  数:无
'================================================
Function Chklogin()
    If Request.Cookies(""&SiteCookie&"")("u_username")<>"" and Request.Cookies(""&SiteCookie&"")("u_password")<>"" Then
        alertMsg "你已登录，请不要重复登录!","admin/manage-profile.asp"
    End If
End Function
'================================================
'函数名:Chkuserrank
'作  用:检测普通用户
'参  数:无
'================================================
Function Chkuserlevel()
    If Request.Cookies(""&SiteCookie&"")("u_level")="" Then
        alertMsg "没有权限","../index.asp"
        Response.End
    End If
End Function
'================================================
'函数名:Chkadmin
'作  用:检测后台登陆权限
'参  数:无
'================================================
Function Chkadmin()
    If Request.Cookies(""&SiteCookie&"")("u_level")<>"0" Then
        alertMsg "没有权限!","../index.asp"
        Response.End
    End If
End Function
'================================================
'函数名:isAllowReg
'作  用:检测是否开放注册
'参  数:无
'================================================
Function isAllowReg()
    If allowreg="1" Then
        echo "<h5 class='text-center'>网站已经关闭注册</h5>"
        Response.End
    End If
End Function
'================================================
'函数名:CheckRegname
'作  用:屏蔽用户名注册词语
'参  数:无
'================================================
Function CheckRegname()
    Dim i,strFilter
    strFilter = Split("'"&BanedWords&"'",",")
    For i = 0 To UBound(strFilter)
        if instr(u_username,strFilter(i)) <> 0 then
            alertMsg "用户名含有敏感词语","javascript:history.back(-1)"
            Response.End
        End If
    Next
End Function
'================================================
'函数名:WordFilter
'作  用:屏蔽禁用词语
'参  数:strInput       ----需要过滤的词语
'================================================
Function WordFilter(strInput)
    Dim i,strOutput,strFilter
    strFilter = Split("'"&BanedWords&"'",",")
    strOutput = strInput
    For i = 0 To UBound(strFilter)
        strOutput = Replace(strOutput,strFilter(i),String(Len(strFilter(i)),"*"))
    Next
    WordFilter = strOutput
End Function
'================================================
'函数名:SendMail
'作  用:用Jmail组件发送邮件
'参  数:SMTPServer    ----服务器地址
'       FromMail      ----发信人地址
'       FromName      ----发信人名称
'       MailServerUserName          ----发送邮件服务器登录名
'       MailServerPassword          ----发送邮件服务器登录密码
'       ToEmail       ----接收人邮件地址
'       Subject       ----邮件标题
'       Body          ----邮件内容
'用  法:SendMail("服务器地址", "发信人地址", "发信人名称", "登录名", "登录密码", "接收人邮件地址", "邮件标题", "邮件内容")
'================================================
Function SendMail(SMTPServer, FromMail, FromName, MailServerUserName, MailServerPassword, ToEmail, Subject, Body)
    Dim jmail
    Set jmail = Server.CreateObject("JMAIL.Message")   '建立发送邮件的对象
    If Err.Number <> 0 Then
        SendMail = 1
        Exit Function
    End If
    jmail.silent = True    '屏蔽例外错误，返回FALSE跟TRUE两值
    jmail.logging = False   '启用邮件日志
    jmail.Charset = "UTF-8"     '邮件的文字编码为中文
    jmail.ISOEncodeHeaders = False '防止邮件标题乱码
    jmail.ContentType = "text/html"    '邮件的格式为HTML格式
    jmail.AddRecipient ToEmail    '邮件收件人的地址
    jmail.From = FromMail  '发件人的E-MAIL地址
    jmail.FromName = FromName   '发件人姓名
    jmail.MailServerUserName = MailServerUserName    '登录邮件服务器所需的用户名
    jmail.MailServerPassword = MailServerPassword     '登录邮件服务器所需的密码
    jmail.Subject = Subject    '邮件的标题 
    jmail.Body = Body      '邮件的内容
    jmail.Priority = 1      '邮件的紧急程序，1 为最快，5 为最慢， 3 为默认值
    jmail.Send(SMTPServer)     '执行邮件发送（通过邮件服务器地址）
    jmail.Close()   '关闭对象
    If jmail.ErrorCode <> 0 Then
        SendMail = 2
    Else
        SendMail = 0
    End If
End Function
'================================================
'函数名:GetFolderSize
'作  用:计算某个文件夹的大小
'参  数:FileName ----文件夹路径及文件夹名称
'返回值:数值
'================================================
Public Function GetFolderSize(Folderpath)
    dim fso,d,size,showsize
    set fso=server.createobject("scripting.filesystemobject")
    drvpath=server.mappath(Folderpath)
    if fso.FolderExists(drvpath) Then
        set d=fso.getfolder(drvpath)
        size=d.size
        GetFolderSize=FormatSize(size)
    Else
        GetFolderSize=Folderpath&"文件夹不存在"
    End If
End Function
'================================================
'函数名:GetFileSize
'作  用:计算某个文件的大小
'参  数:FileName ----文件路径及文件名
'返回值:数值
'================================================
Public Function GetFileSize(FileName)
    Dim fso,drvpath,d,size,showsize
    set fso=server.createobject("scripting.filesystemobject")
    filepath=server.mappath(FileName)
    if fso.FileExists(filepath) then
        set d=fso.getfile(filepath)
        size=d.size
        GetFileSize=FormatSize(size)
    Else
        GetFileSize=FileName&"文件不存在"
    End If
    set fso=nothing
End Function
'================================================
'函数名:IsObjInstalled
'作  用:检查组件是否安装
'参  数:strClassString ----组件名称
'返回值:false不存在，true存在
'================================================
Public Function IsObjInstalled(strClassString)
    IsObjInstalled=False
    Err=0
    Dim xTestObj
    Set xTestObj=Server.CreateObject(strClassString)
    If 0=Err Then IsObjInstalled=True
    Set xTestObj=Nothing
    Err=0
End Function
'================================================
'函数名:FormatSize
'作  用:大小格式化
'参  数:size ----要格式化的大小
'================================================
Public Function FormatSize(dsize)
    if dsize>=1073741824 then
        FormatSize=Formatnumber(dsize/1073741824,2) & " GB"
    elseif dsize>=1048576 then
        FormatSize=Formatnumber(dsize/1048576,2) & " MB"
    elseif dsize>=1024 then
        FormatSize=Formatnumber(dsize/1024,2) & " KB"
    else
        FormatSize=dsize & " Byte"
    end if
End Function
'================================================
'函数名:WriteToTextFile
'作  用:adodb.stream写入文件
'参  数:FileUrl   ----需要写入文件的名称
'       Str       ----需要写入的内容
'       CharSet   ----写入文件的编码
'================================================
Function WriteToTextFile(FileUrl,Str,CharSet)
    Dim stm
    set stm=server.CreateObject("adodb.stream")
    stm.Type=2 '以本模式读取
    stm.mode=3
    stm.charset=CharSet
    stm.open
    stm.WriteText Str
    stm.SaveToFile server.MapPath(FileUrl),2
    stm.flush
    stm.Close
    set stm=nothing
End Function
'================================================
'函数名:ReadFromTextFile
'作  用:adodb.stream读取文件
'参  数:FileUrl   ----需要写入文件的名称
'       CharSet   ----写入文件的编码
'================================================
Function ReadFromTextFile(FileUrl,CharSet)
dim str
set stm=server.CreateObject("adodb.stream")
stm.Type=2 '以本模式读取
stm.mode=3
stm.charset=CharSet
stm.open
stm.loadfromfile server.MapPath(FileUrl)
str=stm.readtext
stm.Close
set stm=nothing
ReadFromTextFile=str
End Function
'================================================
'函数名:BytesToBstr
'作  用:adodb.stream变量转类型
'参  数:Body   ----要转换的变量
'       Cset   ----要转换的类型
'================================================
Function BytesToBstr(Body,Cset) 
   Dim Objstream 
   Set Objstream = Server.CreateObject("adodb.stream") 
   objstream.Type = 1 
   objstream.Mode =3 
   objstream.Open 
   objstream.Write body 
   objstream.Position = 0 
   objstream.Type = 2 
   objstream.Charset = Cset 
   BytesToBstr = objstream.ReadText  
   objstream.Close 
   set objstream = nothing 
End Function
'================================================
'函数名:FSOFileRead
'作  用:FSO读取文件
'参  数:filename   ----需要读取文件的名称
'================================================
Function FSOFileRead(filename)
    Dim objFSO,objCountFile,FiletempData
    Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
    Set objCountFile = objFSO.OpenTextFile(Server.MapPath(filename),1,True)
    If objCountFile.AtEndOfStream = false Then FSOFileRead = objCountFile.ReadAll
    objCountFile.Close
    Set objCountFile=Nothing
    Set objFSO = Nothing
End Function
'================================================
'函数名:CreateFolder
'作  用:FSO创建文件夹
'参  数:fldr       ----需要创建的文件夹名称
'================================================
Function CreateFolder(fldr)
    Dim fso,f
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.CreateFolder(Server.MapPath(fldr))
    CreateFolder = f.Path
    Set f=nothing
    Set fso=nothing
    Select Case Err
        Case 424 Response.Write "路径未找到或者该目录没有写入权限."
    End Select
End Function
'================================================
'函数名:createhtml
'作  用:FSO创建文件
'参  数:path            ----需要创建的文件路径
'       str             ----需要创建的文件名称
'================================================
Function createhtml(path,str)
GetFold=split(path,"/")
For e=0 to Ubound(GetFold)-1
    if fldr="" then
        fldr=GetFold(e)
    else
        fldr=fldr&"/"&GetFold(e)
    end if
    If IsFolder(fldr)=false then
        CreateFolder fldr
    End if
Next
Set fso = Server.CreateObject("Scripting.FileSystemObject")
Set fout = fso.CreateTextFile(server.mappath(path))
fout.Write str
fout.close
set fso = nothing
Select Case Err
    Case 424 Response.Write "路径未找到或者该目录没有写入权限."
End Select
End Function
'================================================
'函数名:delfile
'作  用:FSO删除文件
'参  数:path            ----需要删除文件的路径
'================================================
Function delfile(path)
If IsExists(path)=True Then
    set fso = server.CreateObject("Scripting.FileSystemObject")
    fso.DeleteFile(server.mappath(path))
    set fso = nothing
End If
Select Case Err
    Case 424 Response.Write "路径未找到或者该目录没有写入权限."
End Select
End Function
'================================================
'函数名:delfolder
'作  用:FSO删除文件夹
'参  数:path            ----需要删除文件夹的路径
'================================================
Function delfolder(path)
If IsFolder(path)=True Then
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.DeleteFolder(server.mappath(path))
    set fso = Nothing
End If
End Function
'================================================
'函数名:IsExists
'作  用:FSO检测文件是否存在
'参  数:filespec         ----需要检测文件的路径
'================================================
Function IsExists(filespec)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    If (fso.FileExists(server.MapPath(filespec))) Then
        IsExists = True
    Else
        IsExists = False
    End If
    Set fso=nothing
End Function
'================================================
'函数名:IsFolder
'作  用:FSO检测文件夹是否存在
'参  数:Folder          ----需要检测文件的路径
'================================================
Function IsFolder(Folder)
    Set fso = CreateObject("Scripting.FileSystemObject")
    If FSO.FolderExists(server.MapPath(Folder)) Then
        IsFolder = True
    Else
        IsFolder = False
    End If
    Set fso=nothing
End Function
'================================================
'函数名:fldrename
'作  用:FSO修改文件夹
'参  数:nowfld          ----现在文件的名称及路径
'       newfld          ----修改文件的名称及路径
'================================================
Function fldrename(nowfld,newfld)
    if nowfld<>newfld then
        nowfld=server.mappath(nowfld)
        newfld=server.mappath(newfld)
        Set fso = CreateObject("Scripting.FileSystemObject")
        if not fso.FolderExists(nowfld) then
            response.write("需要修改的文件夹路径不正确")
        else
            fso.CopyFolder nowfld,newfld
            fso.DeleteFolder(nowfld)
        end if
        set fso=nothing
        Select Case Err
            Case 424 Response.Write "路径未找到或者该目录没有写入权限."
        End Select
    end if
End Function
'================================================
'函数名:Checkstr
'作  用:全站SQL防注入
'================================================
Function Checkstr(Str)
    If Isnull(Str) Then 
    CheckStr = "" 
    Exit Function 
    End If 
    Str = Replace(Str,Chr(0),"", 1, -1, 1) 
    Str = Replace(Str, """", """", 1, -1, 1) 
    Str = Replace(Str,"<","<", 1, -1, 1) 
    Str = Replace(Str,">",">", 1, -1, 1) 
    Str = Replace(Str, "script", "script", 1, -1, 0) 
    Str = Replace(Str, "SCRIPT", "SCRIPT", 1, -1, 0) 
    Str = Replace(Str, "Script", "Script", 1, -1, 0) 
    Str = Replace(Str, "script", "Script", 1, -1, 1) 
    Str = Replace(Str, "object", "object", 1, -1, 0) 
    Str = Replace(Str, "OBJECT", "OBJECT", 1, -1, 0) 
    Str = Replace(Str, "Object", "Object", 1, -1, 0) 
    Str = Replace(Str, "object", "Object", 1, -1, 1) 
    Str = Replace(Str, "applet", "applet", 1, -1, 0) 
    Str = Replace(Str, "APPLET", "APPLET", 1, -1, 0) 
    Str = Replace(Str, "Applet", "Applet", 1, -1, 0) 
    Str = Replace(Str, "applet", "Applet", 1, -1, 1) 
    Str = Replace(Str, "[", "[") 
    Str = Replace(Str, "]", "]") 
    Str = Replace(Str, """", "", 1, -1, 1) 
    Str = Replace(Str, "=", "=", 1, -1, 1) 
    Str = Replace(Str, "'", "''", 1, -1, 1) 
    Str = Replace(Str, "select", "select", 1, -1, 1) 
    Str = Replace(Str, "execute", "execute", 1, -1, 1) 
    Str = Replace(Str, "exec", "exec", 1, -1, 1) 
    Str = Replace(Str, "join", "join", 1, -1, 1) 
    Str = Replace(Str, "union", "union", 1, -1, 1) 
    Str = Replace(Str, "where", "where", 1, -1, 1) 
    Str = Replace(Str, "insert", "insert", 1, -1, 1) 
    Str = Replace(Str, "delete", "delete", 1, -1, 1) 
    Str = Replace(Str, "update", "update", 1, -1, 1) 
    Str = Replace(Str, "like", "like", 1, -1, 1) 
    Str = Replace(Str, "drop", "drop", 1, -1, 1) 
    Str = Replace(Str, "create", "create", 1, -1, 1) 
    Str = Replace(Str, "rename", "rename", 1, -1, 1) 
    Str = Replace(Str, "count", "count", 1, -1, 1) 
    Str = Replace(Str, "chr", "chr", 1, -1, 1) 
    Str = Replace(Str, "mid", "mid", 1, -1, 1) 
    Str = Replace(Str, "truncate", "truncate", 1, -1, 1) 
    Str = Replace(Str, "nchar", "nchar", 1, -1, 1) 
    Str = Replace(Str, "char", "char", 1, -1, 1) 
    Str = Replace(Str, "alter", "alter", 1, -1, 1) 
    Str = Replace(Str, "cast", "cast", 1, -1, 1) 
    Str = Replace(Str, "exists", "exists", 1, -1, 1) 
    Str = Replace(Str,Chr(13),"<br>", 1, -1, 1) 
    CheckStr = Replace(Str,"'","''", 1, -1, 1) 
End Function
%>