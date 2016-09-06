<!--#include file="../config.inc.asp"-->
<%
Call CheckPostUrl()
Dim act,des
act=Request.QueryString ("act")
Set des = new DesClass
Dim u_id,u_username,u_password,rememberme,u_level,u_email,u_avatar,u_mobile,u_qq,u_sex,u_age,u_question,u_answer,chkid
Dim a_id,a_title,a_content,a_uptime,a_version,a_view,a_downnum,a_tag,a_vendor,a_score,a_downlink,a_icon,cid,a_keyword,a_description,a_seoname,a_istop
Dim ad_id,ad_title,ad_urllink,ad_piclink
Dim l_id,l_title,l_link,l_order,l_status
Dim c_id,c_title,c_level,c_seoname,c_keywords,c_description,c_appnum,c_order
Dim cm_id,cm_content,cm_status
u_id=Checkstr(Clng(request.QueryString("u_id")))
u_username=Checkstr(request.Form("u_username"))
u_password=des.Encode(Checkstr(request.Form("u_password")))
u_level=Checkstr(Clng(request.Form("u_level")))
u_email=Checkstr(request.Form("u_email"))
u_avatar=Checkstr(request.Form("u_avatar"))
u_mobile=Checkstr(request.Form("u_mobile"))
u_qq=Checkstr(request.Form("u_qq"))
u_sex=Checkstr(Clng(request.Form("u_sex")))
u_age=Checkstr(Clng(request.Form("u_age")))
u_question=Checkstr(Clng(request.Form("u_question")))
u_answer=Checkstr(request.Form("u_answer"))
chkid=Checkstr(request.form("chkid"))
rememberme=Checkstr(Cint(request.Form("rememberme")))
ad_id=Checkstr(Clng(request.QueryString("ad_id")))
ad_title=Checkstr(request.Form("ad_title"))
ad_urllink=Checkstr(request.Form("ad_urllink"))
ad_piclink=Checkstr(request.Form("ad_piclink"))
l_id=Checkstr(Clng(request.QueryString("l_id")))
l_title=Checkstr(request.Form("l_title"))
l_link=Checkstr(request.Form("l_link"))
l_order=Checkstr(Cint(request.Form("l_order")))
l_status=Checkstr(Cint(request.Form("l_status")))
c_id=Checkstr(Clng(request.QueryString("c_id")))
c_title=Checkstr(request.Form("c_title"))
c_level=Checkstr(Cint(request.Form("c_level")))
c_seoname=Checkstr(request.Form("c_seoname"))
c_keywords=Checkstr(request.Form("c_keywords"))
c_description=Checkstr(request.Form("c_description"))
c_appnum=Checkstr(Clng(request.Form("c_appnum")))
c_order=Checkstr(Cint(request.Form("c_order")))
cm_id=Checkstr(Clng(request.QueryString("cm_id")))
cm_content=Checkstr(WordFilter(request.Form("cm_content")))
a_id=Checkstr(Clng(request.QueryString("a_id")))
a_title=Checkstr(request.Form("a_title"))
a_content=Checkstr(request.Form("a_content"))
a_uptime=Checkstr(request.Form("a_uptime"))
a_version=Checkstr(request.Form("a_version"))
a_view=Checkstr(Clng(request.Form("a_view")))
a_downnum=Checkstr(Clng(request.Form("a_downnum")))
a_tag=Checkstr(request.Form("a_tag"))
a_vendor=Checkstr(request.Form("a_vendor"))
a_score=Checkstr(Clng(request.Form("a_score")))
a_downlink=Checkstr(request.Form("a_downlink"))
a_icon=Checkstr(request.Form("a_icon"))
cid=Checkstr(Clng(request.Form("c_id")))
a_keyword=Checkstr(request.Form("a_keyword"))
a_description=Checkstr(request.Form("a_description"))
a_seoname=Checkstr(request.Form("a_seoname"))
a_istop=Checkstr(Clng(request.Form("a_istop")))

Select Case act
    Case "login"
        Call login()
    Case "logout"
        Call logout()
    Case "addlink"
        Call addlink()
    Case "editlinks"
        Call editlinks()
    Case "addad"
        Call addad()
    Case "editads"
        Call editads()
    Case "delad"
        Call delad()
    Case "dellink"
        Call dellink()
    Case "delcomment"
        Call delcomment()
    Case "delcomments"
        Call delcomments()
    Case "register"
        Call adduser()
    Case "editusers"
        Call editusers()
    Case "deluser"
        Call deluser()
    Case "delusers"
        Call delusers()
    Case "addcategory"
        Call addcategory()
    Case "editcategorys"
        Call editcategorys()
    Case "delcategory"
        Call delcategory()
    Case "addapp"
        Call addapp()
    Case "editapps"
        Call editapps()
    Case "delapp"
        Call delapp()
    Case "delapps"
        Call delapps()
    Case "forget"
        Call forget()
    Case "saveconfig"
        Call saveconfig()
    Case "settop"
        Call settop()
    Case "setreview"
        Call setreview()
End Select

Sub login()
    Dim u_loginip,u_logintime
    u_loginip=GetIP()
    u_logintime=FormatTime(now,7)
    If isNul(u_username) Or isNul(u_password) Then
        alertMsg "用户名或密码必须填写!","javascript:history.back(-1)"
    Else
        Set rs = mysql.exec("select * from app_users where u_username='"&u_username&"' and u_password='"&u_password&"'", 1)
        If not(rs.bof and rs.eof) Then
            If u_username=rs("u_username") and u_password=rs("u_password") Then
                Response.Cookies(""&SiteCookie&"")("u_username")=rs("u_username")
				Response.Cookies(""&SiteCookie&"")("u_password")=rs("u_password")
                Response.Cookies(""&SiteCookie&"")("u_level")=rs("u_level")
                If rememberme="1" then
                    Response.Cookies(""&SiteCookie&"").expires=now()+1
                End If
                If Request.Cookies(""&SiteCookie&"")("u_level")="0" Then
                    alertMsg "登录成功","/admin/manage-dashboard.asp"
                Else
                    alertMsg "登录成功","/admin/manage-profile.asp"
                End If
            End If
        Else
            Response.Cookies(""&SiteCookie&"")("u_username")=""
            Response.Cookies(""&SiteCookie&"")("u_password")=""
            Response.Cookies(""&SiteCookie&"")("u_level")=""
            Response.Cookies(""&SiteCookie&"").expires=now()-1
            alertMsg "用户名或密码错误","javascript:history.back(-1)"
        End If
        Call mysql.exec("update app_users set u_logintime='"&u_logintime&"',u_loginip='"&u_loginip&"' where u_username='"&u_username&"'",0)
        Call clear_mysql()
    End If
End Sub

Sub logout()
    Response.Cookies(""&SiteCookie&"")("u_username")=""
    Response.Cookies(""&SiteCookie&"")("u_password")=""
    Response.Cookies(""&SiteCookie&"")("u_level")=""
    Response.Cookies(""&SiteCookie&"").expires=now()-1
    alertMsg "注销成功","../login.asp"
End Sub

Sub addlink()
	If isNul(l_title) or Not isUrl(l_link) or Not isNum(l_order) or Not isNum(l_status) Then
		alertMsg "传入ID参数错误","javascript:history.back(-1)"
	Else
		Call mysql.exec("insert into app_links(l_title,l_link,l_order,l_status) values('"&l_title&"','"&l_link&"','"&l_order&"','"&l_status&"')",0)
		alertMsg "添加链接成功","manage-links.asp"
	End If
End Sub

Sub editlinks()
	If Not isNum(l_id) or isNul(l_title) or Not isUrl(l_link) or Not isNum(l_order) or Not isNum(l_status) Then
		alertMsg "传入ID参数错误","javascript:history.back(-1)"
	Else
        Call mysql.exec("update app_links set l_title='"&l_title&"',l_link='"&l_link&"',l_order='"&l_order&"',l_status='"&l_status&"' where l_id="&l_id, 0)
		alertMsg "修改链接成功","manage-links.asp"
	End If
End Sub

Sub dellink()
    If Not isNum(l_id) Then
        alertMsg "传入ID参数错误","javascript:history.back(-1)"
    Else
        Call mysql.exec("delete from app_links where l_id ="&l_id, 0)
        alertMsg "删除链接成功","manage-links.asp"
    End If
End Sub

Sub addad()
	If isNul(ad_title) or Not isUrl(ad_urllink) or isNul(ad_piclink) Then
		alertMsg "传入ID参数错误","javascript:history.back(-1)"
	Else
		Call mysql.exec("insert into app_ads(ad_title,ad_urllink,ad_piclink) values('"&ad_title&"','"&ad_urllink&"','"&ad_piclink&"')",0)
		alertMsg "添加广告成功","manage-ads.asp"
	End If
End Sub

Sub editads()
	If Not isNum(ad_id) or isNul(ad_title) or Not isUrl(ad_urllink) or isNul(ad_piclink) Then
		alertMsg "传入ID参数错误","javascript:history.back(-1)"
	Else
        Call mysql.exec("update app_ads set ad_title='"&ad_title&"',ad_urllink='"&ad_urllink&"',ad_piclink='"&ad_piclink&"' where ad_id="&ad_id, 0)
		alertMsg "修改广告成功","manage-ads.asp"
	End If
End Sub

Sub delad()
    If Not isNum(ad_id) Then
        alertMsg "传入ID参数错误","javascript:history.back(-1)"
    Else
        Call mysql.exec("delete from app_ads where ad_id ="&ad_id, 0)
        alertMsg "删除广告成功","manage-ads.asp"
    End If
End Sub

Sub delcomment()
    If Not isNum(cm_id) Then
        alertMsg "传入ID参数错误","javascript:history.back(-1)"
    Else
        Call mysql.exec("delete from app_comments where cm_id in("&cm_id&")", 0)
        alertMsg "删除评论成功","manage-comments.asp"
    End If
End Sub

Sub delcomments()
    If isNul(chkid) Then
        alertMsg "传入ID参数错误","javascript:history.back(-1)"
    Else
        Call mysql.exec("delete from app_comments where cm_id in("&chkid&")", 0)
        alertMsg "删除评论成功","manage-comments.asp"
    End If
End Sub

Sub adduser()
    Dim u_regtime
    u_regtime=FormatTime(now,7)
	If isNul(u_username) or isNul(u_password) or isNul(u_email) or Not isNum(u_sex) or Not isNum(u_question) or isNul(u_answer) Then
		alertMsg "注册信息不完整","javascript:history.back(-1)"
	Else
        Call CheckRegname()
        Set rs = mysql.exec("select u_username from app_users where u_username='"&u_username&"'", 1)
        If not(rs.bof and rs.eof) Then
            alertMsg "用户名已存在","javascript:history.back(-1)"
        Else
            Call mysql.exec("insert into app_users(u_username,u_password,u_email,u_level,u_sex,u_regtime,u_question,u_answer) values('"&u_username&"','"&u_password&"','"&u_email&"',1,'"&u_sex&"','"&u_regtime&"','"&u_question&"','"&u_answer&"')",0)
            alertMsg "注册成功","/login.asp"
        End If
	End If
End Sub

Sub editusers()
	If Not isNum(u_id) or isNul(u_username) or isNul(u_password) or isNul(u_email) or Not isNum(u_level) or isNul(u_avatar) or isNul(u_mobile) or isNul(u_qq) or Not isNum(u_sex) or Not isNum(u_age) or Not isNum(u_question) or isNul(u_answer) Then
		alertMsg "传入ID参数错误","javascript:history.back(-1)"
	Else
        Call mysql.exec("update app_users set u_username='"&u_username&"',u_password='"&u_password&"',u_email='"&u_email&"',u_level='"&u_level&"',u_avatar='"&u_avatar&"',u_mobile='"&u_mobile&"',u_qq='"&u_qq&"',u_sex='"&u_sex&"',u_age='"&u_age&"',u_question='"&u_question&"',u_answer='"&u_answer&"' where u_id="&u_id, 0)
		alertMsg "修改用户成功","manage-users.asp"
	End If
End Sub

Sub deluser()
    If Not isNum(u_id) Then
        alertMsg "传入ID参数错误","javascript:history.back(-1)"
    Else
        Call mysql.exec("delete from app_users where u_id in("&u_id&")", 0)
        alertMsg "删除用户成功","manage-users.asp"
    End If
End Sub

Sub delusers()
    If isNul(chkid) Then
        alertMsg "传入ID参数错误","javascript:history.back(-1)"
    Else
        Call mysql.exec("delete from app_users where u_id in("&chkid&")", 0)
        alertMsg "删除用户成功","manage-users.asp"
    End If
End Sub

Sub addcategory()
	If isNul(c_title) or Not isNum(c_level) or isNul(c_seoname) or isNul(c_keywords) or isNul(c_description) or Not isNum(c_order) Then
		alertMsg "传入ID参数错误","javascript:history.back(-1)"
	Else
		Call mysql.exec("insert into app_category(c_title,c_level,c_seoname,c_keywords,c_description,c_appnum,c_order) values('"&c_title&"','"&c_level&"','"&c_seoname&"','"&c_keywords&"','"&c_description&"',0,'"&c_order&"')",0)
		alertMsg "添加分类成功","manage-category.asp"
	End If
End Sub

Sub editcategorys()
	If Not isNum(c_id) or isNul(c_title) or Not isNum(c_level) or isNul(c_seoname) or isNul(c_keywords) or isNul(c_description) or Not isNum(c_order) or Not isNum(c_appnum) Then
		alertMsg "传入ID参数错误","javascript:history.back(-1)"
	Else
        Call mysql.exec("update app_category set c_title='"&c_title&"',c_level='"&c_level&"',c_seoname='"&c_seoname&"',c_keywords='"&c_keywords&"',c_description='"&c_description&"',c_appnum='"&c_appnum&"',c_order='"&c_order&"' where c_id="&c_id, 0)
		alertMsg "修改分类成功","manage-category.asp"
	End If
End Sub

Sub delcategory()
    If Not isNum(c_id) or c_appnum>"0" Then
        alertMsg "传入ID参数错误或该分类下有数据","javascript:history.back(-1)"
    Else
        If c_id="1" or c_id="2" Then
            alertMsg "顶级分类不可删除","javascript:history.back(-1)"
        Else
            Call mysql.exec("delete from app_category where c_id in("&c_id&")", 0)
            alertMsg "删除分类成功","manage-category.asp"
        End If
    End If
End Sub

Sub addapp()
	If isNul(a_title) or isNul(a_content) or isNul(a_uptime) or isNul(a_version) or Not isNum(a_view) or Not isNum(a_downnum) or isNul(a_tag) or isNul(a_vendor) or Not isNum(a_score) or isNul(a_downlink) or isNul(a_icon) or Not isNum(cid) or isNul(a_keyword) or isNul(a_description) or isNul(a_seoname) or Not isNum(a_istop) Then
		alertMsg "传入ID参数错误","javascript:history.back(-1)"
	Else
		Call mysql.exec("insert into app_apps(a_title,a_content,a_uptime,a_version,a_view,a_downnum,a_tag,a_vendor,a_score,a_downlink,a_icon,c_id,a_keyword,a_description,a_seoname,a_istop) values('"&a_title&"','"&a_content&"','"&a_uptime&"','"&a_version&"','"&a_view&"','"&a_downnum&"','"&a_tag&"','"&a_vendor&"','"&a_score&"','"&a_downlink&"','"&a_icon&"','"&cid&"','"&a_keyword&"','"&a_description&"','"&a_seoname&"','"&a_istop&"')",0)
        alertMsg "添加应用成功","manage-apps.asp"
	End If
End Sub

Sub editapps()
	If Not isNum(a_id) or isNul(a_title) or isNul(a_content) or isNul(a_uptime) or isNul(a_version) or Not isNum(a_view) or Not isNum(a_downnum) or isNul(a_tag) or isNul(a_vendor) or Not isNum(a_score) or isNul(a_downlink) or isNul(a_icon) or Not isNum(cid) or isNul(a_keyword) or isNul(a_description) or isNul(a_seoname) or Not isNum(a_istop) Then
		alertMsg "传入ID参数错误","javascript:history.back(-1)"
	Else
        Call mysql.exec("update app_apps set a_title='"&a_title&"',a_content='"&a_content&"',a_uptime='"&a_uptime&"',a_version='"&a_version&"',a_view='"&a_view&"',a_downnum='"&a_downnum&"',a_tag='"&a_tag&"',a_vendor='"&a_vendor&"',a_score='"&a_score&"',a_downlink='"&a_downlink&"',a_icon='"&a_icon&"',c_id='"&cid&"',a_keyword='"&a_keyword&"',a_description='"&a_description&"',a_seoname='"&a_seoname&"',a_istop='"&a_istop&"' where a_id="&a_id, 0)
        alertMsg "修改应用成功","manage-apps.asp"
	End If
End Sub

Sub delapp()
    If Not isNum(a_id) Then
        alertMsg "传入ID参数错误","javascript:history.back(-1)"
    Else
        Call mysql.exec("delete from app_apps where a_id in("&a_id&")", 0)
        alertMsg "删除应用成功","manage-apps.asp"
    End If
End Sub

Sub delapps()
    If isNul(chkid) Then
        alertMsg "传入ID参数错误","javascript:history.back(-1)"
    Else
        Call mysql.exec("delete from app_apps where a_id in("&chkid&")", 0)
        alertMsg "删除应用成功","manage-apps.asp"
    End If
End Sub

Sub forget()
    Dim sql
	If isNul(u_question) and isNul(u_answer) Then
		alertMsg "密保和问题答案不能为空","javascript:history.back(-1)"
	Else
        Call mysql.exec("", -1)
        sql = "select u_username,u_password,u_email,u_question,u_answer from app_users where u_question='"&u_question&"' and u_answer='"&u_answer&"'"
        rs.open sql,conn,1,1
        If Not(rs.Eof And rs.Bof) Then
            Dim FindStatus
            FindStatus=SendMail(""&SMTPServerAdd&"", ""&SMTPMail&"", ""&SiteName&"", ""&SMTPUserName&"", ""&SMTPUserPwd&"", ""&rs("u_email")&"", "友卓网找回密码通知", "尊敬的,"&rs("u_username")&",你的密码是"&des.Decode(rs("u_password"))&",请牢记!")
            Select Case FindStatus
                Case 0
                    alertMsg "找回密码成功,请到邮箱查收邮件","../login.asp"
                Case 1
                    alertMsg "没有安装JMAIL组件,请联系管理员","../login.asp"
                Case 2
                    alertMsg "邮件发送失败,请联系管理员","../login.asp"
            End Select
        Else
            alertMsg "密保答案错误","javascript:history.back(-1)"
        End If
        rs.close
        set rs=nothing
	End If
End Sub

Sub getusername()
    Dim rsc,rsc_uid
    rsc_uid=page(1)
    Set rsc = mysql.exec("select u_username from app_users where u_id="&rsc_uid, 1)
    If not rsc.eof Then
        echo ""&rsc("u_username")&""
    Else
        echo "该用户已被删除"
    End If
    rsc.close
    set rsc=nothing
End Sub

Sub GetCMIDName()
    Dim rscm,rscm_aid
    rscm_aid=page(2)
    Set rscm = mysql.exec("select a_title from app_apps where a_id="&rscm_aid, 1)
    If not rscm.eof Then
        a_title=left(rscm("a_title"),12)
        echo ""&a_title&""
    Else
        echo "该应用已被删除"
    End If
    rscm.close
    set rscm=nothing
End Sub

Sub settop()
    Dim istop
    istop=Checkstr(Clng(request.QueryString("istop")))
	If Not isNum(a_id) or Not isNum(istop) Then
		alertMsg "传入参数错误","javascript:history.back(-1)"
	Else
        Select case istop
            Case 0
                Call mysql.exec("update app_apps set a_istop=1 where a_id="&a_id, 0)
                alertMsg "修改推荐成功","manage-apps.asp"
            Case 1
                Call mysql.exec("update app_apps set a_istop=0 where a_id="&a_id, 0)
                alertMsg "修改推荐成功","manage-apps.asp"
        End Select
	End If
End Sub

Sub setreview()
    Dim isreview
    isreview=Checkstr(Clng(request.QueryString("isreview")))
	If Not isNum(cm_id) or Not isNum(isreview) Then
		alertMsg "传入参数错误","javascript:history.back(-1)"
	Else
        Select case isreview
            Case 0
                Call mysql.exec("update app_comments set cm_status=1 where cm_id="&cm_id, 0)
                alertMsg "修改评论状态成功","manage-comments.asp"
            Case 1
                Call mysql.exec("update app_comments set cm_status=0 where cm_id="&cm_id, 0)
                alertMsg "修改评论状态成功","manage-comments.asp"
        End Select
	End If
End Sub

Sub saveconfig()
    Dim CheckObject
    CheckObject=IsObjInstalled("adodb.stream")
    Dim FStr,AS_SiteName,AS_SiteUrl,AS_WebmasterEmail,AS_SiteDBPath,AS_SiteDBPathBack,AS_SiteCookie,AS_PublicTheme,AS_IndexTheme,AS_AdminTheme,AS_IndexDes,AS_IndexKey
    Dim AS_AllowReg,AS_AllowComments,AS_ListComments,AS_ListApps,AS_AvatarPath,AS_AppPath,AS_AdPath,AS_CountCode,AS_SMTPServerAdd,AS_SMTPMail,AS_SMTPUserName
    Dim AS_SMTPUserPwd,AS_BanedWords,AS_SiteStatus,AS_CopyRight
    AS_SiteName=Checkstr(request.Form("SiteName"))
    AS_SiteUrl=Checkstr(request.Form("SiteUrl"))
    AS_WebmasterEmail=Checkstr(request.Form("WebmasterEmail"))
    AS_SiteDBPath=Checkstr(request.Form("SiteDBPath"))
    AS_SiteDBPathBack=Checkstr(request.Form("SiteDBPathBack"))
    AS_SiteCookie=Checkstr(request.Form("SiteCookie"))
    AS_PublicTheme=Checkstr(request.Form("PublicTheme"))
    AS_IndexTheme=Checkstr(request.Form("IndexTheme"))
    AS_AdminTheme=Checkstr(request.Form("AdminTheme"))
    AS_IndexDes=Checkstr(request.Form("IndexDes"))
    AS_IndexKey=Checkstr(request.Form("IndexKey"))
    AS_AllowReg=Checkstr(Clng(request.Form("AllowReg")))
    AS_AllowComments=Checkstr(Clng(request.Form("AllowComments")))
    AS_ListComments=Checkstr(Clng(request.Form("ListComments")))
    AS_ListApps=Checkstr(Clng(request.Form("ListApps")))
    AS_AvatarPath=Checkstr(request.Form("AvatarPath"))
    AS_AppPath=Checkstr(request.Form("AppPath"))
    AS_AdPath=Checkstr(request.Form("AdPath"))
    AS_CountCode=Checkstr(request.Form("CountCode"))
    AS_SMTPServerAdd=Checkstr(request.Form("SMTPServerAdd"))
    AS_SMTPMail=Checkstr(request.Form("SMTPMail"))
    AS_SMTPUserName=Checkstr(request.Form("SMTPUserName"))
    AS_SMTPUserPwd=Checkstr(request.Form("SMTPUserPwd"))
    AS_BanedWords=Checkstr(request.Form("BanedWords"))
    AS_SiteStatus=Checkstr(Clng(request.Form("SiteStatus")))
    AS_CopyRight=Checkstr(request.Form("CopyRight"))
    If CheckObject=false Then
        alertMsg "没有安装adodb.stream组件","javascript:history.back(-1)"
    Else
        If AS_SiteName="" or AS_SiteDBPath="" or AS_SiteCookie="" Then
            alertMsg "数据不完整","javascript:history.back(-1)"
        Else
            FStr = FStr & "<"&chr(37) & vbCrlf
            FStr = FStr & "Const SiteName="""&AS_SiteName&"""" & vbCrlf
            FStr = FStr & "Const SiteUrl="""&AS_SiteUrl&"""" & vbCrlf
            FStr = FStr & "Const WebmasterEmail="""&AS_WebmasterEmail&"""" & vbCrlf
            FStr = FStr & "Const SiteDBPath="""&AS_SiteDBPath&"""" & vbCrlf
            FStr = FStr & "Const SiteDBPathBack="""&AS_SiteDBPathBack&"""" & vbCrlf
            FStr = FStr & "Const SiteCookie="""&AS_SiteCookie&"""" & vbCrlf
            FStr = FStr & "Const PublicTheme="""&AS_PublicTheme&"""" & vbCrlf
            FStr = FStr & "Const IndexTheme="""&AS_IndexTheme&"""" & vbCrlf
            FStr = FStr & "Const AdminTheme="""&AS_AdminTheme&"""" & vbCrlf
            FStr = FStr & "Const IndexDes="""&AS_IndexDes&"""" & vbCrlf
            FStr = FStr & "Const IndexKey="""&AS_IndexKey&"""" & vbCrlf
            FStr = FStr & "Const AllowReg="""&AS_AllowReg&"""" & vbCrlf
            FStr = FStr & "Const AllowComments="""&AS_AllowComments&"""" & vbCrlf
            FStr = FStr & "Const ListComments="""&AS_ListComments&"""" & vbCrlf
            FStr = FStr & "Const ListApps="""&AS_ListApps&"""" & vbCrlf
            FStr = FStr & "Const AvatarPath="""&AS_AvatarPath&"""" & vbCrlf
            FStr = FStr & "Const AppPath="""&AS_AppPath&"""" & vbCrlf
            FStr = FStr & "Const AdPath="""&AS_AdPath&"""" & vbCrlf
            FStr = FStr & "Const CountCode="""&AS_CountCode&"""" & vbCrlf
            FStr = FStr & "Const SMTPServerAdd="""&AS_SMTPServerAdd&"""" & vbCrlf
            FStr = FStr & "Const SMTPMail="""&AS_SMTPMail&"""" & vbCrlf
            FStr = FStr & "Const SMTPUserName="""&AS_SMTPUserName&"""" & vbCrlf
            FStr = FStr & "Const SMTPUserPwd="""&AS_SMTPUserPwd&"""" & vbCrlf
            FStr = FStr & "Const BanedWords="""&AS_BanedWords&"""" & vbCrlf
            FStr = FStr & "Const SiteStatus="""&AS_SiteStatus&"""" & vbCrlf
            FStr = FStr & "Const CopyRight="""&AS_CopyRight&"""" & vbCrlf
            FStr = FStr & chr(37)&">" & vbCrlf
            Call WriteToTextFile("/include/config.asp",FStr,"UTF-8") 
            alertMsg "修改网站配置成功","manage-setting.asp"
        End If
    End If
End Sub
%>
