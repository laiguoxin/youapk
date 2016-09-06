<!--#include file="../include/admin.asp"-->
<%Call Chkuserlevel()%>
<!DOCTYPE html>
<html>
    <head>
        <meta charset="UTF-8">
        <meta http-equiv="X-UA-Compatible" content="IE=edge, chrome=1">
        <meta name="renderer" content="webkit">
        <meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no">
        <title>个人设置 - <%=SiteName%></title>
        <link href="<%=SiteUrl%><%=PublicTheme%>/css/bootstrap.min.css" rel="stylesheet">
        <link href="<%=SiteUrl%><%=AdminTheme%>/css/admin.css" rel="stylesheet">
        <script src="<%=SiteUrl%><%=PublicTheme%>/js/jquery.min.js"></script>
        <script src="<%=SiteUrl%><%=PublicTheme%>/js/bootstrap.min.js"></script>
        <script src="<%=SiteUrl%><%=AdminTheme%>/js/admin.js"></script>
        <script src="<%=SiteUrl%><%=AdminTheme%>/js/extend.js"></script>
</head>
<body>
<div id="wrapper">
    <nav class="navbar navbar-inverse navbar-static-top" role="navigation" style="margin-bottom: 0">
        <div class="navbar-header">
            <button type="button" class="navbar-toggle" data-toggle="collapse" data-target=".navbar-collapse">
                <span class="sr-only">展开导航</span>
                <span class="icon-bar"></span>
                <span class="icon-bar"></span>
                <span class="icon-bar"></span>
            </button>
            <a class="navbar-brand" href="manage-dashboard.asp"><%=SiteName%></a>
        </div>
        <ul class="nav navbar-top-links navbar-right">
            <li class="dropdown">
                <a class="dropdown-toggle" data-toggle="dropdown" href="#">
                    <i class="glyphicon glyphicon-user"></i>  <i class="glyphicon glyphicon-menu-down"></i>
                </a>
                <%If Request.cookies(""&SiteCookie&"")("u_username")<>"" and Request.cookies(""&SiteCookie&"")("u_password")<>"" Then%>
                <ul class="dropdown-menu dropdown-user">
                    <li><a href="#"><i class="glyphicon glyphicon-user"></i> <%=request.cookies(""&SiteCookie&"")("u_username")%></a></li>
                    <li><a href="manage-profile.asp"><i class="glyphicon glyphicon-cog"></i> 个人设置</a></li>
                    <li class="divider"></li>
                    <li><a href="?act=logout"><i class="glyphicon glyphicon-log-out"></i> 退出</a></li>
                </ul>
                <%End if%>
            </li>
        </ul>
        <div class="navbar-default sidebar" role="navigation">
            <div class="sidebar-nav navbar-collapse">
                <ul class="nav" id="side-menu">
                    <%If Request.Cookies(""&SiteCookie&"")("u_level")="0" Then%>
                    <li><a href="manage-dashboard.asp"><i class="glyphicon glyphicon-dashboard"></i> 控制面板</a></li>
                    <li><a href="manage-category.asp"><i class="glyphicon glyphicon-align-justify"></i> 分类管理</a></li>
                    <li><a href="manage-apps.asp"><i class="glyphicon glyphicon-cloud"></i> 应用管理</a></li>
                    <li><a href="manage-comments.asp"><i class="glyphicon glyphicon-comment"></i> 评论管理</a></li>
                    <li><a href="manage-links.asp"><i class="glyphicon glyphicon-link"></i> 链接管理</a></li>
                    <li><a href="manage-ads.asp"><i class="glyphicon glyphicon-usd"></i> 广告管理</a></li>
                    <li><a href="manage-users.asp"><i class="glyphicon glyphicon-user"></i> 用户管理</a></li>
                    <li><a href="manage-setting.asp"><i class="glyphicon glyphicon-cog"></i> 网站设置</a></li>
                    <%End If%>
                </ul>
            </div>
        </div>
    </nav>
    <%
    u_username=Request.cookies(""&SiteCookie&"")("u_username")
    Set rs = mysql.exec("select u_id,u_username,u_email,u_avatar,u_regtime,u_loginip,u_level,u_mobile,u_qq,u_sex,u_age,u_logintime,u_question,u_answer from app_users where u_username='"&u_username&"'", 1)
    %>
    <div id="page-wrapper">
        <div class="row">
            <div class="col-lg-12">
                <h4 class="page-header">个人信息</h1>
            </div>
        </div>
        <div class="row">
            <div class="col-lg-12">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <i class="glyphicon glyphicon-info-sign"></i> 用户信息
                    </div>
                    <div class="panel-body">
                        <div class="list-group">
                            <i class="glyphicon glyphicon-user"></i> 用户名：
                            <span class="text-muted normal"><%=rs("u_username")%></span>
                        </div>
                        <div class="list-group">
                            <i class="glyphicon glyphicon-envelope"></i> 邮箱：
                            <span class="text-muted normal"><%=rs("u_email")%></span>
                        </div>
                        <div class="list-group">
                            <i class="glyphicon glyphicon-star"></i> 用户级别：
                            <span class="text-muted normal"><%If rs("u_level")="0" Then%>管理员<%Else%>普通会员<%End If%></span>
                        </div>
                        <div class="list-group">
                            <i class="glyphicon glyphicon-time"></i> 注册时间：
                            <span class="text-muted normal"><%=rs("u_regtime")%></span>
                        </div>
                        <div class="list-group">
                            <i class="glyphicon glyphicon-picture"></i> 头像：
                            <span class="text-muted normal"><img src="../<%=rs("u_avatar")%>" class="img-circle" width="40" height="40"></span>
                        </div>
                        <div class="list-group">
                            <i class="glyphicon glyphicon-heart"></i> 性别：
                            <span class="text-muted normal"><%If rs("u_sex")="0" Then%>男<%Else%>女<%End If%></span>
                        </div>
                        <div class="list-group">
                            <i class="glyphicon glyphicon-comment"></i> QQ：
                            <span class="text-muted normal"><%=rs("u_qq")%></span>
                        </div>
                        <div class="list-group">
                            <i class="glyphicon glyphicon-tree-conifer"></i> 年龄：
                            <span class="text-muted normal"><%=rs("u_age")%>岁</span>
                        </div>
                        <div class="list-group">
                            <i class="	glyphicon glyphicon-phone"></i> 手机号码：
                            <span class="text-muted normal"><%=rs("u_mobile")%></span>
                        </div>
                        <div class="list-group">
                            <i class="glyphicon glyphicon-exclamation-sign"></i> 上次登录IP：
                            <span class="text-muted normal"><%=rs("u_loginip")%></span>
                        </div>
                        <div class="list-group">
                            <i class="glyphicon glyphicon-calendar"></i> 上次登录时间：
                            <span class="text-muted normal"><%=rs("u_logintime")%></span>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>
</div>
</body>
</html>
<%Call clear_mysql()%>