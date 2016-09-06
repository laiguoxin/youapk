<!--#include file="../include/admin.asp"-->
<%Call Chkadmin()%>
<!DOCTYPE html>
<html>
    <head>
        <meta charset="UTF-8">
        <meta http-equiv="X-UA-Compatible" content="IE=edge, chrome=1">
        <meta name="renderer" content="webkit">
        <meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no">
        <title>网站设置 - <%=SiteName%></title>
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
    <div id="page-wrapper">
        <div class="row">
            <div class="col-lg-12">
                <h4 class="page-header">网站设置</h1>
            </div>
        </div>
        <div class="row">
            <div class="col-lg-12">
                <div class="panel panel-default">
                    <div class="panel-heading">
                        <i class="glyphicon glyphicon-cog"></i> 网站设置
                    </div>
                    <div class="panel-body">
                        <div class="list-group">
                            <table class="table table-responsive table-hover" id="websetting">
                            <form class="form-horizontal" action="?act=saveconfig" method="post" name="saveconfig" role="form">
                                <tbody>
                                    <tr>
                                        <td class="col-md-2">网站名称：</td>
                                        <td class="col-md-10"><input type="text" value="<%=SiteName%>" id="SiteName" name="SiteName" class="form-control" required autofocus></td>
                                    </tr>
                                    <tr>
                                        <td class="col-md-2">网站地址：</td>
                                        <td class="col-md-10"><input type="text" value="<%=SiteUrl%>" id="SiteUrl" name="SiteUrl" class="form-control"></td>
                                    </tr>
                                    <tr>
                                        <td class="col-md-2">管理员邮箱：</td>
                                        <td class="col-md-10"><input type="text" value="<%=WebmasterEmail%>" id="WebmasterEmail" name="WebmasterEmail" class="form-control"></td>
                                    </tr> 
                                    <tr>
                                        <td class="col-md-2">数据库路径：</td>
                                        <td class="col-md-10"><input type="text" value="<%=SiteDBPath%>" id="SiteDBPath" name="SiteDBPath" class="form-control"></td>
                                    </tr>
                                    <tr>
                                        <td class="col-md-2">数据库备份路径：</td>
                                        <td class="col-md-10"><input type="text" value="<%=SiteDBPathBack%>" id="SiteDBPathBack" name="SiteDBPathBack" class="form-control"></td>
                                    </tr>
                                    <tr>
                                        <td class="col-md-2">Cooke名称：</td>
                                        <td class="col-md-10"><input type="text" value="<%=SiteCookie%>" id="SiteCookie" name="SiteCookie" class="form-control"></td>
                                    </tr>
                                    <tr>
                                        <td class="col-md-2">公共皮肤路径：</td>
                                        <td class="col-md-10"><input type="text" value="<%=PublicTheme%>" id="PublicTheme" name="PublicTheme" class="form-control"></td>
                                    </tr>
                                    <tr>
                                        <td class="col-md-2">前台皮肤路径：</td>
                                        <td class="col-md-10"><input type="text" value="<%=IndexTheme%>" id="IndexTheme" name="IndexTheme" class="form-control"></td>
                                    </tr>
                                    <tr>
                                        <td class="col-md-2">后台皮肤路径：</td>
                                        <td class="col-md-10"><input type="text" value="<%=AdminTheme%>" id="AdminTheme" name="AdminTheme" class="form-control"></td>
                                    </tr>
                                    <tr>
                                        <td class="col-md-2">网站描述：</td>
                                        <td class="col-md-10"><textarea type="text" id="IndexDes" name="IndexDes" class="form-control"><%=IndexDes%></textarea></td>
                                    </tr>
                                    <tr>
                                        <td class="col-md-2">关键词：</td>
                                        <td class="col-md-10"><textarea type="text" value="" id="IndexKey" name="IndexKey" class="form-control"><%=IndexKey%></textarea></td>
                                    </tr>
                                    <tr>
                                        <td class="col-md-2">是否允许注册：</td>
                                        <td class="col-md-10"><select class="form-control" name="AllowReg" id="AllowReg"><option value="0" <%If AllowReg="0" Then%>selected="true"<%End If%>>开启</option><option value="1" <%If AllowReg="1" Then%>selected="true"<%End If%>>关闭</option></select></td>
                                    </tr>
                                    <tr>
                                        <td class="col-md-2">是否允许评论：</td>
                                        <td class="col-md-10"><select class="form-control" name="AllowComments" id="AllowComments"><option value="0" <%If AllowComments="0" Then%>selected="true"<%End If%>>开启</option><option value="1" <%If AllowComments="1" Then%>selected="true"<%End If%>>关闭</option></select></td>
                                    </tr>
                                    <tr>
                                        <td class="col-md-2">每页显示评论数量：</td>
                                        <td class="col-md-10"><input type="text" value="<%=ListComments%>" id="ListComments" name="ListComments" class="form-control"></td>
                                    </tr>
                                    <tr>
                                        <td class="col-md-2">每页显示应用数量：</td>
                                        <td class="col-md-10"><input type="text" value="<%=ListApps%>" id="ListApps" name="ListApps" class="form-control"></td>
                                    </tr>
                                    <tr>
                                        <td class="col-md-2">头像上传地址：</td>
                                        <td class="col-md-10"><input type="text" value="<%=AvatarPath%>" id="AvatarPath" name="AvatarPath" class="form-control"></td>
                                    </tr>
                                    <tr>
                                        <td class="col-md-2">应用上传地址：</td>
                                        <td class="col-md-10"><input type="text" value="<%=AppPath%>" id="AppPath" name="AppPath" class="form-control"></td>
                                    </tr>
                                    <tr>
                                        <td class="col-md-2">广告上传地址：</td>
                                        <td class="col-md-10"><input type="text" value="<%=AdPath%>" id="AdPath" name="AdPath" class="form-control"></td>
                                    </tr>
                                    <tr>
                                        <td class="col-md-2">统计代码：</td>
                                        <td class="col-md-10"><textarea id="CountCode" name="CountCode" class="form-control"><%=CountCode%></textarea></td>
                                    </tr>
                                    <tr>
                                        <td class="col-md-2">SMTP服务器地址：</td>
                                        <td class="col-md-10"><input type="text" value="<%=SMTPServerAdd%>" id="SMTPServerAdd" name="SMTPServerAdd" class="form-control"></td>
                                    </tr>
                                    <tr>
                                        <td class="col-md-2">SMTP发信邮箱：</td>
                                        <td class="col-md-10"><input type="text" value="<%=SMTPMail%>" id="SMTPMail" name="SMTPMail" class="form-control"></td>
                                    </tr>
                                    <tr>
                                        <td class="col-md-2">SMTP发信登录账号：</td>
                                        <td class="col-md-10"><input type="text" value="<%=SMTPUserName%>" id="SMTPUserName" name="SMTPUserName" class="form-control"></td>
                                    </tr>
                                    <tr>
                                        <td class="col-md-2">SMTP发信登录密码：</td>
                                        <td class="col-md-10"><input type="text" value="<%=SMTPUserPwd%>" id="SMTPUserPwd" name="SMTPUserPwd" class="form-control"></td>
                                    </tr>
                                    <tr>
                                        <td class="col-md-2">禁用词：</td>
                                        <td class="col-md-10"><textarea id="BanedWords" name="BanedWords" class="form-control"><%=BanedWords%></textarea></td>
                                    </tr>
                                    <tr>
                                        <td class="col-md-2">版权信息：</td>
                                        <td class="col-md-10"><input type="text" value="<%=CopyRight%>" id="CopyRight" name="CopyRight" class="form-control"></td>
                                    </tr>
                                    <tr>
                                        <td class="col-md-2">网站状态：</td>
                                        <td class="col-md-10"><select class="form-control" name="SiteStatus" id="SiteStatus"><option value="0" <%If SiteStatus="0" Then%>selected="true"<%End If%>>开启</option><option value="1" <%If SiteStatus="1" Then%>selected="true"<%End If%>>关闭</option></select></td>
                                    </tr>
                                    <tr>
                                        <td class="col-md-2"></td>
                                        <td class="col-md-10"><button class="btn btn-lg btn-primary btn-block" type="submit">更新</button></td>
                                    </tr>
                                </tbody>
                            </form>
                            </table>
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