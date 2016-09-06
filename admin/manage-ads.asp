<!--#include file="../include/admin.asp"-->
<%Call Chkadmin()%>
<!DOCTYPE html>
<html>
    <head>
        <meta charset="UTF-8">
        <meta http-equiv="X-UA-Compatible" content="IE=edge, chrome=1">
        <meta name="renderer" content="webkit">
        <meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no">
        <title>广告管理 - <%=SiteName%></title>
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
                <h4 class="page-header">广告管理</h1>
            </div>
        </div>
        <div class="row">
            <div class="col-lg-12">
                <div class="panel panel-default">
                    <div class="panel-heading"><i class="glyphicon glyphicon-usd"></i> 广告列表
                        <div class="pull-right">
                            <div class="btn-group">
                                <button type="button" class="btn btn-primary btn-xs" id="add">
                                    <span class="glyphicon glyphicon-plus" aria-hidden="true"></span>添加
                                </button>
                            </div>
                        </div>
                    </div>
                    <div class="panel-body">
                        <div class="table-responsive">
                            <%
                            If act="" Then
                                Set rs = mysql.exec("select ad_id,ad_title,ad_urllink,ad_piclink from app_ads order by ad_id asc", 1)
                            %>
                            <table class="table table-bordered table-hover" id="table">
                                <thead>
                                    <tr class="info">
                                        <th><input type="checkbox" name="checkall" id="checkall" onclick="selectAll();"></th>
                                        <th>ID</th>
                                        <th>广告位名称</th>
                                        <th>广告链接</th>
                                        <th>广告图片</th>
                                        <th>操作</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <%If not rs.eof Then%>
                                    <%Do While Not Rs.Eof%>
                                    <tr>
                                        <td><input type="checkbox" name="checkbox" id="checkbox" value="<%=rs("ad_id")%>" onclick="setSelectAll();"></td>
                                        <td><a href="?act=editad&ad_id=<%=rs("ad_id")%>"><%=rs("ad_id")%></a></td>
                                        <td><%=rs("ad_title")%></td>
                                        <td><a href="<%=rs("ad_urllink")%>"><%=rs("ad_urllink")%></a></td>
                                        <td><img src="../<%=rs("ad_piclink")%>" width="100" height="50"></td>
                                        <td><a href="?act=delad&ad_id=<%=rs("ad_id")%>">删除</a></td>
                                    </tr>
                                    <%
                                    rs.movenext
                                    loop
                                    rs.close
                                    set rs=nothing
                                    %>
                                    <%Else%>
                                    <tr><td colspan="6">暂无数据</td></tr>
                                    <%End If%>
                                </tbody>
                            </table>                      
                            <%
                            ElseIf act="editad" Then
                                If Not isNum(ad_id) Then
                                    alertMsg "参数非法","manage-ads.asp"
                                End If
                            Set rs = mysql.exec("select ad_id,ad_title,ad_urllink,ad_piclink from app_ads where ad_id="&ad_id, 1)
                            %>
                            <table class="table table-bordered table-hover" id="edittable">
                                <form class="form-horizontal" action="?act=editads&ad_id=<%=rs("ad_id")%>" method="post" name="editads" role="form">
                                <tr>
                                    <td><label class="control-label" for="ad_title">广告位名称</label></td>
                                    <td><input type="text" id="ad_title" name="ad_title" value="<%=rs("ad_title")%>" class="form-control" required autofocus></td>
                                </tr>
                                <tr>
                                    <td><label class="control-label" for="ad_urllink">广告链接</label></td>
                                    <td><input type="text" id="ad_urllink" name="ad_urllink" value="<%=rs("ad_urllink")%>" class="form-control" required></td>
                                </tr>
                                <tr>
                                    <td><label class="control-label" for="ad_piclink">广告图片</label></td>
                                    <td><input type="text" id="ad_piclink" name="ad_piclink" value="<%=rs("ad_piclink")%>" class="form-control" required></td>
                                </tr>
                                <tr>
                                    <td><label class="control-label"></label></td>
                                    <td><button type="submit" class="btn btn-lg btn-primary btn-block">更新</button></td>
                                </tr>
                                </form>
                            </table>
                            <%End If%>
                            <table class="table table-bordered table-hover hidden" id="addtable">
                                <form class="form-horizontal" action="?act=addad" method="post" name="addads" role="form">
                                <tr>
                                    <td><label class="control-label" for="ad_title">广告位名称</label></td>
                                    <td><input type="text" id="ad_title" name="ad_title" class="form-control" required autofocus></td>
                                </tr>
                                <tr>
                                    <td><label class="control-label" for="ad_urllink">广告链接</label></td>
                                    <td><input type="text" id="ad_urllink" name="ad_urllink" class="form-control" required></td>
                                </tr>
                                <tr>
                                    <td><label class="control-label" for="ad_piclink">广告图片</label></td>
                                    <td><input type="text" id="ad_piclink" name="ad_piclink" class="form-control" required></td>
                                </tr>
                                <tr>
                                    <td><label class="control-label"></label></td>
                                    <td><button type="submit" class="btn btn-lg btn-primary btn-block">添加</button></td>
                                </tr>
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