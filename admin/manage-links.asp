<!--#include file="../include/admin.asp"-->
<%Call Chkadmin()%>
<!DOCTYPE html>
<html>
    <head>
        <meta charset="UTF-8">
        <meta http-equiv="X-UA-Compatible" content="IE=edge, chrome=1">
        <meta name="renderer" content="webkit">
        <meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no">
        <title>链接管理 - <%=SiteName%></title>
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
                <h4 class="page-header">链接管理</h1>
            </div>
        </div>
        <div class="row">
            <div class="col-lg-12">
                <div class="panel panel-default">
                    <div class="panel-heading"><i class="glyphicon glyphicon-link"></i> 链接列表
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
                                Set rs = mysql.exec("select l_id,l_title,l_link,l_status,l_order from app_links order by l_id asc", 1)
                            %>
                            <table class="table table-bordered table-hover" id="table">
                                <thead>
                                    <tr class="info">
                                        <th><input type="checkbox" name="checkall" id="checkall" onclick="selectAll();"></th>
                                        <th>ID</th>
                                        <th>链接名称</th>
                                        <th>链接地址</th>
                                        <th>链接状态</th>
                                        <th>排序</th>
                                        <th>操作</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <%If not rs.eof Then%>
                                    <%Do While Not Rs.Eof%>
                                    <tr>
                                        <td><input type="checkbox" name="checkbox" id="checkbox" value="<%=rs("l_id")%>" onclick="setSelectAll();"></td>
                                        <td><a href="?act=editlink&l_id=<%=rs("l_id")%>"><%=rs("l_id")%></a></td>
                                        <td><%=rs("l_title")%></td>
                                        <td><a href="<%=rs("l_link")%>" target="_blank"><%=rs("l_link")%></a></td>
                                        <td><%If rs("l_status")="0" Then%>已审核<%Else%>未审核<%End If%></td>
                                        <td><%=rs("l_order")%></td>
                                        <td><a href="?act=dellink&l_id=<%=rs("l_id")%>">删除</a></td>
                                    </tr>
                                    <%
                                    rs.movenext
                                    loop
                                    rs.close
                                    set rs=nothing
                                    %>
                                    <%Else%>
                                    <tr><td colspan="7">暂无数据</td></tr>
                                    <%End If%>
                                </tbody>
                            </table>
                            <%
                            ElseIf act="editlink" Then
                                If Not isNum(l_id) Then
                                    alertMsg "参数非法","manage-links.asp"
                                End If
                            Set rs = mysql.exec("select l_id,l_title,l_link,l_status,l_order from app_links where l_id="&l_id, 1)
                            %>
                            <table class="table table-bordered table-hover" id="edittable">
                                <form class="form-horizontal" action="?act=editlinks&l_id=<%=rs("l_id")%>" method="post" name="editlinks" role="form">
                                <tr>
                                    <td><label class="control-label" for="l_title">链接名称</label></td>
                                    <td><input type="text" id="l_title" value="<%=rs("l_title")%>" name="l_title" class="form-control" required autofocus></td>
                                </tr>
                                <tr>
                                    <td><label class="control-label" for="l_link">链接地址</label></td>
                                    <td><input type="text" id="l_link" value="<%=rs("l_link")%>" name="l_link" class="form-control" required></td>
                                </tr>
                                <tr>
                                    <td><label class="control-label" for="l_order">链接排序</label></td>
                                    <td><input type="number" id="l_order" value="<%=rs("l_order")%>" name="l_order" class="form-control" required></td>
                                </tr>
                                <tr>
                                    <td><label class="control-label" for="l_status">审核状态</label></td>
                                    <td><select class="form-control" name="l_status" id="l_status"><option value="0" <%If rs("l_status")="0" Then%>selected="true"<%End If%>>审核</option><option value="1" <%If rs("l_status")="1" Then%>selected="true"<%End If%>>未审核</option></select></td>
                                </tr>
                                <tr>
                                    <td><label class="control-label"></label></td>
                                    <td><button type="submit" class="btn btn-lg btn-primary btn-block">修改</button></td>
                                </tr>
                                </form>
                            </table>
                            <%End If%>
                            <table class="table table-bordered table-hover hidden" id="addtable">
                                <form class="form-horizontal" action="?act=addlink" method="post" name="addlink" role="form">
                                <tr>
                                    <td><label class="control-label" for="l_title">链接名称</label></td>
                                    <td><input type="text" id="l_title" name="l_title" class="form-control" required autofocus></td>
                                </tr>
                                <tr>
                                    <td><label class="control-label" for="l_link">链接地址</label></td>
                                    <td><input type="text" id="l_link" name="l_link" class="form-control" required></td>
                                </tr>
                                <tr>
                                    <td><label class="control-label" for="l_order">链接排序</label></td>
                                    <td><input type="number" id="l_order" name="l_order" class="form-control" required></td>
                                </tr>
                                <tr>
                                    <td><label class="control-label" for="l_status">审核状态</label></td>
                                    <td><select class="form-control" name="l_status" id="l_status"><option value="0">审核</option><option value="1">未审核</option></select></td>
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