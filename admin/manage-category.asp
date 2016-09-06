<!--#include file="../include/admin.asp"-->
<%Call Chkadmin()%>
<!DOCTYPE html>
<html>
    <head>
        <meta charset="UTF-8">
        <meta http-equiv="X-UA-Compatible" content="IE=edge, chrome=1">
        <meta name="renderer" content="webkit">
        <meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no">
        <title>分类管理 - <%=SiteName%></title>
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
                <h4 class="page-header">分类管理</h1>
            </div>
        </div>
        <div class="row">
            <div class="col-lg-12">
                <div class="panel panel-default">
                    <div class="panel-heading"><i class="glyphicon glyphicon-align-justify"></i> 分类列表
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
                                Set rs = mysql.exec("select c_id,c_title,c_level,c_seoname,c_keywords,c_description,c_order,c_appnum from app_category order by c_id asc", 1)
                            %>
                            <table class="table table-bordered table-hover" id="table">
                                <thead>
                                    <tr class="info">
                                        <th><input type="checkbox" name="checkall" id="checkall" onclick="selectAll();"></th>
                                        <th>ID</th>
                                        <th>分类名称</th>
                                        <th>分类等级</th>
                                        <th>分类别名</th>
                                        <th>KEY关键词</th>
                                        <th>DES关键词</th>
                                        <th>应用数量</th>
                                        <th>分类排序</th>
                                        <th>操作</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <%If not rs.eof Then%>
                                    <%Do While Not Rs.Eof%>
                                    <tr>
                                        <td><input type="checkbox" name="checkbox" id="checkbox" value="<%=rs("c_id")%>" onclick="setSelectAll();"></td>
                                        <td><a href="?act=editcategory&c_id=<%=rs("c_id")%>"><%=rs("c_id")%></a></td>
                                        <td><%=rs("c_title")%></td>
                                        <td><%If rs("c_level")="0" Then%>一级分类<%Else%>二级分类<%End If%></td>
                                        <td><%=rs("c_seoname")%></td>
                                        <td><%=rs("c_keywords")%></td>
                                        <td><%=rs("c_description")%></td>
                                        <td><%=rs("c_appnum")%></td>
                                        <td><%=rs("c_order")%></td>
                                        <td><a href="?act=delcategory&c_id=<%=rs("c_id")%>&c_appnum=<%=rs("c_appnum")%>">删除</a></td>
                                    </tr>
                                    <%
                                    rs.movenext
                                    loop
                                    rs.close
                                    set rs=nothing
                                    %>
                                    <%Else%>
                                    <tr><td colspan="10">暂无数据</td></tr>
                                    <%End If%>
                                </tbody>
                            </table>
                            <%
                            ElseIf act="editcategory" Then
                                If Not isNum(c_id) Then
                                    alertMsg "参数非法","manage-category.asp"
                                End If
                            Set rs = mysql.exec("select c_id,c_title,c_level,c_seoname,c_keywords,c_description,c_order,c_appnum from app_category where c_id="&c_id, 1)
                            %>
                            <table class="table table-bordered table-hover" id="edittable">
                                <form class="form-horizontal" action="?act=editcategorys&c_id=<%=rs("c_id")%>" method="post" name="editcategorys" role="form">
                                <tr>
                                    <td><label class="control-label" for="c_title">分类名称</label></td>
                                    <td><input type="text" id="c_title" value="<%=rs("c_title")%>" name="c_title" class="form-control" required autofocus></td>
                                </tr>
                                <tr>
                                    <td><label class="control-label" for="c_level">分类等级</label></td>
                                    <td><select class="form-control" name="c_level" id="c_level"><option value="0" <%If rs("c_level")="0" Then%>selected="true"<%End If%>>一级分类</option><option value="1" <%If rs("c_level")="1" Then%>selected="true"<%End If%>>二级分类</option></select></td>
                                </tr>
                                <tr>
                                    <td><label class="control-label" for="c_seoname">分类别名</label></td>
                                    <td><input type="text" id="c_seoname" value="<%=rs("c_seoname")%>" name="c_seoname" class="form-control" required></td>
                                </tr>
                                <tr>
                                    <td><label class="control-label" for="c_keywords">分类KEY关键字</label></td>
                                    <td><input type="text" id="c_keywords" value="<%=rs("c_keywords")%>" name="c_keywords" class="form-control" required></td>
                                </tr>
                                <tr>
                                    <td><label class="control-label" for="c_description">分类DES关键字</label></td>
                                    <td><input type="text" id="c_description" value="<%=rs("c_description")%>" name="c_description" class="form-control" required></td>
                                </tr>
                                <tr>
                                    <td><label class="control-label" for="c_appnum">应用数量</label></td>
                                    <td><input type="text" id="c_appnum" value="<%=rs("c_appnum")%>" name="c_appnum" class="form-control" required></td>
                                </tr>
                                <tr>
                                    <td><label class="control-label" for="c_order">分类排序</label></td>
                                    <td><input type="number" id="c_order" value="<%=rs("c_order")%>" name="c_order" class="form-control" required></td>
                                </tr>
                                <tr>
                                    <td><label class="control-label"></label></td>
                                    <td><button type="submit" class="btn btn-lg btn-primary btn-block">更新</button></td>
                                </tr>
                                </form>
                            </table>
                            <%End If%>
                            <table class="table table-bordered table-hover hidden" id="addtable">
                                <form class="form-horizontal" action="?act=addcategory" method="post" name="addcategory" role="form">
                                <tr>
                                    <td><label class="control-label" for="c_title">分类名称</label></td>
                                    <td><input type="text" id="c_title" name="c_title" class="form-control" required autofocus></td>
                                </tr>
                                <tr>
                                    <td><label class="control-label" for="c_level">分类等级</label></td>
                                    <td><select class="form-control" name="c_level" id="c_level"><option value="0">一级分类</option><option value="1">二级分类</option></select></td>
                                </tr>
                                <tr>
                                    <td><label class="control-label" for="c_seoname">分类别名</label></td>
                                    <td><input type="text" id="c_seoname" name="c_seoname" class="form-control" required></td>
                                </tr>
                                <tr>
                                    <td><label class="control-label" for="c_keywords">分类KEY关键字</label></td>
                                    <td><input type="text" id="c_keywords" name="c_keywords" class="form-control" required></td>
                                </tr>
                                <tr>
                                    <td><label class="control-label" for="c_description">分类DES关键字</label></td>
                                    <td><input type="text" id="c_description" name="c_description" class="form-control" required></td>
                                </tr>
                                <tr>
                                    <td><label class="control-label" for="c_order">分类排序</label></td>
                                    <td><input type="number" id="c_order" name="c_order" class="form-control" required></td>
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