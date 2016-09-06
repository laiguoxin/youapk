<!--#include file="../include/admin.asp"-->
<!--#include file="../include/page.class.asp"-->
<%
Call Chkadmin()
Dim page,i,sql
Set page = New PageClass
%>
<!DOCTYPE html>
<html>
    <head>
        <meta charset="UTF-8">
        <meta http-equiv="X-UA-Compatible" content="IE=edge, chrome=1">
        <meta name="renderer" content="webkit">
        <meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no">
        <title>应用管理 - <%=SiteName%></title>
        <link href="<%=SiteUrl%><%=PublicTheme%>/css/bootstrap.min.css" rel="stylesheet">
        <link href="<%=SiteUrl%><%=AdminTheme%>/css/admin.css" rel="stylesheet">
        <link href="<%=SiteUrl%><%=AdminTheme%>/css/summernote.css" rel="stylesheet" type="text/css">
        <script src="<%=SiteUrl%><%=PublicTheme%>/js/jquery.min.js"></script>
        <script src="<%=SiteUrl%><%=PublicTheme%>/js/bootstrap.min.js"></script>
        <script src="<%=SiteUrl%><%=AdminTheme%>/js/admin.js"></script>
        <script src="<%=SiteUrl%><%=AdminTheme%>/js/summernote.min.js"></script>
        <script src="<%=SiteUrl%><%=AdminTheme%>/js/summernote-zh-CN.js"></script>
        <script src="<%=SiteUrl%><%=AdminTheme%>/js/extend.js"></script>
        <script type="text/javascript">
        $(document).ready(function() {
            $('#contentedit').summernote({
                height: 100,
                minHeight: null,
                maxHeight: null,
                focus: true,
                lang: 'zh-CN'
            });
        });
        </script>
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
                <h4 class="page-header">应用管理</h1>
            </div>
        </div>
        <div class="row">
            <div class="col-lg-12">
                <div class="panel panel-default">
                    <form action="?act=delapps" method="post" name="delall" id="delall" role="form">
                    <div class="panel-heading"><i class="glyphicon glyphicon-cloud"></i> 应用列表
                        <div class="pull-right">
                            <div class="btn-group">
                                <button type="button" class="btn btn-primary btn-xs" id="add">
                                    <span class="glyphicon glyphicon-plus" aria-hidden="true"></span>添加
                                </button>
                                <button type="submit" class="btn btn-danger btn-xs" id="del">
                                    <span class="glyphicon glyphicon-remove" aria-hidden="true"></span>删除
                                </button>
                            </div>
                        </div>
                    </div>
                    <div class="panel-body">
                        <div class="table-responsive">
                            <%
                            If act="" Then
                                Call mysql.exec("", -1)
                                sql="select a_id,c_id,a_title,a_uptime,a_version,a_view,a_downnum,a_score,a_istop from app_apps order by a_id desc"
                                rs.open sql,conn,1,1
                            %>
                            <table class="table table-bordered table-hover" id="table">
                                <thead>
                                    <tr class="info">
                                        <th><input type="checkbox" name="checkbox" id="checkbox" onClick="javascript:SelectAll()"></th>
                                        <th>ID</th>
                                        <th>所属分类</th>
                                        <th>应用名称</th>
                                        <th>更新时间</th>
                                        <th>版本号</th>
                                        <th>浏览次数</th>
                                        <th>下载次数</th>
                                        <th>评分</th>
                                        <th>是否推荐</th>
                                        <th>操作</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <%
                                    If rs.eof or rs.bof Then
                                        echo ("<tr><td colspan='14'>暂无数据</td></tr>")
                                        die
                                    Else
                                        page.rs=rs
                                        page.PageSize=20
                                        page.setCount=-1
                                        page.PageID="p"
                                        i=0
                                        Do While Not page.EOF
                                        i=i+1
                                    %>
                                    <tr>
                                        <td><input type="checkbox" name="chkid" id="chkid" value="<%=page(0)%>"></td>
                                        <td><a href="?act=editapp&a_id=<%=page(0)%>"><%=page(0)%></a></td>
                                        <td><%If page(1)="1" Then%>软件<%Else%>游戏<%End If%></td>
                                        <td><%=page(2)%></td>
                                        <td><%=page(3)%></td>
                                        <td><%=page(4)%></td>
                                        <td><%=page(5)%></td>
                                        <td><%=page(6)%></td>
                                        <td><%=page(7)%></td>
                                        <td><%If page(8)="0" Then%><a href="?act=settop&a_id=<%=page(0)%>&istop=<%=page(8)%>">推荐</a><%Else%><a href="?act=settop&a_id=<%=page(0)%>&istop=<%=page(8)%>">不推荐</a><%End If%></td>
                                        <td><a href="?act=delapp&a_id=<%=page(0)%>">删除</a></td>
                                    </tr>
                                    <%
                                        page.MoveNext
                                        Loop
                                    End If
                                    %>
                                </tbody>
                            </table>
                            <ul class="pagination pagination-sm pull-right" id="pagination">
                                <%=page.Control(10,"{Ctrl}{State}")%><%Set page=nothing%>
                            </ul>
                            </form>
                            <%
                            ElseIf act="editapp" Then
                                If Not isNum(a_id) Then
                                    alertMsg "参数非法","manage-apps.asp"
                                End If
                            Set rs = mysql.exec("select * from app_apps where a_id="&a_id, 1)
                            %>
                            <table class="table table-bordered table-hover" id="editapp">
                            <form class="form-horizontal" action="?act=editapps&a_id=<%=rs("a_id")%>" method="post" name="editapps" role="form">
                                <tr>
                                    <td class="col-md-2"><label class="control-label" for="a_title">应用名称</label></td>
                                    <td class="col-md-10"><input type="text" id="a_title" value="<%=rs("a_title")%>" name="a_title" class="form-control" required autofocus></td>
                                </tr>
                                <tr>
                                    <td class="col-md-2"><label class="control-label" for="c_id">应用分类</label></td>
                                    <td class="col-md-10">
                                        <select class="form-control" name="c_id" id="c_id">
                                            <option value="1" <%If rs("c_id")="1" Then%>selected="true"<%End If%>>软件</option>
                                            <option value="2" <%If rs("c_id")="2" Then%>selected="true"<%End If%>>游戏</option>
                                        </select>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="col-md-2"><label class="control-label" for="a_seoname">SEO名称</label></td>
                                    <td class="col-md-10"><input type="text" id="a_seoname" value="<%=rs("a_seoname")%>" name="a_seoname" class="form-control" required autofocus></td>
                                </tr>
                                <tr>
                                    <td class="col-md-2"><label class="control-label" for="a_istop">是否推荐</label></td>
                                    <td class="col-md-10">
                                        <select class="form-control" name="a_istop" id="a_istop">
                                            <option value="0" <%If rs("a_istop")="0" Then%>selected="true"<%End If%>>推荐</option>
                                            <option value="1" <%If rs("a_istop")="1" Then%>selected="true"<%End If%>>不推荐</option>
                                        </select>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="col-md-2"><label class="control-label" for="a_uptime">更新时间</label></td>
                                    <td class="col-md-10"><input type="text" id="a_uptime" value="<%=FormatTime(now,7)%>" name="a_uptime" class="form-control" required autofocus></td>
                                </tr>
                                <tr>
                                    <td class="col-md-2"><label class="control-label" for="a_content">应用简介</label></td>
                                    <td class="col-md-10"><textarea id="contentedit" name="a_content"><%=rs("a_content")%></textarea></td>
                                </tr>
                                <tr>
                                    <td class="col-md-2"><label class="control-label" for="a_version">应用版本</label></td>
                                    <td class="col-md-10"><input type="text" id="a_version" value="<%=rs("a_version")%>" name="a_version" class="form-control" required autofocus></td>
                                </tr>
                                <tr>
                                    <td class="col-md-2"><label class="control-label" for="a_view">浏览次数</label></td>
                                    <td class="col-md-10"><input type="text" id="a_view" value="<%=rs("a_view")%>" name="a_view" class="form-control" required autofocus></td>
                                </tr>
                                <tr>
                                    <td class="col-md-2"><label class="control-label" for="a_downnum">下载次数</label></td>
                                    <td class="col-md-10"><input type="text" id="a_downnum" value="<%=rs("a_downnum")%>" name="a_downnum" class="form-control" required autofocus></td>
                                </tr>
                                <tr>
                                    <td class="col-md-2"><label class="control-label" for="a_tag">标签</label></td>
                                    <td class="col-md-10"><input type="text" id="a_tag" value="<%=rs("a_tag")%>" name="a_tag" class="form-control" required autofocus></td>
                                </tr>
                                <tr>
                                    <td class="col-md-2"><label class="control-label" for="a_vendor">开发厂商</label></td>
                                    <td class="col-md-10"><input type="text" id="a_vendor" value="<%=rs("a_vendor")%>" name="a_vendor" class="form-control" required autofocus></td>
                                </tr>
                                <tr>
                                    <td class="col-md-2"><label class="control-label" for="a_score">应用评分</label></td>
                                    <td class="col-md-10"><input type="text" id="a_score" value="<%=rs("a_score")%>" name="a_score" class="form-control" required autofocus></td>
                                </tr>
                                <tr>
                                    <td class="col-md-2"><label class="control-label" for="a_downlink">下载链接</label></td>
                                    <td class="col-md-10"><input type="text" id="a_downlink" value="<%=rs("a_downlink")%>" name="a_downlink" class="form-control" required autofocus></td>
                                </tr>
                                <tr>
                                    <td class="col-md-2"><label class="control-label" for="a_icon">应用图标</label></td>
                                    <td class="col-md-10"><input type="text" id="a_icon" value="<%=rs("a_icon")%>" name="a_icon" class="form-control" required autofocus></td>
                                </tr>
                                <tr>
                                    <td class="col-md-2"><label class="control-label" for="a_keyword">搜索关键字</label></td>
                                    <td class="col-md-10"><textarea id="a_keyword" name="a_keyword" class="form-control" rows="3" required autofocus><%=rs("a_keyword")%></textarea></td>
                                </tr>
                                <tr>
                                    <td class="col-md-2"><label class="control-label" for="a_description">搜索介绍关键字</label></td>
                                    <td class="col-md-10"><textarea id="a_description" name="a_description" class="form-control" rows="3" required autofocus><%=rs("a_description")%></textarea></td>
                                </tr>
                                <tr>
                                    <td class="col-md-2"><label class="control-label"></label></td>
                                    <td class="col-md-10"><button type="submit" class="btn btn-lg btn-primary btn-block">修改</button></td>
                                </tr>
                            </form>
                            </table>
                            <%End If%>
                            <table class="table table-bordered table-hover hidden" id="addtable">
                            <form class="form-horizontal" action="?act=addapp" method="post" name="addapp" role="form">
                                <tr>
                                    <td class="col-md-2"><label class="control-label" for="a_title">应用名称</label></td>
                                    <td class="col-md-10"><input type="text" id="a_title" value="" name="a_title" class="form-control" required autofocus></td>
                                </tr>
                                <tr>
                                    <td class="col-md-2"><label class="control-label" for="c_id">应用分类</label></td>
                                    <td class="col-md-10">
                                        <select class="form-control" name="c_id" id="c_id">
                                            <option value="1">软件</option>
                                            <option value="2">游戏</option>
                                        </select>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="col-md-2"><label class="control-label" for="a_seoname">SEO名称</label></td>
                                    <td class="col-md-10"><input type="text" id="a_seoname" value="" name="a_seoname" class="form-control" required autofocus></td>
                                </tr>
                                <tr>
                                    <td class="col-md-2"><label class="control-label" for="a_istop">是否推荐</label></td>
                                    <td class="col-md-10">
                                        <select class="form-control" name="a_istop" id="a_istop">
                                            <option value="0">推荐</option>
                                            <option value="1">不推荐</option>
                                        </select>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="col-md-2"><label class="control-label" for="a_uptime">更新时间</label></td>
                                    <td class="col-md-10"><input type="text" id="a_uptime" value="<%=FormatTime(now,7)%>" name="a_uptime" class="form-control" required autofocus></td>
                                </tr>
                                <tr>
                                    <td class="col-md-2"><label class="control-label" for="a_content">应用简介</label></td>
                                    <td class="col-md-10"><textarea id="contentedit" name="a_content"></textarea></td>
                                </tr>
                                <tr>
                                    <td class="col-md-2"><label class="control-label" for="a_version">应用版本</label></td>
                                    <td class="col-md-10"><input type="text" id="a_version" value="" name="a_version" class="form-control" required autofocus></td>
                                </tr>
                                <tr>
                                    <td class="col-md-2"><label class="control-label" for="a_view">浏览次数</label></td>
                                    <td class="col-md-10"><input type="text" id="a_view" value="" name="a_view" class="form-control" required autofocus></td>
                                </tr>
                                <tr>
                                    <td class="col-md-2"><label class="control-label" for="a_downnum">下载次数</label></td>
                                    <td class="col-md-10"><input type="text" id="a_downnum" value="" name="a_downnum" class="form-control" required autofocus></td>
                                </tr>
                                <tr>
                                    <td class="col-md-2"><label class="control-label" for="a_tag">标签</label></td>
                                    <td class="col-md-10"><input type="text" id="a_tag" value="" name="a_tag" class="form-control" required autofocus></td>
                                </tr>
                                <tr>
                                    <td class="col-md-2"><label class="control-label" for="a_vendor">开发厂商</label></td>
                                    <td class="col-md-10"><input type="text" id="a_vendor" value="" name="a_vendor" class="form-control" required autofocus></td>
                                </tr>
                                <tr>
                                    <td class="col-md-2"><label class="control-label" for="a_score">应用评分</label></td>
                                    <td class="col-md-10"><input type="text" id="a_score" value="" name="a_score" class="form-control" required autofocus></td>
                                </tr>
                                <tr>
                                    <td class="col-md-2"><label class="control-label" for="a_downlink">下载链接</label></td>
                                    <td class="col-md-10"><input type="text" id="a_downlink" value="" name="a_downlink" class="form-control" required autofocus></td>
                                </tr>
                                <tr>
                                    <td class="col-md-2"><label class="control-label" for="a_icon">应用图标</label></td>
                                    <td class="col-md-10"><input type="text" id="a_icon" value="" name="a_icon" class="form-control" required autofocus></td>
                                </tr>
                                <tr>
                                    <td class="col-md-2"><label class="control-label" for="a_keyword">搜索关键字</label></td>
                                    <td class="col-md-10"><textarea id="a_keyword" name="a_keyword" class="form-control" rows="3" required autofocus></textarea></td>
                                </tr>
                                <tr>
                                    <td class="col-md-2"><label class="control-label" for="a_description">搜索介绍关键字</label></td>
                                    <td class="col-md-10"><textarea id="a_description" name="a_description" class="form-control" rows="3" required autofocus></textarea></td>
                                </tr>
                                <tr>
                                    <td class="col-md-2"><label class="control-label"></label></td>
                                    <td class="col-md-10"><button type="submit" class="btn btn-lg btn-primary btn-block">添加</button></td>
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