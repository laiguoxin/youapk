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
        <title>评论管理 - <%=SiteName%></title>
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
                <h4 class="page-header">评论管理</h1>
            </div>
        </div>
        <div class="row">
            <div class="col-lg-12">
                <div class="panel panel-default">
                    <form action="?act=delcomments" method="post" name="delall" id="delall" role="form">
                    <div class="panel-heading"><i class="glyphicon glyphicon-comment"></i> 评论列表
                        <div class="pull-right">
                            <div class="btn-group">
                                <button type="submit" class="btn btn-danger btn-xs" id="del">
                                    <span class="glyphicon glyphicon-remove" aria-hidden="true"></span>删除
                                </button>
                            </div>
                        </div>
                    </div>
                    <div class="panel-body">
                        <div class="table-responsive">
                            <table class="table table-bordered table-hover" id="table">
                                <thead>
                                    <tr class="info">
                                        <th><input type="checkbox" name="checkbox" id="checkbox" onClick="javascript:SelectAll()"></th>
                                        <th>ID</th>
                                        <th>评论用户</th>
                                        <th>所属应用</th>
                                        <th>评论内容</th>
                                        <th>评论时间</th>
                                        <th>评论IP</th>
                                        <th>评论状态</th>
                                        <th>赞</th>
                                        <th>回复数</th>
                                        <th>操作</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <%
                                    Call mysql.exec("", -1)
                                    sql="select cm_id,u_id,a_id,cm_content,cm_time,cm_ip,cm_status,cm_top,cm_replynum from app_comments order by cm_id desc"
                                    rs.open sql,conn,1,1
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
                                        <td><%=page(0)%></td>
                                        <td><%Call getusername%></td>
                                        <td><%Call GetCMIDName%></td>
                                        <td><%=delHtml(left(page(3),15))%></td>
                                        <td><%=page(4)%></td>
                                        <td><%=page(5)%></td>
                                        <td><%If page(6)="0" Then%><a href="?act=setreview&cm_id=<%=page(0)%>&isreview=<%=page(6)%>">已审核</a><%Else%><a href="?act=setreview&cm_id=<%=page(0)%>&isreview=<%=page(6)%>">未审核</a><%End If%></td>
                                        <td><%=page(7)%></td>
                                        <td><%=page(8)%></td>
                                        <td><a href="?act=delcomment&cm_id=<%=page(0)%>">删除</a></td>
                                    </tr>
                                    <%
                                        page.MoveNext
                                        Loop
                                    End If
                                    %>
                                </tbody>
                            </table>
                            <ul class="pagination pagination-sm pull-right">
                                <%=page.Control(10,"{Ctrl}{State}")%>
                            </ul>
                        </div>
                    </div>
                    </form>
                </div>
            </div>
        </div>
    </div>
</div>
</body>
</html>
<%
Set page=nothing
Call clear_mysql()
%>