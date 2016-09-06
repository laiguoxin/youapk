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
        <title>用户管理 - <%=SiteName%></title>
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
                <h4 class="page-header">用户管理</h1>
            </div>
        </div>
        <div class="row">
            <div class="col-lg-12">
                <div class="panel panel-default">
                    <form action="?act=delusers" method="post" name="delall" id="delall" role="form">
                    <div class="panel-heading"><i class="glyphicon glyphicon-user"></i> 用户列表
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
                            <%
                            If act="" Then
                                Call mysql.exec("", -1)
                                sql="select u_id,u_username,u_password,u_regtime,u_loginip,u_level,u_mobile,u_qq,u_sex,u_age,u_logintime from app_users order by u_id desc"
                                rs.open sql,conn,1,1
                            %>
                            <table class="table table-bordered table-hover" id="table">
                                <thead>
                                    <tr class="info">
                                        <th><input type="checkbox" name="checkbox" id="checkbox" onClick="javascript:SelectAll()"></th>
                                        <th>ID</th>
                                        <th>用户名</th>
                                        <th>登陆密码</th>
                                        <th>注册时间</th>
                                        <th>上次登录IP</th>
                                        <th>用户权限</th>
                                        <th>手机号码</th>
                                        <th>绑定QQ</th>
                                        <th>性别</th>
                                        <th>年龄</th>
                                        <th>上次登录时间</th>
                                        <th>操作</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <%
                                    If rs.eof or rs.bof Then
                                        echo ("<tr><td colspan='13'>暂无数据</td></tr>")
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
                                        <td><a href="?act=edituser&u_id=<%=page(0)%>"><%=page(0)%></a></td>
                                        <td><%=page(1)%></td>
                                        <td><%=des.Decode(page(2))%></td>
                                        <td><%=page(3)%></td>
                                        <td><%=page(4)%></td>
                                        <td><%If page(5)="0" Then%>管理员<%Else%>普通会员<%End If%></td>
                                        <td><%=page(6)%></td>
                                        <td><%=page(7)%></td>
                                        <td><%If page(8)="0" Then%>男<%Else%>女<%End If%></td>
                                        <td><%=page(9)%></td>
                                        <td><%=page(10)%></td>
                                        <td><a href="?act=deluser&u_id=<%=page(0)%>">删除</a></td>
                                    </tr>
                                    <%
                                        page.MoveNext
                                        Loop
                                    End If
                                    %>
                                </tbody>
                            </table>
                            <ul class="pagination pagination-sm pull-right">
                                <%=page.Control(10,"{Ctrl}{State}")%><%Set page=nothing%>
                            </ul>
                            <%
                            ElseIf act="edituser" Then
                                If Not isNum(u_id) Then
                                    alertMsg "参数非法","manage-users.asp"
                                End If
                            Set rs = mysql.exec("select * from app_users where u_id="&u_id, 1)
                            %>
                            <table class="table table-bordered table-hover" id="edituser">
                            <form class="form-horizontal" action="?act=editusers&u_id=<%=rs("u_id")%>" method="post" name="editusers" role="form">
                                <tr>
                                    <td><label class="control-label" for="u_username">用户名</label></td>
                                    <td><input type="text" id="u_username" value="<%=rs("u_username")%>" name="u_username" class="form-control" required autofocus></td>
                                </tr>
                                <tr>
                                    <td><label class="control-label" for="u_password">密码</label></td>
                                    <td><input type="password" id="u_password" value="" name="u_password" class="form-control" placeholder="必须输入旧密码或输入新密码" required></td>
                                </tr>
                                <tr>
                                    <td><label class="control-label" for="u_email">邮箱</label></td>
                                    <td><input type="text" id="u_email" value="<%=rs("u_email")%>" name="u_email" class="form-control" required></td>
                                </tr>
                                <tr>
                                    <td><label class="control-label" for="u_level">用户级别</label></td>
                                    <td><select class="form-control" name="u_level" id="u_level"><option value="0" <%If rs("u_level")="0" Then%>selected="true"<%End If%>>管理员</option><option value="1" <%If rs("u_level")="1" Then%>selected="true"<%End If%>>普通会员</option></select></td>
                                </tr>
                                <tr>
                                    <td><label class="control-label" for="u_avatar">头像</label></td>
                                    <td><input type="text" id="u_avatar" value="<%=rs("u_avatar")%>" name="u_avatar" class="form-control" required></td>
                                </tr>
                                <tr>
                                    <td><label class="control-label" for="u_mobile">手机号码</label></td>
                                    <td><input type="number" size="11" id="u_mobile" value="<%=rs("u_mobile")%>" name="u_mobile" class="form-control" required></td>
                                </tr>
                                <tr>
                                    <td><label class="control-label" for="u_mobile">QQ</label></td>
                                    <td><input type="number" id="u_qq" value="<%=rs("u_qq")%>" name="u_qq" class="form-control" required></td>
                                </tr>
                                <tr>
                                    <td><label class="control-label" for="u_sex">性别</label></td>
                                    <td><select class="form-control" name="u_sex" id="u_sex"><option value="0" <%If rs("u_level")="0" Then%>selected="true"<%End If%>>男</option><option value="1" <%If rs("u_sex")="1" Then%>selected="true"<%End If%>>女</option></select></td>
                                </tr>
                                <tr>
                                    <td><label class="control-label" for="u_age">年龄</label></td>
                                    <td><input type="number" id="u_age" value="<%=rs("u_age")%>" name="u_age" class="form-control" required></td>
                                </tr>
                                <tr>
                                    <td><label class="control-label" for="u_question">密保问题</label></td>
                                    <td>
                                        <select class="form-control" name="u_question" id="u_question">
                                            <option value="0" <%If rs("u_question")="0" Then%>selected="true"<%End If%>>您的出生地是什么？</option>
                                            <option value="1" <%If rs("u_question")="1" Then%>selected="true"<%End If%>>您父亲的姓名是什么？</option>
                                            <option value="2" <%If rs("u_question")="2" Then%>selected="true"<%End If%>>您母亲的姓名是什么？</option>
                                        </select>
                                    </td>
                                </tr>
                                <tr>
                                    <td><label class="control-label" for="u_answer">密保答案</label></td>
                                    <td><input type="text" id="u_answer" value="<%=rs("u_answer")%>" name="u_answer" class="form-control" required></td>
                                </tr>
                                <tr>
                                    <td><label class="control-label"></label></td>
                                    <td><button type="submit" class="btn btn-lg btn-primary btn-block">修改</button></td>
                                </tr>
                            </form>
                            </table>
                            <%End If%>
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
<%Call clear_mysql()%>