<!--#include file="include/admin.asp"-->
<%Call Chklogin()%>
<!DOCTYPE html>
<html>
    <head>
        <meta charset="UTF-8">
        <meta http-equiv="X-UA-Compatible" content="IE=edge, chrome=1">
        <meta name="renderer" content="webkit">
        <meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no">
        <title>用户中心 - <%=SiteName%></title>
        <link href="<%=SiteUrl%><%=PublicTheme%>/css/bootstrap.min.css" rel="stylesheet">
        <link href="<%=SiteUrl%><%=AdminTheme%>/css/admin.css" rel="stylesheet">
        <link href="<%=SiteUrl%><%=PublicTheme%>/css/Validator.min.css" rel="stylesheet">
        <script src="<%=SiteUrl%><%=PublicTheme%>/js/jquery.min.js"></script>
        <script src="<%=SiteUrl%><%=PublicTheme%>/js/bootstrap.min.js"></script>
        <script src="<%=SiteUrl%><%=PublicTheme%>/js/Validator.min.js"></script>
        <script src="<%=SiteUrl%><%=PublicTheme%>/js/Validator.zh_CN.js"></script>
</head>
<body>
<div id="wrapper">
    <nav class="navbar navbar-inverse navbar-static-top" role="navigation">
        <div class="navbar-header">
            <a class="navbar-brand" href="../login.asp"><%=SiteName%></a><p class="navbar-text"><a href="../index.asp">回到首页</a></p>
        </div>
    </nav>
    <div class="container">
        <div class="row">
            <div class="col-md-4 col-md-offset-4">
                <%If act="" Then%>
                <div class="login-panel panel panel-default">
                    <ul class="nav nav-tabs">
                        <li class="active"><a href="#login" data-toggle="tab" class="login-paneltab">登录</a></li>
                        <li><a href="#register" data-toggle="tab">注册</a></li>
                        <li><a href="#forget" data-toggle="tab">找回密码</a></li>
                    </ul>
                    <div class="tab-content">
                        <div class="tab-pane fade in active" id="login">
                            <div class="panel-body">
                                <form class="form-signin" action="include/admin.asp?act=login" method="post" name="login" id="Loginform" role="form">
                                    <fieldset>
                                        <div class="form-group">
                                            <input type="text" name="u_username" id="u_username" class="form-control" placeholder="账号">
                                        </div>
                                        <div class="form-group">
                                            <input type="password" name="u_password" id="u_password" class="form-control" placeholder="密码">
                                        </div>
                                        <div class="form-group">
                                            <input type="checkbox" name="rememberme" id="rememberme">记住我
                                        </div>
                                        <div class="form-group">
                                            <button class="btn btn-lg btn-primary btn-block" type="submit">登录</button>
                                        </div>
                                    </fieldset>
                                </form>
                            </div>
                        </div>
                        <div class="tab-pane fade" id="register">
                            <div class="panel-body">
                                <%Call isAllowReg%>
                                <form class="form-signin" action="include/admin.asp?act=register" method="post" name="register" id="Regform" role="form">
                                    <fieldset>
                                        <div class="form-group">
                                            <input type="text" name="u_username" id="u_username" class="form-control" placeholder="用户名">
                                        </div>
                                        <div class="form-group">
                                            <input type="password" name="u_password" id="u_password" class="form-control" placeholder="密码">
                                        </div>
                                        <div class="form-group">
                                            <input type="password" name="confirmPassword" id="confirmPassword" class="form-control" placeholder="确认密码">
                                        </div>
                                        <div class="form-group">
                                            <input type="text" name="u_email" id="u_email" class="form-control" placeholder="邮箱">
                                        </div>
                                        <div class="form-group">
                                            <select class="form-control" name="u_sex" id="u_sex">
                                                <option value="0">男</option>
                                                <option value="1">女</option>
                                            </select>
                                        </div>
                                        <div class="form-group">
                                            <select class="form-control" name="u_question" id="u_question">
                                                <option value="0">您的出生地是什么？</option>
                                                <option value="1">您父亲的姓名是什么？</option>
                                                <option value="2">您母亲的姓名是什么？</option>
                                            </select>
                                        </div>
                                        <div class="form-group">
                                            <input type="text" name="u_answer" id="u_answer" class="form-control" placeholder="密保答案">
                                        </div>
                                        <div class="form-group">
                                            <input type="checkbox" name="readme" id="readme">我同意接受本站的<a href="../copyright.html">版权声明</a>
                                        </div>
                                        <div class="form-group">
                                            <button class="btn btn-lg btn-primary btn-block" type="submit">确认注册</button>
                                        </div>
                                    </fieldset>
                                </form>
                            </div>
                        </div>
                        <div class="tab-pane fade" id="forget">
                            <div class="panel-body">
                                <form class="form-signin" action="?act=getverify" method="post" name="forget" id="Forgetform" role="form">
                                    <fieldset>
                                        <div class="form-group">
                                            <input type="text" name="u_username" id="u_username" class="form-control" placeholder="用户名">
                                        </div>
                                        <div class="form-group">
                                            <input type="text" name="u_email" id="u_email" class="form-control" placeholder="邮箱">
                                        </div>
                                        <div class="form-group">
                                            <button class="btn btn-lg btn-primary btn-block" type="submit">找回密码</button>
                                        </div>
                                    </fieldset>
                                </form>
                            </div>
                        </div>
                    </div>
                </div>
                <%
                ElseIf act="getverify" Then
                    Call CheckPostUrl()
                    If isNul(u_username) or isNul(u_email) Then
                        alertMsg "信息不完整","javascript:history.back(-1)"
                    Else
                    Set rs = mysql.exec("select u_username,u_email,u_question from app_users where u_username='"&u_username&"' and u_email='"&u_email&"'", 1)
                %>
                <div class="login-panel panel panel-default">
                    <ul class="nav nav-tabs">
                        <li class="active"><a class="login-paneltab">找回密码</a></li>
                    </ul>
                    <div class="tab-content">
                        <div class="panel-body">
                            <%If rs.bof and rs.eof Then%>
                            <h5 class='text-center'>用户名或者邮箱错误</h5>
                            <%Else%>
                            <form class="form-signin" action="include/admin.asp?act=forget" method="post" name="getverify" id="Forgetform" role="form">
                                <fieldset>
                                    <div class="form-group">
                                        <select class="form-control" name="u_question" id="u_question">
                                            <%If rs("u_question")="0" Then%>
                                                <option value="0" selected="true">您的出生地是什么？</option>
                                            <%ElseIf rs("u_question")="1" Then%>
                                                <option value="1" selected="true">您父亲的姓名是什么？</option>
                                            <%ElseIf rs("u_question")="2" Then%>
                                                <option value="2" selected="true">您母亲的姓名是什么？</option>
                                            <%End If%>
                                        </select>
                                    </div>
                                    <div class="form-group">
                                        <input type="text" id="u_answer" value="" name="u_answer" class="form-control" required>
                                    </div>
                                    <div class="form-group">
                                        <button class="btn btn-lg btn-primary btn-block" type="submit">验证问题</button>
                                     </div>
                                    </fieldset>
                                </form>
                                <%End If%>
                            </div>
                    </div>
                </div>
                    <%End If%>
                <%End If%>
            </div>
        </div>
    </div>
</div>
<script type="text/javascript">
$(document).ready(function(){
    $("#rememberme").click(function(){
    if(this.checked){
        $("input[name='rememberme']").prop('checked', true)
        $("input[name='rememberme']").prop('value', 1)}
    else{    
        $("input[name='rememberme']").prop('checked', false)
        $("input[name='rememberme']").prop('value', 0)}
    })
});

$(document).ready(function() {
    $('#Loginform').bootstrapValidator({
        feedbackIcons: {
            valid: 'glyphicon glyphicon-ok',
            invalid: 'glyphicon glyphicon-remove',
            validating: 'glyphicon glyphicon-refresh'
        },
        fields: {
            u_username: {
                message: '用户名无效',
                validators: {
                    notEmpty: {
                        message: '用户名不能为空'
                    },
                    stringLength: {
                        min: 5,
                        max: 16,
                        message: '用户名必须大于5，小于16个字'
                    },
                    regexp: {
                        regexp: /^[a-zA-Z0-9_\.]+$/,
                        message: '用户名只能由字母、数字、点和下划线组成'
                    }
                }
            },
            u_password: {
                validators: {
                    notEmpty: {
                        message: '密码不能为空'
                    },
                    stringLength: {
                        min: 8,
                        max: 16,
                        message: '密码必须大于8，小于16个字'
                    }
                }
            }
        }
    });
});

$(document).ready(function() {
    $('#Regform').bootstrapValidator({
        feedbackIcons: {
            valid: 'glyphicon glyphicon-ok',
            invalid: 'glyphicon glyphicon-remove',
            validating: 'glyphicon glyphicon-refresh'
        },
        fields: {
            u_username: {
                message: '用户名无效',
                validators: {
                    notEmpty: {
                        message: '用户名不能为空'
                    },
                    stringLength: {
                        min: 5,
                        max: 16,
                        message: '用户名必须大于5，小于16个字'
                    },
                    regexp: {
                        regexp: /^[a-zA-Z0-9_\.]+$/,
                        message: '用户名只能由字母、数字、点和下划线组成'
                    },
                    different: {
                        field: 'password',
                        message: '用户名和密码不能相同'
                    }
                }
            },
            u_email: {
                validators: {
                    notEmpty: {
                        message: '邮件地址不能为空'
                    },
                    emailAddress: {
                        message: '输入不是有效的邮件地址'
                    }
                }
            },
            u_password: {
                validators: {
                    notEmpty: {
                        message: '密码不能为空'
                    },
                    stringLength: {
                        min: 8,
                        max: 16,
                        message: '密码必须大于8，小于16个字'
                    },
                    identical: {
                        field: 'confirmPassword',
                        message: '两次密码不一致'
                    },
                    different: {
                        field: 'u_username',
                        message: '用户名和密码不能相同'
                    }
                }
            },
            confirmPassword: {
                validators: {
                    notEmpty: {
                        message: '密码不能为空'
                    },
                    stringLength: {
                        min: 8,
                        max: 16,
                        message: '密码必须大于8，小于16个字'
                    },
                    identical: {
                        field: 'u_password',
                        message: '两次密码不一致'
                    },
                    different: {
                        field: 'u_username',
                        message: '用户名和密码不能相同'
                    }
                }
            },
            u_answer: {
                validators: {
                    notEmpty: {
                        message: '密保答案不能为空'
                    },
                    stringLength: {
                        min: 2,
                        message: '密保答案不能小于2个字'
                    }
                }
            },
            readme: {
                validators: {
                    notEmpty: {
                        message: '请先仔细阅读本站协议'
                    }
                }
            }
        }
    });
});
$(document).ready(function() {
    $('#Forgetform').bootstrapValidator({
        feedbackIcons: {
            valid: 'glyphicon glyphicon-ok',
            invalid: 'glyphicon glyphicon-remove',
            validating: 'glyphicon glyphicon-refresh'
        },
        fields: {
            u_username: {
                message: '用户名无效',
                validators: {
                    notEmpty: {
                        message: '用户名不能为空'
                    },
                    stringLength: {
                        min: 5,
                        max: 16,
                        message: '用户名必须大于5，小于16个字'
                    },
                    regexp: {
                        regexp: /^[a-zA-Z0-9_\.]+$/,
                        message: '用户名只能由字母、数字、点和下划线组成'
                    }
                }
            },
            u_email: {
                validators: {
                    notEmpty: {
                        message: '邮件地址不能为空'
                    },
                    emailAddress: {
                        message: '输入不是有效的邮件地址'
                    }
                }
            },
            u_answer: {
                validators: {
                    notEmpty: {
                        message: '密保答案不能为空'
                    },
                    stringLength: {
                        min: 2,
                        message: '密保答案不能小于2个字'
                    }
                }
            }
        }
    });
});
</script>
</body>
</html>