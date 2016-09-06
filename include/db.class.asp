<%
'//数据库操作类
Class dbClass
    Public connstr,isconn
    Private Sub Class_Initialize()
        isconn=False
    End Sub

    Sub conn_open()
        set conn=server.createobject("adodb.connection")
        set rs=server.createobject("adodb.recordset")
        conn.open connstr
        isconn=True
	End Sub

    Function exec(sql,stype)
        If not isconn=true Then Call conn_open()
        Select Case stype
            Case 0
                conn.execute(sql)
            Case 1
                Set exec=server.CreateObject("adodb.recordset")
                exec.open sql,conn,1,1
        End Select
    End Function
End Class

Dim conn,rs,mysql
Set mysql = new dbClass
    mysql.connstr = "DRIVER={SQLite3 ODBC Driver};Database=" &Server.MapPath(SiteDBPath)


Sub clear_mysql()
    If isobject(rs) Then Set rs=nothing
    If mysql.isconn Then
        conn.close
        Set conn=nothing
    End If
    Set mysql=nothing
End Sub
%>