<%@ page contentType="text/html;charset=UTF-8" language="java" %>
<html>
<head>
    <title>导入测试</title>
</head>
    <body>
        <form action="${pageContext.request.scheme}://${pageContext.request.serverName}:${pageContext.request.serverPort}${pageContext.request.contextPath}/uploadFile_file"
              method="post" enctype="multipart/form-data">
            <input id="file" name="file" type="file"/><br/><br/>　　
            <button type="submit"> 导 入</button>

            <a href="${pageContext.request.scheme}://${pageContext.request.serverName}:${pageContext.request.serverPort}${pageContext.request.contextPath}/excel_load">导出</a>
        </form>
    </body>
</html>
