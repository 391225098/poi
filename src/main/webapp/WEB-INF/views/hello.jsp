<%--
  Created by IntelliJ IDEA.
  User: wb252654
  Date: 17-9-6
  Time: 上午10:15
  To change this template use File | Settings | File Templates.
--%>
<%@ page contentType="text/html;charset=UTF-8" language="java" %>
<html>
<head>
    <title></title>
</head>
<body>
<form action="/test/upload" method="post" enctype="multipart/form-data">
    <input type="file" name="file">
    <button type="submit">submit</button>
</form>
</body>
</html>