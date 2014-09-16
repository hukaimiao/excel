<%@ page language="java" import="java.util.*" pageEncoding="UTF-8"%>
<%@ taglib uri="http://java.sun.com/jsp/jstl/core" prefix="c"%>
<%@ taglib uri="http://java.sun.com/jsp/jstl/fmt" prefix="fmt" %> 
<%
	String path = request.getContextPath();
	String basePath = request.getScheme() + "://"
			+ request.getServerName() + ":" + request.getServerPort()
			+ path + "/";
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<base href="<%=basePath%>">

<title>My JSP 'inputdata.jsp' starting page</title>

<meta http-equiv="pragma" content="no-cache">
<meta http-equiv="cache-control" content="no-cache">
<meta http-equiv="expires" content="0">
<meta http-equiv="keywords" content="keyword1,keyword2,keyword3">
<meta http-equiv="description" content="This is my page">
<!--
	<link rel="stylesheet" type="text/css" href="styles.css">
	-->

</head>

<body>
	<table>
		<tr>
			<td>学号</td>
			<td>姓名</td>
			<td>年龄</td>
			<td>性别</td>
			<td>出生日期</td>
		</tr>
		<c:forEach var="s" items="${list}">
		<tr>
			<td>${s.id}</td>
			<td>${s.name }</td>
			<td>${s.age }</td>
			<td>${s.sex }</td>
			<td><fmt:formatDate value="${s.birthday }" pattern="yyyy-MM-dd" /></td>
		</tr>
		</c:forEach>
	</table>
</body>
</html>
