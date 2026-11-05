<%@ page language="java" contentType="application/json; charset=UTF-8" pageEncoding="UTF-8" %>
<%
    response.setStatus(404);
    response.setContentType("application/json;charset=UTF-8");
    response.setHeader("Cache-Control", "private, no-store");
    out.print("{\"code\":404,\"errorCode\":\"OFFICE_PREVIEW_UNSUPPORTED\","
            + "\"msg\":\"旧版 Office 转换入口未启用\"}");
%>
