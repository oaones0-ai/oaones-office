<%@ page language="java" contentType="application/json; charset=UTF-8" pageEncoding="UTF-8" %>
<%@ page import="com.alibaba.fastjson.JSONObject" %>
<%@ page import="com.customization.oaones.officeunit.common.OfficePreviewException" %>
<%@ page import="com.customization.oaones.officeunit.common.OfficeUnitLog" %>
<%@ page import="com.customization.oaones.officeunit.preview.OfficePreviewController" %>
<%@ page import="com.customization.oaones.officeunit.preview.OfficeUnitAdminController" %>
<%@ page import="weaver.hrm.HrmUserVarify" %>
<%@ page import="weaver.hrm.User" %>
<%
    request.setCharacterEncoding("UTF-8");
    response.setCharacterEncoding("UTF-8");
    response.setContentType("application/json;charset=UTF-8");
    response.setHeader("Cache-Control", "private, no-store, no-cache, must-revalidate");
    response.setHeader("Pragma", "no-cache");
    response.setHeader("X-Content-Type-Options", "nosniff");
    response.setHeader("Referrer-Policy", "no-referrer");

    User user = HrmUserVarify.getUser(request, response);
    if (user == null) {
        response.setStatus(401);
        out.print("{\"code\":401,\"errorCode\":\"OFFICE_PREVIEW_FORBIDDEN\","
                + "\"msg\":\"未登录或会话已失效\"}");
        return;
    }

    String action = request.getParameter("action");
    JSONObject result;
    try {
        boolean writeAction = "public.officeunit.preview.prepare".equals(action)
                || "admin.save".equals(action);
        if (writeAction) {
            if (!"POST".equalsIgnoreCase(request.getMethod())
                    || !"XMLHttpRequest".equalsIgnoreCase(
                    request.getHeader("X-Requested-With"))) {
                response.setStatus(403);
                result = new JSONObject();
                result.put("code", Integer.valueOf(403));
                result.put("errorCode", "OFFICE_PREVIEW_FORBIDDEN");
                result.put("msg", "非法请求");
                out.print(result.toJSONString());
                return;
            }
        } else if (!"GET".equalsIgnoreCase(request.getMethod())) {
            response.setStatus(405);
            result = new JSONObject();
            result.put("code", Integer.valueOf(405));
            result.put("msg", "请求方法不支持");
            out.print(result.toJSONString());
            return;
        }

        if ("public.officeunit.preview.prepare".equals(action)) {
            result = new OfficePreviewController().prepare(request, user);
        } else if (action != null && action.startsWith("admin.")) {
            result = new OfficeUnitAdminController().handle(action, request, user.getUID());
        } else {
            response.setStatus(400);
            result = new JSONObject();
            result.put("code", Integer.valueOf(400));
            result.put("msg", "未知 action");
        }
    } catch (OfficePreviewException e) {
        response.setStatus(e.getHttpStatus());
        result = new JSONObject();
        result.put("code", Integer.valueOf(-1));
        result.put("errorCode", e.getCode());
        result.put("msg", e.getMessage());
    } catch (SecurityException e) {
        response.setStatus(403);
        result = new JSONObject();
        result.put("code", Integer.valueOf(403));
        result.put("errorCode", "OFFICE_PREVIEW_FORBIDDEN");
        result.put("msg", e.getMessage());
    } catch (IllegalArgumentException e) {
        response.setStatus(400);
        result = new JSONObject();
        result.put("code", Integer.valueOf(400));
        result.put("msg", e.getMessage());
    } catch (Throwable t) {
        OfficeUnitLog.error("OfficeUnit API 处理失败 action=" + action, t);
        response.setStatus(500);
        result = new JSONObject();
        result.put("code", Integer.valueOf(-1));
        result.put("errorCode", "OFFICE_PREVIEW_PARSE_FAILED");
        result.put("msg", "Office 预览请求处理失败，请稍后重试");
    }
    out.print(result.toJSONString());
%>
