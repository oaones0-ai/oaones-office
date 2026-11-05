<%@ page language="java" contentType="application/json; charset=UTF-8" pageEncoding="UTF-8" %>
<%@ page import="com.alibaba.fastjson.JSONObject" %>
<%@ page import="com.customization.oaones.officeunit.common.OfficeFileAccessSupport" %>
<%@ page import="com.customization.oaones.officeunit.common.OfficePreviewException" %>
<%@ page import="com.customization.oaones.officeunit.common.OfficeUnitConfig" %>
<%@ page import="com.customization.oaones.officeunit.common.OfficeUnitLog" %>
<%@ page import="com.customization.oaones.officeunit.common.OfficeUnitUtil" %>
<%@ page import="com.customization.oaones.officeunit.preview.OfficePreviewTicketService" %>
<%@ page import="weaver.file.ImageFileManager" %>
<%@ page import="weaver.hrm.HrmUserVarify" %>
<%@ page import="weaver.hrm.User" %>
<%@ page import="java.io.InputStream" %>
<%@ page import="java.io.OutputStream" %>
<%
    response.setCharacterEncoding("UTF-8");
    response.setContentType("application/json;charset=UTF-8");
    response.setHeader("Cache-Control", "private, no-store, no-cache, must-revalidate");
    response.setHeader("Pragma", "no-cache");
    response.setHeader("Expires", "0");
    response.setHeader("X-Content-Type-Options", "nosniff");
    response.setHeader("Referrer-Policy", "no-referrer");

    User user = HrmUserVarify.getUser(request, response);
    if (user == null) {
        response.setStatus(401);
        out.print("{\"code\":401,\"errorCode\":\"OFFICE_PREVIEW_FORBIDDEN\","
                + "\"msg\":\"未登录或会话已失效\"}");
        return;
    }
    if (!"GET".equalsIgnoreCase(request.getMethod())) {
        response.setStatus(405);
        out.print("{\"code\":405,\"msg\":\"请求方法不支持\"}");
        return;
    }

    InputStream input = null;
    try {
        OfficePreviewTicketService ticketService = new OfficePreviewTicketService();
        OfficePreviewTicketService.Claims claims = ticketService.verify(
                request.getParameter("token"), user.getUID());
        OfficeUnitConfig.Config config = OfficeUnitConfig.load();
        if (!config.enabled || !config.browserMode()) {
            throw new OfficePreviewException("OFFICE_PREVIEW_UNSUPPORTED",
                    "Office 浏览器预览当前未启用", 503);
        }
        OfficeFileAccessSupport.FileInfo file = OfficeFileAccessSupport.resolveBound(
                claims.docId, claims.imageFileId, claims.versionId, user, request, config);
        if (!claims.extension.equals(file.extension)
                || claims.fileSize != file.fileSize
                || claims.fileSize > config.maxInputBytes) {
            throw new OfficePreviewException("OFFICE_PREVIEW_FILE_NOT_FOUND",
                    "文档文件已变化，请重新打开预览", 409);
        }

        input = ImageFileManager.getInputStreamById(file.imageFileId);
        if (input == null) {
            throw OfficePreviewException.fileNotFound();
        }
        response.reset();
        response.setStatus(200);
        response.setContentType("application/octet-stream");
        response.setHeader("Content-Disposition",
                OfficeUnitUtil.contentDisposition(file.fileName, file.extension));
        response.setHeader("Content-Length", String.valueOf(file.fileSize));
        response.setHeader("Cache-Control", "private, no-store, no-cache, must-revalidate");
        response.setHeader("Pragma", "no-cache");
        response.setHeader("Expires", "0");
        response.setHeader("X-Content-Type-Options", "nosniff");
        response.setHeader("Referrer-Policy", "no-referrer");
        response.setHeader("X-Download-Options", "noopen");

        OutputStream output = response.getOutputStream();
        byte[] buffer = new byte[16384];
        long total = 0L;
        int count;
        while ((count = input.read(buffer)) >= 0) {
            total += count;
            if (total > config.maxInputBytes || total > claims.fileSize) {
                throw new IllegalStateException("文件大小与预览票据不一致");
            }
            output.write(buffer, 0, count);
        }
        if (total != claims.fileSize) {
            throw new IllegalStateException("文件大小与预览票据不一致");
        }
        output.flush();
        OfficeUnitLog.info("文件流完成 docId=" + file.docId
                + " imageFileId=" + file.imageFileId
                + " versionId=" + file.versionId
                + " userId=" + user.getUID());
    } catch (OfficePreviewException e) {
        if (!response.isCommitted()) {
            response.reset();
            response.setStatus(e.getHttpStatus());
            response.setContentType("application/json;charset=UTF-8");
            response.setHeader("Cache-Control", "private, no-store");
            response.setHeader("X-Content-Type-Options", "nosniff");
            JSONObject error = new JSONObject();
            error.put("code", Integer.valueOf(-1));
            error.put("errorCode", e.getCode());
            error.put("msg", e.getMessage());
            out.print(error.toJSONString());
        }
    } catch (Throwable t) {
        OfficeUnitLog.error("文件流处理失败 userId=" + user.getUID(), t);
        if (!response.isCommitted()) {
            response.reset();
            response.setStatus(500);
            response.setContentType("application/json;charset=UTF-8");
            response.setHeader("Cache-Control", "private, no-store");
            response.setHeader("X-Content-Type-Options", "nosniff");
            out.print("{\"code\":-1,\"errorCode\":\"OFFICE_PREVIEW_FILE_FETCH_FAILED\","
                    + "\"msg\":\"文档加载失败，请稍后重试\"}");
        }
    } finally {
        OfficeUnitUtil.closeQuietly(input);
    }
%>
