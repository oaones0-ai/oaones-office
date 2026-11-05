package com.customization.oaones.officeunit.preview;

import com.alibaba.fastjson.JSONObject;
import com.customization.oaones.officeunit.common.OfficeFileAccessSupport;
import com.customization.oaones.officeunit.common.OfficePreviewException;
import com.customization.oaones.officeunit.common.OfficeUnitConfig;
import com.customization.oaones.officeunit.common.OfficeUnitLog;
import com.customization.oaones.officeunit.common.OfficeUnitUtil;
import weaver.file.ImageFileManager;
import weaver.hrm.HrmUserVarify;
import weaver.hrm.User;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.InputStream;
import java.io.OutputStream;

/**
 * Streams a permission-checked Office file before the JSP container obtains its
 * text writer. This is mapped by OfficePreviewFilter for preview/file.jsp.
 */
public final class OfficePreviewFileStreamer {

    private OfficePreviewFileStreamer() {
    }

    public static boolean isRequest(HttpServletRequest request) {
        String path = request == null ? "" : request.getRequestURI();
        return path != null && (path.endsWith("/oaones/officeunit/preview/file.jsp")
                || path.endsWith("/oaones/officeunit/preview/file"));
    }

    public static void handle(HttpServletRequest request, HttpServletResponse response) {
        setNoStoreHeaders(response);
        User user = HrmUserVarify.getUser(request, response);
        if (user == null) {
            writeError(response, 401, "OFFICE_PREVIEW_FORBIDDEN",
                    "未登录或会话已失效");
            return;
        }
        if (!"GET".equalsIgnoreCase(request.getMethod())) {
            writeError(response, 405, "OFFICE_PREVIEW_UNSUPPORTED", "请求方法不支持");
            return;
        }

        InputStream input = null;
        try {
            OfficePreviewTicketService service = new OfficePreviewTicketService();
            OfficePreviewTicketService.Claims claims = service.verify(
                    request.getParameter("token"), user.getUID());
            OfficeUnitConfig.Config config = OfficeUnitConfig.load();
            if (!config.enabled || !config.browserMode()) {
                throw new OfficePreviewException("OFFICE_PREVIEW_UNSUPPORTED",
                        "Office 浏览器预览当前未启用", 503);
            }
            OfficeFileAccessSupport.FileInfo file = OfficeFileAccessSupport.resolveBound(
                    claims.docId, claims.imageFileId, claims.versionId,
                    user, request, config);
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
            setNoStoreHeaders(response);
            response.setHeader("X-Download-Options", "noopen");

            OutputStream output = response.getOutputStream();
            byte[] buffer = new byte[16384];
            long total = 0L;
            int count;
            while ((count = input.read(buffer)) >= 0) {
                total += count;
                if (total > config.maxInputBytes || total > claims.fileSize) {
                    throw new IllegalStateException("stream size exceeds ticket");
                }
                output.write(buffer, 0, count);
            }
            if (total != claims.fileSize) {
                throw new IllegalStateException("stream size differs from ticket");
            }
            output.flush();
            OfficeUnitLog.info("文件流完成 docId=" + file.docId
                    + " imageFileId=" + file.imageFileId
                    + " versionId=" + file.versionId
                    + " userId=" + user.getUID());
        } catch (OfficePreviewException e) {
            if (!response.isCommitted()) {
                writeError(response, e.getHttpStatus(), e.getCode(), e.getMessage());
            }
        } catch (Throwable t) {
            OfficeUnitLog.error("文件流处理失败 userId=" + user.getUID(), t);
            if (!response.isCommitted()) {
                writeError(response, 500, "OFFICE_PREVIEW_FILE_FETCH_FAILED",
                        "文档加载失败，请稍后重试");
            }
        } finally {
            OfficeUnitUtil.closeQuietly(input);
        }
    }

    private static void setNoStoreHeaders(HttpServletResponse response) {
        response.setHeader("Cache-Control", "private, no-store, no-cache, must-revalidate");
        response.setHeader("Pragma", "no-cache");
        response.setHeader("Expires", "0");
        response.setHeader("X-Content-Type-Options", "nosniff");
        response.setHeader("Referrer-Policy", "no-referrer");
    }

    private static void writeError(HttpServletResponse response, int status,
                                   String code, String message) {
        try {
            response.reset();
            response.setStatus(status);
            response.setContentType("application/json;charset=UTF-8");
            setNoStoreHeaders(response);
            JSONObject error = new JSONObject();
            error.put("code", Integer.valueOf(-1));
            error.put("errorCode", code);
            error.put("msg", message);
            response.getWriter().print(error.toJSONString());
        } catch (Exception ignore) {
            // The response may already be committed after a streaming failure.
        }
    }
}
