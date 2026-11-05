package com.customization.oaones.officeunit.preview;

import com.alibaba.fastjson.JSONObject;
import com.customization.oaones.officeunit.common.OfficeFileAccessSupport;
import com.customization.oaones.officeunit.common.OfficePreviewException;
import com.customization.oaones.officeunit.common.OfficeUnitConfig;
import com.customization.oaones.officeunit.common.OfficeUnitLog;
import com.customization.oaones.officeunit.common.OfficeUnitUtil;
import weaver.hrm.HrmUserVarify;
import weaver.hrm.User;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;

/**
 * Streams permission-checked legacy Office conversion output.
 */
public final class OfficeLegacyPreviewStreamer {

    private OfficeLegacyPreviewStreamer() {
    }

    public static boolean isRequest(HttpServletRequest request) {
        String path = request == null ? "" : request.getRequestURI();
        return path != null && (path.endsWith("/oaones/officeunit/preview/legacy.jsp")
                || path.endsWith("/oaones/officeunit/preview/legacy"));
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
            writeError(response, 405, "OFFICE_PREVIEW_UNSUPPORTED",
                    "请求方法不支持");
            return;
        }

        try {
            OfficePreviewTicketService tickets = new OfficePreviewTicketService();
            OfficePreviewTicketService.Claims claims = tickets.verify(
                    request.getParameter("token"), user.getUID());
            OfficeUnitConfig.Config config = OfficeUnitConfig.load();
            if (!config.enabled || !config.browserMode()) {
                throw new OfficePreviewException("OFFICE_PREVIEW_UNSUPPORTED",
                        "Office 预览当前未启用", 503);
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
            if (!"doc".equals(file.extension) && !"ppt".equals(file.extension)) {
                throw OfficePreviewException.unsupported(
                        "OFFICE_PREVIEW_UNSUPPORTED");
            }

            OfficeLegacyPreviewService.Artifact artifact =
                    OfficeLegacyPreviewService.load(file, config);
            if ("doc".equals(file.extension)) {
                writeDoc(response, artifact);
            } else {
                String part = OfficeUnitUtil.lower(request.getParameter("part"));
                if ("manifest".equals(part)) {
                    writeManifest(response, artifact);
                } else if ("slide".equals(part)) {
                    writeSlide(response, artifact,
                            OfficeUnitUtil.i(request.getParameter("index"), -1));
                } else {
                    throw new IllegalArgumentException("不支持的预览输出类型");
                }
            }
        } catch (OfficePreviewException e) {
            if (!response.isCommitted()) {
                writeError(response, e.getHttpStatus(), e.getCode(), e.getMessage());
            }
        } catch (IllegalArgumentException e) {
            if (!response.isCommitted()) {
                writeError(response, 400, "OFFICE_PREVIEW_UNSUPPORTED",
                        e.getMessage());
            }
        } catch (Throwable t) {
            OfficeUnitLog.error("旧版 Office 预览输出失败 userId="
                    + user.getUID(), t);
            if (!response.isCommitted()) {
                writeError(response, 500, "OFFICE_PREVIEW_PARSE_FAILED",
                        "旧版 Office 文档转换失败，请稍后重试");
            }
        }
    }

    private static void writeDoc(HttpServletResponse response,
                                 OfficeLegacyPreviewService.Artifact artifact)
            throws Exception {
        byte[] bytes = artifact.html.getBytes(StandardCharsets.UTF_8);
        response.reset();
        response.setStatus(200);
        response.setContentType("text/html;charset=UTF-8");
        response.setContentLength(bytes.length);
        response.setHeader("Content-Disposition", "inline; filename=\"document.html\"");
        response.setHeader("Content-Security-Policy",
                "default-src 'none'; style-src 'unsafe-inline'; img-src data:");
        setNoStoreHeaders(response);
        OutputStream output = response.getOutputStream();
        output.write(bytes);
        output.flush();
    }

    private static void writeManifest(HttpServletResponse response,
                                      OfficeLegacyPreviewService.Artifact artifact)
            throws Exception {
        JSONObject data = new JSONObject();
        data.put("slideCount", Integer.valueOf(artifact.slides.size()));
        data.put("width", Integer.valueOf(artifact.slideWidth));
        data.put("height", Integer.valueOf(artifact.slideHeight));
        JSONObject result = new JSONObject();
        result.put("code", Integer.valueOf(0));
        result.put("msg", "success");
        result.put("data", data);
        byte[] bytes = result.toJSONString().getBytes(StandardCharsets.UTF_8);
        response.reset();
        response.setStatus(200);
        response.setContentType("application/json;charset=UTF-8");
        response.setContentLength(bytes.length);
        setNoStoreHeaders(response);
        OutputStream output = response.getOutputStream();
        output.write(bytes);
        output.flush();
    }

    private static void writeSlide(HttpServletResponse response,
                                   OfficeLegacyPreviewService.Artifact artifact,
                                   int index) throws Exception {
        if (index < 0 || index >= artifact.slides.size()) {
            throw OfficePreviewException.fileNotFound();
        }
        byte[] bytes = artifact.slides.get(index);
        response.reset();
        response.setStatus(200);
        response.setContentType("image/png");
        response.setContentLength(bytes.length);
        response.setHeader("X-Content-Type-Options", "nosniff");
        response.setHeader("Cache-Control",
                "private, max-age=0, no-store, no-cache, must-revalidate");
        OutputStream output = response.getOutputStream();
        output.write(bytes);
        output.flush();
    }

    private static void setNoStoreHeaders(HttpServletResponse response) {
        response.setHeader("Cache-Control",
                "private, no-store, no-cache, must-revalidate");
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
            // Do not hide the original conversion failure.
        }
    }
}
