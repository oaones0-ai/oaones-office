package com.customization.oaones.officeunit.filter;

import com.customization.oaones.officeunit.common.OfficeFileAccessSupport;
import com.customization.oaones.officeunit.common.OfficeUnitConfig;
import com.customization.oaones.officeunit.common.OfficeUnitLog;
import com.customization.oaones.officeunit.common.OfficeUnitUtil;
import com.customization.oaones.officeunit.preview.OfficePreviewFileStreamer;
import com.customization.oaones.officeunit.preview.OfficeLegacyPreviewStreamer;
import com.customization.oaones.officeunit.preview.OfficeLegacyPreviewService;
import weaver.hrm.HrmUserVarify;
import weaver.hrm.User;

import javax.servlet.Filter;
import javax.servlet.FilterChain;
import javax.servlet.FilterConfig;
import javax.servlet.ServletException;
import javax.servlet.ServletRequest;
import javax.servlet.ServletResponse;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.PrintWriter;

/**
 * Replaces the E9 default preview region for supported Office files.
 * Every failure or unsupported format falls back to the untouched E9 page.
 */
public class OfficePreviewFilter implements Filter {

    private static final String ORIGINAL_MARKER = "officeunit.original";
    private static final String INTERCEPT_MARKER = "oaones.officeunit.browser-preview";

    @Override
    public void init(FilterConfig filterConfig) {
        // Hot-loaded from WEB-INF/prop/oaones.officeunit.properties.
    }

    @Override
    public void doFilter(ServletRequest request, ServletResponse response, FilterChain chain)
            throws IOException, ServletException {
        if (!(request instanceof HttpServletRequest)
                || !(response instanceof HttpServletResponse)) {
            chain.doFilter(request, response);
            return;
        }
        HttpServletRequest httpRequest = (HttpServletRequest) request;
        HttpServletResponse httpResponse = (HttpServletResponse) response;
        if (OfficeLegacyPreviewStreamer.isRequest(httpRequest)) {
            OfficeLegacyPreviewStreamer.handle(httpRequest, httpResponse);
            return;
        }
        if (OfficePreviewFileStreamer.isRequest(httpRequest)) {
            OfficePreviewFileStreamer.handle(httpRequest, httpResponse);
            return;
        }
        if (!isCandidate(httpRequest)
                || httpRequest.getAttribute(INTERCEPT_MARKER) != null) {
            chain.doFilter(request, response);
            return;
        }

        OfficeUnitConfig.Config config = OfficeUnitConfig.load();
        if (!config.enabled || !config.browserMode()
                || !OfficeUnitConfig.isOaDefaultPreviewEnabled()) {
            chain.doFilter(request, response);
            return;
        }

        User user = HrmUserVarify.getUser(httpRequest, httpResponse);
        if (user == null) {
            chain.doFilter(request, response);
            return;
        }
        OfficeFileAccessSupport.FileInfo file;
        try {
            file = OfficeFileAccessSupport.resolveForPrepare(httpRequest, user, config);
        } catch (Throwable t) {
            // Permission failures and unsupported files retain E9's original handling.
            chain.doFilter(request, response);
            return;
        }

        httpRequest.setAttribute(INTERCEPT_MARKER, Boolean.TRUE);
        try {
            render(httpRequest, httpResponse, file);
        } catch (Throwable t) {
            OfficeUnitLog.error("浏览器预览入口渲染失败 docId=" + file.docId
                    + " imageFileId=" + file.imageFileId, t);
            if (!httpResponse.isCommitted()) {
                httpResponse.reset();
                chain.doFilter(request, response);
            }
        }
    }

    @Override
    public void destroy() {
        OfficeLegacyPreviewService.shutdown();
    }

    private boolean isCandidate(HttpServletRequest request) {
        if ("1".equals(request.getParameter(ORIGINAL_MARKER))) {
            return false;
        }
        String path = request.getRequestURI();
        return path != null && (path.endsWith("/docs/docpreview/main.jsp")
                || path.endsWith("/docs/view/main.jsp"));
    }

    private void render(HttpServletRequest request, HttpServletResponse response,
                        OfficeFileAccessSupport.FileInfo file) throws Exception {
        response.setContentType("text/html;charset=UTF-8");
        response.setHeader("Cache-Control", "private, no-store, no-cache, must-revalidate");
        response.setHeader("Pragma", "no-cache");
        response.setHeader("Expires", "0");
        response.setHeader("X-Content-Type-Options", "nosniff");
        response.setHeader("Referrer-Policy", "no-referrer");
        response.setHeader("X-Frame-Options", "SAMEORIGIN");
        response.setHeader("Permissions-Policy", "camera=(), microphone=(), geolocation=()");
        response.setHeader("Content-Security-Policy",
                "default-src 'none'; "
                        + "script-src 'self'; "
                        + "worker-src 'self' blob:; "
                        + "style-src 'self' 'unsafe-inline'; "
                        + "img-src 'self' blob: data:; "
                        + "font-src 'self'; "
                        + "connect-src 'self'; "
                        + "media-src 'self' blob:; "
                        + "object-src 'none'; frame-src 'none'; "
                        + "base-uri 'none'; form-action 'none'; frame-ancestors 'self'");

        String contextPath = request.getContextPath();
        String originalUrl = originalUrl(request);
        String title = file.title.length() == 0 ? file.fileName : file.title;
        PrintWriter out = response.getWriter();
        out.print("<!doctype html><html lang=\"zh-CN\"><head><meta charset=\"UTF-8\">"
                + "<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">"
                + "<title>" + OfficeUnitUtil.escapeHtml(title) + " - 在线预览</title>"
                + "<link rel=\"stylesheet\" href=\"" + contextPath
                + "/oaones/scripts/css/style.css?v=2026110401\">"
                + "<link rel=\"stylesheet\" href=\"" + contextPath
                + "/oaones/officeunit/css/office-viewer.css?v=2026110403\">"
                + "<script src=\"" + contextPath
                + "/oaones/scripts/js/loading.js?v=2026110401\" "
                + "data-oaones-loading-auto=\"文档加载中ing\" "
                + "data-oaones-loading-key=\"oa-office-document\"></script>"
                + "</head><body class=\"oa-office-preview-page\">"
                + "<a class=\"oa-office-preview-original\" href=\""
                + OfficeUnitUtil.escapeHtml(originalUrl)
                + "\" target=\"_top\" rel=\"noopener noreferrer\">显示原格式</a>"
                + "<main id=\"oaOfficeViewerRoot\" class=\"oa-office-viewer-root\" "
                + "data-oa-office-viewer-auto "
                + "data-api=\"" + OfficeUnitUtil.escapeHtml(contextPath)
                + "/oaones/officeunit/page/api.jsp\" "
                + "data-doc-id=\"" + file.docId + "\" "
                + "data-image-file-id=\"" + file.imageFileId + "\" "
                + "data-file-version=\"" + file.versionId + "\" "
                + "data-file-name=\"" + OfficeUnitUtil.escapeHtml(file.fileName) + "\" "
                + "data-loading-key=\"oa-office-document\" "
                + "data-loading-text=\"文档加载中ing\" "
                + "data-original-url=\"" + OfficeUnitUtil.escapeHtml(originalUrl) + "\"></main>"
                + "<noscript><p class=\"oa-office-preview-noscript\">需要启用 JavaScript"
                + " 才能使用浏览器 Office 预览。</p></noscript>");
        if ("ofd".equals(file.extension)) {
            out.print("<script src=\"" + contextPath
                    + "/oaones/officeunit/js/vendor/ofdrw/ofdrw-viewer-ui.min.js?v=0.1.7\"></script>");
            out.print("<script src=\"" + contextPath
                    + "/oaones/officeunit/js/render/ofd-renderer.js?v=2026110402\"></script>");
        } else if ("text".equals(file.format.viewer)) {
            out.print("<script src=\"" + contextPath
                    + "/oaones/officeunit/js/vendor/codemirror/lib/codemirror.js"
                    + "?v=5.65.21\"></script>");
            out.print("<script src=\"" + contextPath
                    + "/oaones/officeunit/js/render/text-renderer.js?v=2026110501\"></script>");
        } else {
            out.print("<script src=\"" + contextPath
                    + "/oaones/officeunit/js/vendor/jszip.min.js?v=3.10.1\"></script>");
            out.print("<script src=\"" + contextPath
                    + "/oaones/officeunit/js/vendor/docx-preview.min.js?v=0.4.1\"></script>");
            out.print("<script src=\"" + contextPath
                    + "/oaones/officeunit/js/security/sanitizer.js?v=2026110301\"></script>");
            out.print("<script src=\"" + contextPath
                    + "/oaones/officeunit/js/render/docx-renderer.js?v=2026110301\"></script>");
            out.print("<script src=\"" + contextPath
                    + "/oaones/officeunit/js/render/spreadsheet-renderer.js?v=2026110304\"></script>");
            out.print("<script src=\"" + contextPath
                    + "/oaones/officeunit/js/render/pptx-renderer.js?v=2026110301\"></script>");
            out.print("<script src=\"" + contextPath
                    + "/oaones/officeunit/js/render/legacy-office-renderer.js?v=2026110302\"></script>");
        }
        out.print("<script src=\"" + contextPath
                + "/oaones/officeunit/js/office-viewer.js?v=2026110404\"></script>");
        out.print("</body></html>");
        out.flush();
    }

    private String originalUrl(HttpServletRequest request) {
        String query = OfficeUnitUtil.trim(request.getQueryString());
        if (query.length() == 0) {
            return request.getRequestURI() + "?" + ORIGINAL_MARKER + "=1";
        }
        return request.getRequestURI() + "?" + query + "&" + ORIGINAL_MARKER + "=1";
    }
}
