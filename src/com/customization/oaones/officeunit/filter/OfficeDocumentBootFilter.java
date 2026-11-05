package com.customization.oaones.officeunit.filter;

import com.customization.oaones.officeunit.common.OfficeUnitUtil;

import javax.servlet.Filter;
import javax.servlet.FilterChain;
import javax.servlet.FilterConfig;
import javax.servlet.ServletException;
import javax.servlet.ServletRequest;
import javax.servlet.ServletResponse;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.HttpServletResponseWrapper;
import java.io.IOException;
import java.io.PrintWriter;
import java.io.Writer;

/**
 * Injects the OfficeUnit Loading bootstrap into the document SPA response.
 *
 * This is a custom filter rather than an edit to E9's standard detail JSP.
 * Only the first HTML head/body transition is buffered; the remainder of the
 * response continues to stream to the browser.
 */
public class OfficeDocumentBootFilter implements Filter {

    private static final String INJECTED_ATTRIBUTE =
            "oaones.officeunit.boot.injected";
    private static final int MAX_HEAD_BYTES = 262144;

    @Override
    public void init(FilterConfig filterConfig) {
        // No initialization parameters are required.
    }

    @Override
    public void doFilter(ServletRequest request, ServletResponse response,
                         FilterChain chain) throws IOException, ServletException {
        if (!(request instanceof HttpServletRequest)
                || !(response instanceof HttpServletResponse)) {
            chain.doFilter(request, response);
            return;
        }
        HttpServletRequest httpRequest = (HttpServletRequest) request;
        HttpServletResponse httpResponse = (HttpServletResponse) response;
        if (!shouldInject(httpRequest)
                || httpRequest.getAttribute(INJECTED_ATTRIBUTE) != null) {
            chain.doFilter(request, response);
            return;
        }
        httpRequest.setAttribute(INJECTED_ATTRIBUTE, Boolean.TRUE);
        chain.doFilter(httpRequest, new BootInjectingResponse(
                httpResponse, injection(httpRequest)));
    }

    @Override
    public void destroy() {
        // No resources.
    }

    private boolean shouldInject(HttpServletRequest request) {
        if (!"GET".equalsIgnoreCase(request.getMethod())) {
            return false;
        }
        String path = request.getRequestURI();
        if (path == null || !path.endsWith("/spa/document/index.jsp")) {
            return false;
        }
        // The loading indicator must be emitted before any database or preview
        // capability lookup, and E9's request wrappers do not expose every
        // camel-case document query parameter consistently. The detail-page
        // path is the stable first-paint boundary. The browser script removes
        // the indicator again if the frame is not an OfficeUnit preview.
        return true;
    }

    private String injection(HttpServletRequest request) {
        String contextPath = OfficeUnitUtil.trim(request.getContextPath());
        String scriptUrl = OfficeUnitUtil.escapeHtml(contextPath
                + "/oaones/officeunit/ecode/boot-loading.js?v=2026110408");
        return "<style data-oaones-officeunit-boot>"
                + ".oa-office-ecode-loading{position:fixed;left:0;top:0;width:100%;"
                + "height:100%;z-index:2147483000;display:flex;align-items:center;"
                + "justify-content:center;box-sizing:border-box;padding:24px;"
                + "background:rgba(255,255,255,.9);font-family:-apple-system,"
                + "BlinkMacSystemFont,\"Segoe UI\",\"Microsoft YaHei\",sans-serif}"
                + ".oa-office-ecode-loading-panel{display:flex;align-items:center;"
                + "gap:11px;padding:15px 20px;border:1px solid rgba(0,85,204,.12);"
                + "border-radius:10px;background:#fff;box-shadow:0 8px 28px "
                + "rgba(24,39,75,.14);color:#24324a;font-size:14px;line-height:20px}"
                + ".oa-office-ecode-loading-spinner{width:20px;height:20px;"
                + "border:2px solid #dce8fb;border-top-color:#1a73e8;border-radius:50%;"
                + "animation:oa-office-ecode-spin .75s linear infinite}"
                + "@keyframes oa-office-ecode-spin{to{transform:rotate(360deg)}}"
                + "@media (prefers-reduced-motion:reduce){"
                + ".oa-office-ecode-loading-spinner{animation:none}}"
                + "</style>"
                + "<script src=\"" + scriptUrl + "\"></script>";
    }

    private static final class BootInjectingResponse
            extends HttpServletResponseWrapper {
        private final HttpServletResponse response;
        private final String injection;
        private PrintWriter writer;

        private BootInjectingResponse(HttpServletResponse response,
                                      String injection) {
            super(response);
            this.response = response;
            this.injection = injection;
        }

        @Override
        public PrintWriter getWriter() throws IOException {
            if (writer == null) {
                writer = new PrintWriter(new BootInjectingWriter(
                        response.getWriter(), injection), true);
            }
            return writer;
        }

        @Override
        public void setContentLength(int length) {
            // Injection changes the response length.
        }

        @Override
        public void setHeader(String name, String value) {
            if (!"Content-Length".equalsIgnoreCase(name)) {
                super.setHeader(name, value);
            }
        }

        @Override
        public void addHeader(String name, String value) {
            if (!"Content-Length".equalsIgnoreCase(name)) {
                super.addHeader(name, value);
            }
        }
    }

    private static final class BootInjectingWriter extends Writer {
        private final Writer target;
        private final String injection;
        private final StringBuilder pending = new StringBuilder(8192);
        private boolean injected;

        private BootInjectingWriter(Writer target, String injection) {
            this.target = target;
            this.injection = injection;
        }

        @Override
        public void write(char[] buffer, int offset, int length) throws IOException {
            if (injected) {
                target.write(buffer, offset, length);
                return;
            }
            pending.append(buffer, offset, length);
            injectIfReady(false);
        }

        @Override
        public void flush() throws IOException {
            if (injected) {
                target.flush();
            }
        }

        @Override
        public void close() throws IOException {
            injectIfReady(true);
            target.flush();
            target.close();
        }

        private void injectIfReady(boolean closing) throws IOException {
            if (injected) {
                return;
            }
            int bodyStart = findBodyStart();
            if (bodyStart >= 0) {
                int tagEnd = pending.indexOf(">", bodyStart + 5);
                if (tagEnd >= 0) {
                    target.write(pending.substring(0, tagEnd + 1));
                    target.write(injection);
                    if (tagEnd + 1 < pending.length()) {
                        target.write(pending.substring(tagEnd + 1));
                    }
                    injected = true;
                    pending.setLength(0);
                    target.flush();
                    return;
                }
            }
            if (pending.length() > MAX_HEAD_BYTES) {
                target.write(pending.toString());
                pending.setLength(0);
                injected = true;
            } else if (closing) {
                target.write(pending.toString());
                pending.setLength(0);
                injected = true;
            }
        }

        private int findBodyStart() {
            for (int i = 0; i <= pending.length() - 5; i++) {
                if (pending.charAt(i) == '<'
                        && pending.charAt(i + 1) == 'b'
                        && pending.charAt(i + 2) == 'o'
                        && pending.charAt(i + 3) == 'd'
                        && pending.charAt(i + 4) == 'y'
                        && (i + 5 == pending.length()
                            || Character.isWhitespace(pending.charAt(i + 5))
                            || pending.charAt(i + 5) == '>')) {
                    return i;
                }
            }
            return -1;
        }
    }
}
