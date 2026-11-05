package com.customization.oaones.officeunit.common;

/**
 * Controlled public preview error.
 */
public class OfficePreviewException extends RuntimeException {

    private final String code;
    private final int httpStatus;

    public OfficePreviewException(String code, String message) {
        this(code, message, 400);
    }

    public OfficePreviewException(String code, String message, int httpStatus) {
        super(message == null || message.length() == 0 ? defaultMessage(code) : message);
        this.code = code == null || code.length() == 0
                ? "OFFICE_PREVIEW_UNSUPPORTED" : code;
        this.httpStatus = httpStatus;
    }

    public String getCode() {
        return code;
    }

    public int getHttpStatus() {
        return httpStatus;
    }

    public static OfficePreviewException forbidden() {
        return new OfficePreviewException("OFFICE_PREVIEW_FORBIDDEN",
                "无权查看该文档", 403);
    }

    public static OfficePreviewException fileNotFound() {
        return new OfficePreviewException("OFFICE_PREVIEW_FILE_NOT_FOUND",
                "未找到文档正文文件", 404);
    }

    public static OfficePreviewException unsupported(String code) {
        return new OfficePreviewException(code, defaultMessage(code), 400);
    }

    public static String defaultMessage(String code) {
        if ("OFFICE_PREVIEW_FORBIDDEN".equals(code)) {
            return "无权查看该文档";
        }
        if ("OFFICE_PREVIEW_FILE_NOT_FOUND".equals(code)) {
            return "未找到文档正文文件";
        }
        if ("OFFICE_PREVIEW_FILE_TOO_LARGE".equals(code)) {
            return "文件过大，暂不支持在线预览";
        }
        if ("OFFICE_PREVIEW_LEGACY_UNSUPPORTED".equals(code)) {
            return "该旧版 Office 格式暂不支持在线预览，请下载查看";
        }
        if ("OFFICE_PREVIEW_MACRO_BLOCKED".equals(code)) {
            return "出于安全原因，宏格式文档已被禁止预览";
        }
        if ("OFFICE_PREVIEW_ENCRYPTED".equals(code)) {
            return "文档已加密，请下载后输入密码打开";
        }
        if ("OFFICE_PREVIEW_TICKET_EXPIRED".equals(code)) {
            return "预览会话已失效，请重新打开";
        }
        return "暂不支持预览此文件格式";
    }
}
