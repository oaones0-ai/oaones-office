package com.customization.oaones.officeunit.preview;

import com.alibaba.fastjson.JSONObject;
import com.customization.oaones.officeunit.common.OfficeFileAccessSupport;
import com.customization.oaones.officeunit.common.OfficeFormatPolicy;
import com.customization.oaones.officeunit.common.OfficePreviewException;
import com.customization.oaones.officeunit.common.OfficeUnitConfig;
import com.customization.oaones.officeunit.common.OfficeUnitLog;
import com.customization.oaones.officeunit.common.OfficeUnitUtil;
import weaver.hrm.User;

import javax.servlet.http.HttpServletRequest;

/**
 * Issues a short-lived ticket for a permission-checked browser preview.
 */
public final class OfficePreviewController {

    private final OfficePreviewTicketService tickets = new OfficePreviewTicketService();

    public JSONObject prepare(HttpServletRequest request, User user) {
        OfficeUnitConfig.Config config = OfficeUnitConfig.load();
        if (!config.enabled || !config.browserMode()) {
            throw new OfficePreviewException("OFFICE_PREVIEW_UNSUPPORTED",
                    "Office 浏览器预览当前未启用", 503);
        }
        OfficeFileAccessSupport.FileInfo file = OfficeFileAccessSupport
                .resolveForPrepare(request, user, config);
        OfficePreviewTicketService.Ticket ticket = tickets.issue(file, user, config);

        OfficeUnitLog.info("准备预览 docId=" + file.docId + " imageFileId=" + file.imageFileId
                + " versionId=" + file.versionId + " userId=" + user.getUID()
                + " ext=" + file.extension);

        JSONObject data = new JSONObject();
        data.put("fileToken", ticket.token);
        data.put("fileName", file.fileName);
        data.put("extension", file.extension);
        data.put("fileSize", Long.valueOf(file.fileSize));
        data.put("supportLevel", file.format.supportLevel);
        data.put("viewer", file.format.viewer);
        data.put("expiresAt", Long.valueOf(ticket.expiresAt));
        data.put("limits", limits(config));
        data.put("fileUrl", "/oaones/officeunit/preview/file.jsp");
        if ("legacy-doc".equals(file.format.viewer)
                || "legacy-ppt".equals(file.format.viewer)) {
            data.put("legacyUrl", "/oaones/officeunit/preview/legacy.jsp");
        }
        return ok(data);
    }

    private JSONObject limits(OfficeUnitConfig.Config config) {
        JSONObject value = new JSONObject();
        value.put("maxInputBytes", Long.valueOf(config.maxInputBytes));
        value.put("maxUncompressedBytes", Long.valueOf(config.maxUncompressedBytes));
        value.put("maxArchiveEntries", Integer.valueOf(config.maxArchiveEntries));
        value.put("maxCompressionRatio", Integer.valueOf(config.maxCompressionRatio));
        value.put("maxXmlBytes", Long.valueOf(config.maxXmlBytes));
        value.put("maxPageCount", Integer.valueOf(config.maxPageCount));
        value.put("maxSheetCount", Integer.valueOf(config.maxSheetCount));
        value.put("maxCellCount", Integer.valueOf(config.maxCellCount));
        value.put("maxSlideCount", Integer.valueOf(config.maxSlideCount));
        value.put("maxShapeCount", Integer.valueOf(config.maxShapeCount));
        value.put("maxAssetBytes", Long.valueOf(config.maxAssetBytes));
        value.put("parseTimeoutSeconds", Integer.valueOf(config.parseTimeoutSeconds));
        value.put("allowExternalLinks", Boolean.valueOf(config.allowExternalLinks));
        value.put("allowRemoteFonts", Boolean.valueOf(config.allowRemoteFonts));
        return value;
    }

    private JSONObject ok(Object data) {
        JSONObject result = new JSONObject();
        result.put("code", Integer.valueOf(0));
        result.put("msg", "success");
        result.put("data", data);
        return result;
    }
}
