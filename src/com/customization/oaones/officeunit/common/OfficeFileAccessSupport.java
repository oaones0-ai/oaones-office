package com.customization.oaones.officeunit.common;

import com.api.doc.detail.service.DocViewPermission;
import weaver.conn.RecordSet;
import weaver.file.ImageFileManager;
import weaver.hrm.User;

import javax.servlet.http.HttpServletRequest;
import java.io.InputStream;
import java.util.Locale;
import java.util.Map;

/**
 * Resolves an OA document file and applies ownership, version, permission and
 * format policy checks. No filesystem path is exposed to callers.
 */
public final class OfficeFileAccessSupport {

    private OfficeFileAccessSupport() {
    }

    public static FileInfo resolveForPrepare(HttpServletRequest request, User user,
                                             OfficeUnitConfig.Config config) {
        int docId = OfficeUnitUtil.i(first(request, "docId", "docid"), 0);
        int imageFileId = OfficeUnitUtil.i(request.getParameter("imageFileId"), 0);
        int requestedVersion = OfficeUnitUtil.i(first(request, "fileVersion", "versionId"), 0);
        return resolve(docId, imageFileId, requestedVersion, user, request, config);
    }

    public static FileInfo resolveBound(int docId, int imageFileId, int versionId,
                                        User user, HttpServletRequest request,
                                        OfficeUnitConfig.Config config) {
        return resolve(docId, imageFileId, versionId, user, request, config);
    }

    private static FileInfo resolve(int docId, int imageFileId, int requestedVersion,
                                    User user, HttpServletRequest request,
                                    OfficeUnitConfig.Config config) {
        if (user == null || user.getUID() <= 0) {
            throw OfficePreviewException.forbidden();
        }
        if (docId <= 0 || imageFileId <= 0) {
            throw OfficePreviewException.fileNotFound();
        }
        if (!hasReadPermission(docId, user, request)) {
            throw OfficePreviewException.forbidden();
        }
        if (!KnowledgeDirectoryScope.allows(docId, config)) {
            throw OfficePreviewException.forbidden();
        }

        RecordSet rs = new RecordSet();
        rs.executeQuery("select docid, imagefileid, imagefilename, versionid, docfiletype"
                        + " from DocImageFile where docid = ? and imagefileid = ?",
                Integer.valueOf(docId), Integer.valueOf(imageFileId));
        if (!rs.next()) {
            throw OfficePreviewException.fileNotFound();
        }

        FileInfo info = new FileInfo();
        info.docId = rs.getInt("docid");
        info.imageFileId = rs.getInt("imagefileid");
        info.versionId = rs.getInt("versionid");
        info.fileName = OfficeUnitUtil.safeFileName(rs.getString("imagefilename"));
        info.docFileType = OfficeUnitUtil.trim(rs.getString("docfiletype"));
        info.title = documentTitle(docId);

        if (info.docId != docId || info.imageFileId != imageFileId) {
            throw OfficePreviewException.fileNotFound();
        }
        if (requestedVersion > 0 && requestedVersion != info.versionId) {
            throw new OfficePreviewException("OFFICE_PREVIEW_FILE_NOT_FOUND",
                    "文档版本已变化，请重新打开预览", 409);
        }

        info.extension = OfficeUnitUtil.extensionOf(info.fileName);
        OfficeFormatPolicy.Format format = OfficeFormatPolicy.resolve(info.extension);
        if (!format.supported) {
            throw OfficePreviewException.unsupported(format.errorCode);
        }
        if (!config.isFormatEnabled(format.extension)) {
            throw OfficePreviewException.unsupported("OFFICE_PREVIEW_FORMAT_DISABLED");
        }
        if (!config.allowMacroFormats
                && ("docm".equals(format.extension) || "pptm".equals(format.extension)
                || "xlsm".equals(format.extension))) {
            throw OfficePreviewException.unsupported("OFFICE_PREVIEW_MACRO_BLOCKED");
        }
        info.format = format;

        ImageFileManager manager = new ImageFileManager();
        try {
            manager.getImageFileInfoById(imageFileId);
            if ("1".equals(OfficeUnitUtil.trim(manager.getIsencrypt()))) {
                throw OfficePreviewException.unsupported("OFFICE_PREVIEW_ENCRYPTED");
            }
            info.fileSize = parseSize(manager.getFilesize());
        } catch (OfficePreviewException e) {
            throw e;
        } catch (Throwable t) {
            OfficeUnitLog.error("读取文件元数据失败 imageFileId=" + imageFileId
                    + " userId=" + user.getUID(), t);
        }
        if (info.fileSize <= 0L) {
            info.fileSize = countBytes(imageFileId, config.maxInputBytes);
        }
        if (info.fileSize > config.maxInputBytes) {
            throw OfficePreviewException.unsupported("OFFICE_PREVIEW_FILE_TOO_LARGE");
        }

        byte[] header = readHeader(imageFileId, 32);
        validateContainer(info.extension, header);
        if (!info.format.supported) {
            throw OfficePreviewException.unsupported(info.format.errorCode);
        }
        return info;
    }

    private static boolean hasReadPermission(int docId, User user, HttpServletRequest request) {
        try {
            DocViewPermission permission = new DocViewPermission();
            Map<String, Boolean> levels = permission.getShareLevel(docId, user, false);
            boolean readable = levels != null
                    && Boolean.TRUE.equals(levels.get(DocViewPermission.READ));
            if (!readable && request != null) {
                readable = permission.hasRightFromOtherMould(docId, user, request);
            }
            return readable;
        } catch (Throwable t) {
            OfficeUnitLog.error("文档阅读权限校验失败 docId=" + docId
                    + " userId=" + (user == null ? 0 : user.getUID()), t);
            return false;
        }
    }

    private static String documentTitle(int docId) {
        RecordSet rs = new RecordSet();
        rs.executeQuery("select docsubject from DocDetail where id = ?", Integer.valueOf(docId));
        return rs.next() ? OfficeUnitUtil.trim(rs.getString("docsubject")) : "";
    }

    private static byte[] readHeader(int imageFileId, int limit) {
        InputStream input = null;
        try {
            input = ImageFileManager.getInputStreamById(imageFileId);
            if (input == null) {
                throw OfficePreviewException.fileNotFound();
            }
            byte[] buffer = new byte[limit];
            int total = 0;
            while (total < limit) {
                int count = input.read(buffer, total, limit - total);
                if (count < 0) {
                    break;
                }
                total += count;
            }
            if (total == buffer.length) {
                return buffer;
            }
            byte[] result = new byte[total];
            System.arraycopy(buffer, 0, result, 0, total);
            return result;
        } catch (OfficePreviewException e) {
            throw e;
        } catch (Exception e) {
            OfficeUnitLog.error("读取文件头失败 imageFileId=" + imageFileId, e);
            throw OfficePreviewException.fileNotFound();
        } finally {
            OfficeUnitUtil.closeQuietly(input);
        }
    }

    private static long countBytes(int imageFileId, long maxBytes) {
        InputStream input = null;
        try {
            input = ImageFileManager.getInputStreamById(imageFileId);
            if (input == null) {
                throw OfficePreviewException.fileNotFound();
            }
            byte[] buffer = new byte[8192];
            long total = 0L;
            int count;
            while ((count = input.read(buffer)) >= 0) {
                total += count;
                if (total > maxBytes) {
                    throw OfficePreviewException.unsupported(
                            "OFFICE_PREVIEW_FILE_TOO_LARGE");
                }
            }
            return total;
        } catch (OfficePreviewException e) {
            throw e;
        } catch (Exception e) {
            OfficeUnitLog.error("计算文件大小失败 imageFileId=" + imageFileId, e);
            throw OfficePreviewException.fileNotFound();
        } finally {
            OfficeUnitUtil.closeQuietly(input);
        }
    }

    private static void validateContainer(String extension, byte[] header) {
        String ext = OfficeUnitUtil.lower(extension);
        boolean zip = isZip(header);
        boolean ole = isOle(header);
        if (OfficeFormatPolicy.isText(ext)) {
            if (!isTextHeader(header)) {
                throw OfficePreviewException.unsupported("OFFICE_PREVIEW_PARSE_FAILED");
            }
        } else if ("docx".equals(ext) || "dotx".equals(ext)
                || "xlsx".equals(ext) || "pptx".equals(ext)
                || "potx".equals(ext) || "ppsx".equals(ext) || "ods".equals(ext)) {
            if (ole) {
                throw OfficePreviewException.unsupported("OFFICE_PREVIEW_ENCRYPTED");
            }
            if (!zip) {
                throw OfficePreviewException.unsupported("OFFICE_PREVIEW_ARCHIVE_INVALID");
            }
        } else if ("xls".equals(ext) && !ole && !zip) {
            // Some producers write BIFF directly; SheetJS applies the final check.
            if (header.length < 4 || (header[0] & 0xff) != 0x09
                    || ((header[1] & 0xff) != 0x00 && (header[1] & 0xff) != 0x02
                    && (header[1] & 0xff) != 0x04 && (header[1] & 0xff) != 0x08)) {
                throw OfficePreviewException.unsupported("OFFICE_PREVIEW_ARCHIVE_INVALID");
            }
        } else if ("csv".equals(ext) || "tsv".equals(ext)) {
            for (int i = 0; i < Math.min(header.length, 512); i++) {
                if (header[i] == 0) {
                    throw OfficePreviewException.unsupported("OFFICE_PREVIEW_PARSE_FAILED");
                }
            }
        } else if ("ofd".equals(ext)) {
            if (!zip) {
                throw OfficePreviewException.unsupported(
                        "OFFICE_PREVIEW_ARCHIVE_INVALID");
            }
        } else if ("doc".equals(ext) || "ppt".equals(ext)) {
            if (!ole) {
                throw OfficePreviewException.unsupported(
                        "OFFICE_PREVIEW_ARCHIVE_INVALID");
            }
        }
    }

    private static boolean isTextHeader(byte[] header) {
        if (header == null || header.length == 0) {
            return true;
        }
        if (header.length >= 2
                && (((header[0] & 0xff) == 0xff && (header[1] & 0xff) == 0xfe)
                || ((header[0] & 0xff) == 0xfe && (header[1] & 0xff) == 0xff))) {
            return true;
        }
        if (header.length >= 3
                && (header[0] & 0xff) == 0xef
                && (header[1] & 0xff) == 0xbb
                && (header[2] & 0xff) == 0xbf) {
            return true;
        }
        int oddNulls = 0;
        int evenNulls = 0;
        for (int i = 0; i < header.length; i++) {
            int value = header[i] & 0xff;
            if (value == 0) {
                if (i % 2 == 0) {
                    evenNulls++;
                } else {
                    oddNulls++;
                }
                continue;
            }
            if (value < 0x20 && value != 0x09 && value != 0x0a
                    && value != 0x0c && value != 0x0d && value != 0x1b) {
                return false;
            }
        }
        int pairs = Math.max(1, header.length / 2);
        boolean utf16Le = oddNulls > pairs / 4 && evenNulls < pairs / 20;
        boolean utf16Be = evenNulls > pairs / 4 && oddNulls < pairs / 20;
        return (oddNulls == 0 && evenNulls == 0) || utf16Le || utf16Be;
    }

    private static boolean isZip(byte[] header) {
        return header.length >= 4 && header[0] == 0x50 && header[1] == 0x4b
                && (header[2] == 0x03 || header[2] == 0x05 || header[2] == 0x07);
    }

    private static boolean isOle(byte[] header) {
        int[] magic = {0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1};
        if (header.length < magic.length) {
            return false;
        }
        for (int i = 0; i < magic.length; i++) {
            if ((header[i] & 0xff) != magic[i]) {
                return false;
            }
        }
        return true;
    }

    private static long parseSize(String value) {
        String text = OfficeUnitUtil.trim(value).replace(",", "").toUpperCase(Locale.ROOT);
        if (text.length() == 0) {
            return 0L;
        }
        int end = 0;
        while (end < text.length()
                && ((text.charAt(end) >= '0' && text.charAt(end) <= '9')
                || text.charAt(end) == '.')) {
            end++;
        }
        if (end == 0) {
            return 0L;
        }
        double number;
        try {
            number = Double.parseDouble(text.substring(0, end));
        } catch (Exception e) {
            return 0L;
        }
        String unit = text.substring(end).trim();
        if (unit.startsWith("G")) {
            number *= 1024D * 1024D * 1024D;
        } else if (unit.startsWith("M")) {
            number *= 1024D * 1024D;
        } else if (unit.startsWith("K")) {
            number *= 1024D;
        }
        return Math.max(0L, Math.round(number));
    }

    private static String first(HttpServletRequest request, String first, String second) {
        String value = request == null ? null : request.getParameter(first);
        if (value == null || value.length() == 0) {
            value = request == null ? null : request.getParameter(second);
        }
        return value;
    }

    public static final class FileInfo {
        public int docId;
        public int imageFileId;
        public int versionId;
        public long fileSize;
        public String fileName;
        public String docFileType;
        public String title;
        public String extension;
        public OfficeFormatPolicy.Format format;
    }
}
