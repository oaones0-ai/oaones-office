package com.customization.oaones.officeunit.common;

import java.io.Closeable;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Set;

/**
 * OfficeUnit browser-preview helpers.
 */
public final class OfficeUnitUtil {

    private OfficeUnitUtil() {
    }

    public static String s(Object value) {
        return value == null ? "" : String.valueOf(value);
    }

    public static String trim(Object value) {
        return s(value).trim();
    }

    public static int i(Object value, int defaultValue) {
        try {
            String text = trim(value);
            return text.length() == 0 ? defaultValue : Integer.parseInt(text);
        } catch (Exception e) {
            return defaultValue;
        }
    }

    public static long l(Object value, long defaultValue) {
        try {
            String text = trim(value);
            return text.length() == 0 ? defaultValue : Long.parseLong(text);
        } catch (Exception e) {
            return defaultValue;
        }
    }

    public static boolean b(Object value, boolean defaultValue) {
        String text = trim(value).toLowerCase(Locale.ROOT);
        if (text.length() == 0) {
            return defaultValue;
        }
        return "1".equals(text) || "true".equals(text) || "yes".equals(text) || "on".equals(text);
    }

    public static String lower(Object value) {
        return trim(value).toLowerCase(Locale.ROOT);
    }

    public static String extensionOf(String fileName) {
        String name = safeFileName(fileName);
        int dot = name.lastIndexOf('.');
        if (dot < 0 || dot == name.length() - 1) {
            return "";
        }
        return lower(name.substring(dot + 1));
    }

    public static String safeFileName(String fileName) {
        String name = trim(fileName);
        int slash = Math.max(name.lastIndexOf('/'), name.lastIndexOf('\\'));
        if (slash >= 0) {
            name = name.substring(slash + 1);
        }
        name = name.replace('\r', '_').replace('\n', '_').replace('\t', '_')
                .replace('\u0000', '_');
        return name.length() == 0 ? "document" : name;
    }

    public static Set<Integer> parseIds(String csv) {
        Set<Integer> ids = new LinkedHashSet<Integer>();
        String text = trim(csv);
        if (text.length() == 0) {
            return ids;
        }
        String[] values = text.split("[,，\\s]+");
        for (String value : values) {
            int id = i(value, 0);
            if (id > 0) {
                ids.add(Integer.valueOf(id));
            }
        }
        return ids;
    }

    public static String joinIds(Set<Integer> ids) {
        List<String> values = new ArrayList<String>();
        if (ids != null) {
            for (Integer id : ids) {
                if (id != null && id.intValue() > 0) {
                    values.add(String.valueOf(id));
                }
            }
        }
        return join(values, ",");
    }

    public static String join(List<String> values, String separator) {
        if (values == null || values.isEmpty()) {
            return "";
        }
        StringBuilder result = new StringBuilder();
        for (String value : values) {
            if (value == null) {
                continue;
            }
            if (result.length() > 0) {
                result.append(separator);
            }
            result.append(value);
        }
        return result.toString();
    }

    public static String escapeHtml(Object value) {
        String text = s(value);
        StringBuilder escaped = new StringBuilder(text.length() + 16);
        for (int i = 0; i < text.length(); i++) {
            char c = text.charAt(i);
            switch (c) {
                case '&':
                    escaped.append("&amp;");
                    break;
                case '<':
                    escaped.append("&lt;");
                    break;
                case '>':
                    escaped.append("&gt;");
                    break;
                case '"':
                    escaped.append("&quot;");
                    break;
                case '\'':
                    escaped.append("&#39;");
                    break;
                default:
                    escaped.append(c);
            }
        }
        return escaped.toString();
    }

    public static String contentDisposition(String fileName, String extension) {
        String safe = safeFileName(fileName);
        String fallback = "document";
        String ext = lower(extension);
        if (ext.length() > 0) {
            fallback += "." + ext;
        }
        String encoded;
        try {
            encoded = URLEncoder.encode(safe, "UTF-8").replace("+", "%20");
        } catch (UnsupportedEncodingException e) {
            encoded = fallback;
        }
        return "attachment; filename=\"" + fallback + "\"; filename*=UTF-8''" + encoded;
    }

    public static void closeQuietly(Closeable closeable) {
        if (closeable == null) {
            return;
        }
        try {
            closeable.close();
        } catch (Exception ignore) {
            // Resource cleanup must not hide the main result.
        }
    }
}
