package com.customization.oaones.officeunit.common;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Set;

/**
 * Server-side capability policy. The browser re-detects the actual container.
 */
public final class OfficeFormatPolicy {

    public static final String LEVEL_STANDARD = "standard";
    public static final String LEVEL_BEST_EFFORT = "best_effort";

    private static final Set<String> TEXT = immutable(
            "txt", "log", "text", "cfg", "py", "pyw", "js", "jsx", "ts", "tsx",
            "java", "c", "h", "cpp", "cc", "cxx", "hpp", "cs", "kt", "scala",
            "m", "mm", "go", "rs", "php", "rb", "lua", "r", "sql", "json",
            "xml", "html", "htm", "css", "scss", "less", "yaml", "yml", "md",
            "markdown", "sh", "bash", "zsh", "bat", "cmd", "ps1", "psd1", "psm1",
            "properties", "ini", "in", "toml", "swift", "vue", "diff", "patch",
            "vb", "vbs", "groovy", "gradle", "dart", "coffee", "pl", "pm", "conf");
    private static final Set<String> STANDARD = immutable(
            "doc", "docx", "xlsx", "xls", "csv", "tsv", "ppt", "pptx", "ofd");
    private static final Set<String> BEST_EFFORT = immutable(
            "dotx", "xlsb", "ods", "potx", "ppsx");
    private static final Set<String> MACROS = immutable(
            "docm", "dotm", "xlsm", "xltm", "pptm", "potm", "ppsm");
    private static final Set<String> LEGACY = immutable(
            "wps", "et", "dps");
    private static final Set<String> OTHER_UNSUPPORTED = immutable(
            "rtf", "odt", "ott", "ots", "otp", "xlt", "xltx", "pot", "pps");
    private static final Set<String> NATIVE = immutable(
            "pdf", "ofd", "txt", "log", "sql", "js", "css", "json", "xml",
            "html", "htm", "md", "png", "jpg", "jpeg", "gif", "bmp",
            "webp", "svg", "tif", "tiff", "cad", "dwg", "dxf");

    private OfficeFormatPolicy() {
    }

    public static Format resolve(String extension) {
        String ext = OfficeUnitUtil.lower(extension);
        if (TEXT.contains(ext)) {
            return new Format(ext, "text", LEVEL_STANDARD, true, "");
        }
        if (STANDARD.contains(ext)) {
            return new Format(ext, viewer(ext), LEVEL_STANDARD, true, "");
        }
        if (BEST_EFFORT.contains(ext)) {
            return new Format(ext, viewer(ext), LEVEL_BEST_EFFORT, true, "");
        }
        if (MACROS.contains(ext)) {
            return new Format(ext, "", "", false, "OFFICE_PREVIEW_MACRO_BLOCKED");
        }
        if (LEGACY.contains(ext)) {
            return new Format(ext, "", "", false, "OFFICE_PREVIEW_LEGACY_UNSUPPORTED");
        }
        if (OTHER_UNSUPPORTED.contains(ext)) {
            return new Format(ext, "", "", false, "OFFICE_PREVIEW_UNSUPPORTED");
        }
        return new Format(ext, "", "", false, "OFFICE_PREVIEW_UNSUPPORTED");
    }

    public static List<String> convertibleFormats() {
        List<String> values = new ArrayList<String>();
        values.addAll(TEXT);
        values.addAll(STANDARD);
        values.addAll(BEST_EFFORT);
        return values;
    }

    public static List<String> textFormats() {
        return new ArrayList<String>(TEXT);
    }

    public static boolean isText(String extension) {
        return TEXT.contains(OfficeUnitUtil.lower(extension));
    }

    public static List<String> fallbackFormats(OfficeUnitConfig.Config config) {
        Set<String> values = new LinkedHashSet<String>(config.nativeSet());
        values.addAll(NATIVE);
        values.addAll(LEGACY);
        values.addAll(MACROS);
        values.addAll(OTHER_UNSUPPORTED);
        return new ArrayList<String>(values);
    }

    private static String viewer(String ext) {
        if ("ofd".equals(ext)) {
            return "ofd";
        }
        if ("doc".equals(ext)) {
            return "legacy-doc";
        }
        if ("ppt".equals(ext)) {
            return "legacy-ppt";
        }
        if ("docx".equals(ext) || "dotx".equals(ext)) {
            return "docx";
        }
        if ("pptx".equals(ext) || "potx".equals(ext) || "ppsx".equals(ext)) {
            return "pptx";
        }
        return "spreadsheet";
    }

    private static Set<String> immutable(String... values) {
        return Collections.unmodifiableSet(new LinkedHashSet<String>(Arrays.asList(values)));
    }

    public static final class Format {
        public final String extension;
        public final String viewer;
        public final String supportLevel;
        public final boolean supported;
        public final String errorCode;

        private Format(String extension, String viewer, String supportLevel,
                       boolean supported, String errorCode) {
            this.extension = extension;
            this.viewer = viewer;
            this.supportLevel = supportLevel;
            this.supported = supported;
            this.errorCode = errorCode;
        }

        public boolean standard() {
            return LEVEL_STANDARD.equals(supportLevel);
        }

        public String normalizedExtension() {
            return extension == null ? "" : extension.toLowerCase(Locale.ROOT);
        }
    }
}
