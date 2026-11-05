package com.customization.oaones.officeunit.common;

import weaver.general.GCONST;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Properties;
import java.util.Set;

/**
 * Hot-loaded browser-preview configuration.
 */
public final class OfficeUnitConfig {

    public static final String PROP_FILE = "oaones.officeunit.properties";
    public static final String MODE_BROWSER_JS = "browser-js";

    private static final String DEFAULT_NATIVE =
            "pdf,ofd,txt,log,sql,js,css,json,xml,html,htm,md,png,jpg,jpeg,gif,bmp,"
                    + "webp,svg,tif,tiff,cad,dwg,dxf";
    private static final String DEFAULT_FORMATS =
            "txt,log,text,cfg,py,pyw,js,jsx,ts,tsx,java,c,h,cpp,cc,cxx,hpp,"
                    + "cs,kt,scala,m,mm,go,rs,php,rb,lua,r,sql,json,xml,html,htm,"
                    + "css,scss,less,yaml,yml,md,markdown,sh,bash,zsh,bat,cmd,"
                    + "ps1,psd1,psm1,properties,ini,in,toml,swift,vue,diff,patch,"
                    + "vb,vbs,groovy,gradle,dart,coffee,pl,pm,conf,"
                    + "doc,docx,xlsx,xls,csv,tsv,ppt,pptx,dotx,xlsb,ods,potx,ppsx,ofd";

    private static volatile Snapshot cache;

    private OfficeUnitConfig() {
    }

    public static Config load() {
        File file = configFile();
        long stamp = file.exists() ? file.lastModified() : 0L;
        long length = file.exists() ? file.length() : 0L;
        Snapshot current = cache;
        if (current != null && current.stamp == stamp && current.length == length) {
            return current.value;
        }
        synchronized (OfficeUnitConfig.class) {
            current = cache;
            stamp = file.exists() ? file.lastModified() : 0L;
            length = file.exists() ? file.length() : 0L;
            if (current != null && current.stamp == stamp && current.length == length) {
                return current.value;
            }
            Config value = read(file);
            cache = new Snapshot(value, stamp, length);
            return value;
        }
    }

    public static Config save(boolean enabled, String includeDirs, String excludeDirs,
                              String enabledFormats) {
        synchronized (OfficeUnitConfig.class) {
            Config old = load();
            Config updated = old.withAdminSettings(enabled, includeDirs, excludeDirs,
                    enabledFormats);
            File target = configFile();
            File parent = target.getParentFile();
            if (parent != null && !parent.exists() && !parent.mkdirs()) {
                throw new IllegalStateException("无法创建配置目录");
            }
            File temp = new File(parent, PROP_FILE + ".saving-" + System.nanoTime());
            try {
                write(temp, updated);
                try {
                    Files.move(temp.toPath(), target.toPath(),
                            StandardCopyOption.REPLACE_EXISTING, StandardCopyOption.ATOMIC_MOVE);
                } catch (Exception atomicFailure) {
                    Files.move(temp.toPath(), target.toPath(), StandardCopyOption.REPLACE_EXISTING);
                }
                cache = null;
                OfficeUnitLog.info("配置已更新 enabled=" + updated.enabled
                        + " includeDirs=" + updated.includeDirs
                        + " excludeDirs=" + updated.excludeDirs
                        + " formats=" + updated.formats);
                return load();
            } catch (Exception e) {
                throw new IllegalStateException("保存预览配置失败", e);
            } finally {
                if (temp.exists() && !temp.delete()) {
                    OfficeUnitLog.error("未能删除配置临时文件", null);
                }
            }
        }
    }

    public static boolean isOaDefaultPreviewEnabled() {
        try {
            weaver.general.BaseBean bean = new weaver.general.BaseBean();
            String usePreview = bean.getPropValue("docpreview", "IsUseDocPreview");
            String previewType = bean.getPropValue("docpreview", "DocPreviewType");
            return "1".equals(OfficeUnitUtil.trim(usePreview))
                    && "2".equals(OfficeUnitUtil.trim(previewType));
        } catch (Throwable t) {
            return false;
        }
    }

    public static File configFile() {
        return new File(new File(OfficeUnitUtil.trim(GCONST.getRootPath()), "WEB-INF/prop"),
                PROP_FILE);
    }

    private static Config read(File file) {
        Properties properties = new Properties();
        if (file.exists() && file.isFile()) {
            FileInputStream input = null;
            try {
                input = new FileInputStream(file);
                properties.load(input);
            } catch (Exception e) {
                OfficeUnitLog.error("读取 OfficeUnit 配置失败，使用安全默认值", e);
            } finally {
                OfficeUnitUtil.closeQuietly(input);
            }
        }
        return new Config(
                OfficeUnitUtil.b(properties.getProperty("officeunit.preview.enabled"), true),
                normalizeIds(properties.getProperty("officeunit.preview.include_dirs")),
                normalizeIds(properties.getProperty("officeunit.preview.exclude_dirs")),
                normalizeIds(properties.getProperty("officeunit.preview.admin_ids")),
                defaultIfEmpty(properties.getProperty("officeunit.preview.mode"), MODE_BROWSER_JS),
                boundedLong(properties.getProperty("officeunit.preview.max_input_bytes"),
                        52428800L, 1048576L, 1073741824L),
                boundedLong(properties.getProperty("officeunit.preview.max_uncompressed_bytes"),
                        209715200L, 1048576L, 2147483648L),
                boundedInt(properties.getProperty("officeunit.preview.max_archive_entries"),
                        10000, 1, 100000),
                boundedInt(properties.getProperty("officeunit.preview.max_compression_ratio"),
                        100, 1, 1000),
                boundedLong(properties.getProperty("officeunit.preview.max_xml_bytes"),
                        20971520L, 65536L, 268435456L),
                boundedInt(properties.getProperty("officeunit.preview.max_page_count"),
                        2000, 1, 50000),
                boundedInt(properties.getProperty("officeunit.preview.max_sheet_count"),
                        50, 1, 500),
                boundedInt(properties.getProperty("officeunit.preview.max_cell_count"),
                        1000000, 1, 10000000),
                boundedInt(properties.getProperty("officeunit.preview.max_slide_count"),
                        500, 1, 5000),
                boundedInt(properties.getProperty("officeunit.preview.max_shape_count"),
                        50000, 1, 500000),
                boundedLong(properties.getProperty("officeunit.preview.max_asset_bytes"),
                        52428800L, 1048576L, 536870912L),
                boundedInt(properties.getProperty("officeunit.preview.parse_timeout_seconds"),
                        30, 1, 300),
                boundedInt(properties.getProperty("officeunit.preview.ticket_ttl_seconds"),
                        600, 60, 3600),
                OfficeUnitUtil.b(properties.getProperty("officeunit.preview.allow_external_links"),
                        false),
                OfficeUnitUtil.b(properties.getProperty("officeunit.preview.allow_remote_fonts"),
                        false),
                OfficeUnitUtil.b(properties.getProperty("officeunit.preview.allow_macro_formats"),
                        false),
                defaultIfEmpty(properties.getProperty("officeunit.preview.legacy_action"), "download"),
                normalizeFormats(properties.getProperty("officeunit.preview.native_formats"),
                        DEFAULT_NATIVE),
                optionalFormats(properties.getProperty("officeunit.preview.formats"),
                        DEFAULT_FORMATS));
    }

    private static void write(File file, Config config) throws Exception {
        String content = "# OAones OfficeUnit browser preview. Keep this file ASCII-only.\n"
                + "# Empty include_dirs means all knowledge-document directories.\n"
                + "# Includes match descendants; exclude_dirs wins.\n"
                + "officeunit.preview.enabled=" + (config.enabled ? "true" : "false") + "\n"
                + "officeunit.preview.include_dirs=" + config.includeDirs + "\n"
                + "officeunit.preview.exclude_dirs=" + config.excludeDirs + "\n"
                + "officeunit.preview.admin_ids=" + config.adminIds + "\n"
                + "officeunit.preview.mode=" + ascii(config.mode) + "\n"
                + "officeunit.preview.max_input_bytes=" + config.maxInputBytes + "\n"
                + "officeunit.preview.max_uncompressed_bytes="
                + config.maxUncompressedBytes + "\n"
                + "officeunit.preview.max_archive_entries=" + config.maxArchiveEntries + "\n"
                + "officeunit.preview.max_compression_ratio=" + config.maxCompressionRatio + "\n"
                + "officeunit.preview.max_xml_bytes=" + config.maxXmlBytes + "\n"
                + "officeunit.preview.max_page_count=" + config.maxPageCount + "\n"
                + "officeunit.preview.max_sheet_count=" + config.maxSheetCount + "\n"
                + "officeunit.preview.max_cell_count=" + config.maxCellCount + "\n"
                + "officeunit.preview.max_slide_count=" + config.maxSlideCount + "\n"
                + "officeunit.preview.max_shape_count=" + config.maxShapeCount + "\n"
                + "officeunit.preview.max_asset_bytes=" + config.maxAssetBytes + "\n"
                + "officeunit.preview.parse_timeout_seconds=" + config.parseTimeoutSeconds + "\n"
                + "officeunit.preview.ticket_ttl_seconds=" + config.ticketTtlSeconds + "\n"
                + "officeunit.preview.allow_external_links="
                + (config.allowExternalLinks ? "true" : "false") + "\n"
                + "officeunit.preview.allow_remote_fonts="
                + (config.allowRemoteFonts ? "true" : "false") + "\n"
                + "officeunit.preview.allow_macro_formats="
                + (config.allowMacroFormats ? "true" : "false") + "\n"
                + "officeunit.preview.legacy_action=" + ascii(config.legacyAction) + "\n"
                + "officeunit.preview.native_formats=" + ascii(config.nativeFormats) + "\n"
                + "officeunit.preview.formats=" + ascii(config.formats) + "\n";
        OutputStream output = null;
        try {
            output = new FileOutputStream(file);
            output.write(content.getBytes(StandardCharsets.US_ASCII));
            output.flush();
        } finally {
            OfficeUnitUtil.closeQuietly(output);
        }
    }

    private static int boundedInt(String value, int defaultValue, int min, int max) {
        int parsed = OfficeUnitUtil.i(value, defaultValue);
        if (parsed < min || parsed > max) {
            return defaultValue;
        }
        return parsed;
    }

    private static long boundedLong(String value, long defaultValue, long min, long max) {
        long parsed = OfficeUnitUtil.l(value, defaultValue);
        if (parsed < min || parsed > max) {
            return defaultValue;
        }
        return parsed;
    }

    private static String normalizeIds(String value) {
        return OfficeUnitUtil.joinIds(new LinkedHashSet<Integer>(
                OfficeUnitUtil.parseIds(OfficeUnitUtil.s(value))));
    }

    private static String normalizeFormats(String value, String defaultValue) {
        String text = OfficeUnitUtil.trim(value);
        if (text.length() == 0) {
            text = defaultValue;
        }
        return normalizeFormatList(text);
    }

    private static String optionalFormats(String value, String defaultValue) {
        return value == null ? normalizeFormatList(defaultValue)
                : normalizeFormatList(value);
    }

    private static String normalizeFormatList(String value) {
        String text = OfficeUnitUtil.trim(value);
        Set<String> values = new LinkedHashSet<String>();
        String[] parts = text.split("[,，\\s]+");
        for (String part : parts) {
            String ext = OfficeUnitUtil.lower(part);
            if (ext.startsWith(".")) {
                ext = ext.substring(1);
            }
            if (ext.matches("[a-z0-9]{1,16}")) {
                values.add(ext);
            }
        }
        return OfficeUnitUtil.join(new ArrayList<String>(values), ",");
    }

    private static String defaultIfEmpty(String value, String defaultValue) {
        String text = OfficeUnitUtil.trim(value);
        return text.length() == 0 ? defaultValue : text;
    }

    private static String ascii(String value) {
        String text = OfficeUnitUtil.s(value);
        StringBuilder result = new StringBuilder(text.length());
        for (int i = 0; i < text.length(); i++) {
            char c = text.charAt(i);
            result.append(c >= 32 && c <= 126 ? c : '_');
        }
        return result.toString();
    }

    public static final class Config {
        public final boolean enabled;
        public final String includeDirs;
        public final String excludeDirs;
        public final String adminIds;
        public final String mode;
        public final long maxInputBytes;
        public final long maxUncompressedBytes;
        public final int maxArchiveEntries;
        public final int maxCompressionRatio;
        public final long maxXmlBytes;
        public final int maxPageCount;
        public final int maxSheetCount;
        public final int maxCellCount;
        public final int maxSlideCount;
        public final int maxShapeCount;
        public final long maxAssetBytes;
        public final int parseTimeoutSeconds;
        public final int ticketTtlSeconds;
        public final boolean allowExternalLinks;
        public final boolean allowRemoteFonts;
        public final boolean allowMacroFormats;
        public final String legacyAction;
        public final String nativeFormats;
        public final String formats;

        private Config(boolean enabled, String includeDirs, String excludeDirs, String adminIds,
                       String mode, long maxInputBytes, long maxUncompressedBytes,
                       int maxArchiveEntries, int maxCompressionRatio, long maxXmlBytes,
                       int maxPageCount, int maxSheetCount, int maxCellCount, int maxSlideCount,
                       int maxShapeCount, long maxAssetBytes, int parseTimeoutSeconds,
                       int ticketTtlSeconds, boolean allowExternalLinks, boolean allowRemoteFonts,
                       boolean allowMacroFormats, String legacyAction, String nativeFormats,
                       String formats) {
            this.enabled = enabled;
            this.includeDirs = includeDirs;
            this.excludeDirs = excludeDirs;
            this.adminIds = adminIds;
            this.mode = mode;
            this.maxInputBytes = maxInputBytes;
            this.maxUncompressedBytes = maxUncompressedBytes;
            this.maxArchiveEntries = maxArchiveEntries;
            this.maxCompressionRatio = maxCompressionRatio;
            this.maxXmlBytes = maxXmlBytes;
            this.maxPageCount = maxPageCount;
            this.maxSheetCount = maxSheetCount;
            this.maxCellCount = maxCellCount;
            this.maxSlideCount = maxSlideCount;
            this.maxShapeCount = maxShapeCount;
            this.maxAssetBytes = maxAssetBytes;
            this.parseTimeoutSeconds = parseTimeoutSeconds;
            this.ticketTtlSeconds = ticketTtlSeconds;
            this.allowExternalLinks = allowExternalLinks;
            this.allowRemoteFonts = allowRemoteFonts;
            this.allowMacroFormats = allowMacroFormats;
            this.legacyAction = legacyAction;
            this.nativeFormats = nativeFormats;
            this.formats = formats;
        }

        public boolean browserMode() {
            return MODE_BROWSER_JS.equals(OfficeUnitUtil.lower(mode));
        }

        public Set<Integer> includeSet() {
            return OfficeUnitUtil.parseIds(includeDirs);
        }

        public Set<Integer> excludeSet() {
            return OfficeUnitUtil.parseIds(excludeDirs);
        }

        public Set<Integer> adminIds() {
            return OfficeUnitUtil.parseIds(adminIds);
        }

        public Set<String> nativeSet() {
            Set<String> result = new LinkedHashSet<String>();
            String[] values = nativeFormats.split("[,，\\s]+");
            for (String value : values) {
                String ext = OfficeUnitUtil.lower(value);
                if (ext.length() > 0) {
                    result.add(ext);
                }
            }
            return result;
        }

        public Set<String> formatSet() {
            Set<String> result = new LinkedHashSet<String>();
            String[] values = formats.split("[,，\\s]+");
            for (String value : values) {
                String ext = OfficeUnitUtil.lower(value);
                if (ext.length() > 0) {
                    result.add(ext);
                }
            }
            return result;
        }

        public boolean isFormatEnabled(String extension) {
            return formatSet().contains(OfficeUnitUtil.lower(extension));
        }

        private Config withAdminSettings(boolean nextEnabled, String nextInclude,
                                         String nextExclude, String nextFormats) {
            return new Config(nextEnabled, normalizeIds(nextInclude), normalizeIds(nextExclude),
                    adminIds, mode, maxInputBytes, maxUncompressedBytes, maxArchiveEntries,
                    maxCompressionRatio, maxXmlBytes, maxPageCount, maxSheetCount, maxCellCount,
                    maxSlideCount, maxShapeCount, maxAssetBytes, parseTimeoutSeconds,
                    ticketTtlSeconds, allowExternalLinks, allowRemoteFonts, allowMacroFormats,
                    legacyAction, nativeFormats, normalizeFormatList(nextFormats));
        }
    }

    private static final class Snapshot {
        private final Config value;
        private final long stamp;
        private final long length;

        private Snapshot(Config value, long stamp, long length) {
            this.value = value;
            this.stamp = stamp;
            this.length = length;
        }
    }
}
