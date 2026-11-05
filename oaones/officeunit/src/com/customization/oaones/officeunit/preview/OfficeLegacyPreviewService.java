package com.customization.oaones.officeunit.preview;

import com.customization.oaones.officeunit.common.OfficeFileAccessSupport;
import com.customization.oaones.officeunit.common.OfficePreviewException;
import com.customization.oaones.officeunit.common.OfficeUnitConfig;
import com.customization.oaones.officeunit.common.OfficeUnitLog;
import com.customization.oaones.officeunit.common.OfficeUnitUtil;
import org.apache.poi.hslf.usermodel.HSLFSlide;
import org.apache.poi.hslf.usermodel.HSLFSlideShow;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.HWPFDocumentCore;
import org.apache.poi.hwpf.HWPFOldDocument;
import org.apache.poi.hwpf.OldWordFileFormatException;
import org.apache.poi.hwpf.converter.PicturesManager;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.hwpf.usermodel.PictureType;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import weaver.file.ImageFileManager;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.awt.Color;
import java.awt.Dimension;
import java.awt.Graphics2D;
import java.awt.RenderingHints;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.Base64;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.concurrent.Callable;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.Future;
import java.util.concurrent.ThreadFactory;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.TimeoutException;

/**
 * Converts legacy binary Office files with the POI implementation already
 * deployed in E9. DOC becomes HTML; PPT becomes a list of PNG slides.
 */
public final class OfficeLegacyPreviewService {

    private static final int MAX_CACHE_ENTRIES = 16;
    private static final long CACHE_TTL_MILLIS = 10L * 60L * 1000L;
    private static final Object CACHE_LOCK = new Object();
    private static final Map<String, Artifact> CACHE =
            new LinkedHashMap<String, Artifact>(16, 0.75F, true);
    private static long cacheBytes;

    private static final ExecutorService CONVERTERS =
            Executors.newFixedThreadPool(2, new ThreadFactory() {
                @Override
                public Thread newThread(Runnable runnable) {
                    Thread thread = new Thread(runnable,
                            "oaones-officeunit-legacy-converter");
                    thread.setDaemon(true);
                    return thread;
                }
            });

    private OfficeLegacyPreviewService() {
    }

    public static Artifact load(OfficeFileAccessSupport.FileInfo file,
                                OfficeUnitConfig.Config config) {
        String extension = OfficeUnitUtil.lower(file.extension);
        if (!"doc".equals(extension) && !"ppt".equals(extension)) {
            throw OfficePreviewException.unsupported("OFFICE_PREVIEW_UNSUPPORTED");
        }
        String key = cacheKey(file, extension);
        long now = System.currentTimeMillis();
        synchronized (CACHE_LOCK) {
            purgeExpiredLocked(now);
            Artifact cached = CACHE.get(key);
            if (cached != null && cached.accepts(config)) {
                return cached;
            }
            CACHE.remove(key);
        }

        final byte[] source = readSource(file, config.maxInputBytes);
        Future<Artifact> future = CONVERTERS.submit(new Callable<Artifact>() {
            @Override
            public Artifact call() {
                return "doc".equals(extension)
                        ? convertDoc(source, config) : convertPpt(source, config);
            }
        });
        Artifact artifact;
        try {
            artifact = future.get(Math.max(1, config.parseTimeoutSeconds),
                    TimeUnit.SECONDS);
        } catch (TimeoutException e) {
            future.cancel(true);
            throw new OfficePreviewException("OFFICE_PREVIEW_PARSE_TIMEOUT",
                    "旧版 Office 文档转换超时，请稍后重试", 504);
        } catch (InterruptedException e) {
            future.cancel(true);
            Thread.currentThread().interrupt();
            throw new OfficePreviewException("OFFICE_PREVIEW_PARSE_FAILED",
                    "旧版 Office 文档转换已中断", 503);
        } catch (ExecutionException e) {
            Throwable cause = e.getCause();
            if (cause instanceof OfficePreviewException) {
                throw (OfficePreviewException) cause;
            }
            OfficeUnitLog.error("旧版 Office 转换失败 imageFileId="
                    + file.imageFileId, cause);
            throw new OfficePreviewException("OFFICE_PREVIEW_PARSE_FAILED",
                    "旧版 Office 文档解析失败，请下载后查看", 422);
        }

        if (artifact.sizeBytes > config.maxAssetBytes) {
            throw OfficePreviewException.unsupported(
                    "OFFICE_PREVIEW_ARCHIVE_TOO_LARGE");
        }
        synchronized (CACHE_LOCK) {
            Artifact current = CACHE.get(key);
            if (current != null && current.accepts(config)) {
                return current;
            }
            putCacheLocked(key, artifact, config.maxAssetBytes);
        }
        return artifact;
    }

    public static void shutdown() {
        CONVERTERS.shutdownNow();
        synchronized (CACHE_LOCK) {
            CACHE.clear();
            cacheBytes = 0L;
        }
    }

    private static byte[] readSource(OfficeFileAccessSupport.FileInfo file,
                                     long maxBytes) {
        InputStream input = null;
        try {
            input = ImageFileManager.getInputStreamById(file.imageFileId);
            if (input == null) {
                throw OfficePreviewException.fileNotFound();
            }
            ByteArrayOutputStream output = new ByteArrayOutputStream(
                    (int) Math.min(65536L, Math.max(1024L, file.fileSize)));
            byte[] buffer = new byte[16384];
            long total = 0L;
            int count;
            while ((count = input.read(buffer)) >= 0) {
                total += count;
                if (total > maxBytes || total > file.fileSize) {
                    throw OfficePreviewException.unsupported(
                            "OFFICE_PREVIEW_FILE_TOO_LARGE");
                }
                output.write(buffer, 0, count);
            }
            if (total != file.fileSize) {
                throw new OfficePreviewException("OFFICE_PREVIEW_FILE_NOT_FOUND",
                        "文档文件已变化，请重新打开预览", 409);
            }
            return output.toByteArray();
        } catch (OfficePreviewException e) {
            throw e;
        } catch (Exception e) {
            OfficeUnitLog.error("读取旧版 Office 文件失败 imageFileId="
                    + file.imageFileId, e);
            throw OfficePreviewException.fileNotFound();
        } finally {
            OfficeUnitUtil.closeQuietly(input);
        }
    }

    private static Artifact convertDoc(byte[] source,
                                       OfficeUnitConfig.Config config) {
        HWPFDocumentCore word = null;
        POIFSFileSystem fs = null;
        try {
            fs = new POIFSFileSystem(new ByteArrayInputStream(source));
            try {
                word = new HWPFDocument(fs);
            } catch (OldWordFileFormatException oldFormat) {
                word = new HWPFOldDocument(fs);
            }

            org.w3c.dom.Document html = DocumentBuilderFactory.newInstance()
                    .newDocumentBuilder().newDocument();
            WordToHtmlConverter converter = new WordToHtmlConverter(html);
            converter.setPicturesManager(new EmbeddedPicturesManager(
                    config.maxAssetBytes));
            converter.processDocument(word);

            Transformer transformer = TransformerFactory.newInstance().newTransformer();
            transformer.setOutputProperty(OutputKeys.METHOD, "html");
            transformer.setOutputProperty(OutputKeys.ENCODING, "UTF-8");
            transformer.setOutputProperty(OutputKeys.OMIT_XML_DECLARATION, "yes");
            ByteArrayOutputStream output = new ByteArrayOutputStream();
            transformer.transform(new DOMSource(converter.getDocument()),
                    new StreamResult(output));
            byte[] htmlBytes = output.toByteArray();
            return Artifact.html(new String(htmlBytes, StandardCharsets.UTF_8));
        } catch (OfficePreviewException e) {
            throw e;
        } catch (Exception e) {
            OfficeUnitLog.error("DOC 转 HTML 失败", e);
            throw new OfficePreviewException("OFFICE_PREVIEW_PARSE_FAILED",
                    "Word 文档解析失败，请下载后查看", 422);
        } finally {
            OfficeUnitUtil.closeQuietly(word);
            OfficeUnitUtil.closeQuietly(fs);
        }
    }

    private static Artifact convertPpt(byte[] source,
                                       OfficeUnitConfig.Config config) {
        HSLFSlideShow show = null;
        try {
            show = new HSLFSlideShow(new ByteArrayInputStream(source));
            List<HSLFSlide> slides = show.getSlides();
            if (slides.isEmpty()) {
                throw new OfficePreviewException("OFFICE_PREVIEW_PARSE_FAILED",
                        "PowerPoint 文档中没有可显示的幻灯片", 422);
            }
            if (slides.size() > config.maxSlideCount) {
                throw OfficePreviewException.unsupported(
                        "OFFICE_PREVIEW_ARCHIVE_TOO_LARGE");
            }

            Dimension logicalSize = show.getPageSize();
            int logicalWidth = logicalSize == null || logicalSize.width <= 0
                    ? 720 : logicalSize.width;
            int logicalHeight = logicalSize == null || logicalSize.height <= 0
                    ? 540 : logicalSize.height;
            double renderScale = Math.min(2D,
                    Math.min(1600D / logicalWidth, 1200D / logicalHeight));
            renderScale = Math.max(1D, renderScale);
            int imageWidth = (int) Math.round(logicalWidth * renderScale);
            int imageHeight = (int) Math.round(logicalHeight * renderScale);

            List<byte[]> images = new ArrayList<byte[]>(slides.size());
            long totalBytes = 0L;
            for (HSLFSlide slide : slides) {
                BufferedImage image = new BufferedImage(
                        imageWidth, imageHeight, BufferedImage.TYPE_INT_RGB);
                Graphics2D graphics = image.createGraphics();
                try {
                    graphics.setRenderingHint(RenderingHints.KEY_ANTIALIASING,
                            RenderingHints.VALUE_ANTIALIAS_ON);
                    graphics.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING,
                            RenderingHints.VALUE_TEXT_ANTIALIAS_ON);
                    graphics.setRenderingHint(RenderingHints.KEY_RENDERING,
                            RenderingHints.VALUE_RENDER_QUALITY);
                    graphics.setPaint(Color.WHITE);
                    graphics.fillRect(0, 0, imageWidth, imageHeight);
                    graphics.scale(renderScale, renderScale);
                    slide.draw(graphics);
                } finally {
                    graphics.dispose();
                }
                ByteArrayOutputStream output = new ByteArrayOutputStream();
                if (!javax.imageio.ImageIO.write(image, "png", output)) {
                    throw new IOException("PNG encoder unavailable");
                }
                byte[] png = output.toByteArray();
                totalBytes += png.length;
                if (totalBytes > config.maxAssetBytes) {
                    throw OfficePreviewException.unsupported(
                            "OFFICE_PREVIEW_ARCHIVE_TOO_LARGE");
                }
                images.add(png);
            }
            return Artifact.slides(logicalWidth, logicalHeight, images);
        } catch (OfficePreviewException e) {
            throw e;
        } catch (Exception e) {
            OfficeUnitLog.error("PPT 转 PNG 失败", e);
            throw new OfficePreviewException("OFFICE_PREVIEW_PARSE_FAILED",
                    "PowerPoint 文档解析失败，请下载后查看", 422);
        } finally {
            OfficeUnitUtil.closeQuietly(show);
        }
    }

    private static String cacheKey(OfficeFileAccessSupport.FileInfo file,
                                   String extension) {
        return extension + ":" + file.docId + ":" + file.imageFileId + ":"
                + file.versionId + ":" + file.fileSize;
    }

    private static void putCacheLocked(String key, Artifact artifact, long maxBytes) {
        Artifact previous = CACHE.remove(key);
        if (previous != null) {
            cacheBytes -= previous.sizeBytes;
        }
        CACHE.put(key, artifact);
        cacheBytes += artifact.sizeBytes;
        while ((cacheBytes > maxBytes || CACHE.size() > MAX_CACHE_ENTRIES)
                && !CACHE.isEmpty()) {
            Iterator<Map.Entry<String, Artifact>> iterator =
                    CACHE.entrySet().iterator();
            if (!iterator.hasNext()) {
                break;
            }
            Artifact removed = iterator.next().getValue();
            iterator.remove();
            cacheBytes -= removed.sizeBytes;
        }
        cacheBytes = Math.max(0L, cacheBytes);
    }

    private static void purgeExpiredLocked(long now) {
        Iterator<Map.Entry<String, Artifact>> iterator =
                CACHE.entrySet().iterator();
        while (iterator.hasNext()) {
            Artifact artifact = iterator.next().getValue();
            if (now - artifact.createdAt > CACHE_TTL_MILLIS) {
                iterator.remove();
                cacheBytes -= artifact.sizeBytes;
            }
        }
        cacheBytes = Math.max(0L, cacheBytes);
    }

    private static final class EmbeddedPicturesManager implements PicturesManager {
        private final long maxBytes;
        private long totalBytes;

        private EmbeddedPicturesManager(long maxBytes) {
            this.maxBytes = maxBytes;
        }

        @Override
        public String savePicture(byte[] content, PictureType pictureType,
                                  String suggestedName, float widthInches,
                                  float heightInches) {
            if (content == null || content.length == 0 || pictureType == null) {
                return "";
            }
            String mime = OfficeUnitUtil.lower(pictureType.getMime());
            if (!isBrowserImageMime(mime)) {
                return "";
            }
            totalBytes += content.length;
            if (totalBytes > maxBytes) {
                throw OfficePreviewException.unsupported(
                        "OFFICE_PREVIEW_ARCHIVE_TOO_LARGE");
            }
            return "data:" + mime + ";base64,"
                    + Base64.getEncoder().encodeToString(content);
        }
    }

    private static boolean isBrowserImageMime(String mime) {
        String value = OfficeUnitUtil.lower(mime).toLowerCase(Locale.ROOT);
        return "image/png".equals(value)
                || "image/jpeg".equals(value)
                || "image/jpg".equals(value)
                || "image/gif".equals(value)
                || "image/bmp".equals(value)
                || "image/webp".equals(value);
    }

    public static final class Artifact {
        public final String kind;
        public final String html;
        public final int slideWidth;
        public final int slideHeight;
        public final List<byte[]> slides;
        public final long sizeBytes;
        public final long createdAt;

        private Artifact(String kind, String html, int slideWidth, int slideHeight,
                         List<byte[]> slides, long sizeBytes) {
            this.kind = kind;
            this.html = html;
            this.slideWidth = slideWidth;
            this.slideHeight = slideHeight;
            this.slides = slides;
            this.sizeBytes = sizeBytes;
            this.createdAt = System.currentTimeMillis();
        }

        private static Artifact html(String html) {
            return new Artifact("doc", html == null ? "" : html, 0, 0, null,
                    html == null ? 0L : html.getBytes(StandardCharsets.UTF_8).length);
        }

        private static Artifact slides(int width, int height, List<byte[]> slides) {
            long bytes = 0L;
            for (byte[] slide : slides) {
                bytes += slide.length;
            }
            return new Artifact("ppt", null, width, height, slides, bytes);
        }

        private boolean accepts(OfficeUnitConfig.Config config) {
            if ("doc".equals(kind)) {
                return sizeBytes <= config.maxAssetBytes;
            }
            return slides != null && slides.size() <= config.maxSlideCount
                    && sizeBytes <= config.maxAssetBytes;
        }
    }
}
