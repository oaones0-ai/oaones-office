package com.customization.oaones.officeunit.preview;

import com.alibaba.fastjson.JSONObject;
import com.customization.oaones.officeunit.common.OfficeFileAccessSupport;
import com.customization.oaones.officeunit.common.OfficePreviewException;
import com.customization.oaones.officeunit.common.OfficeUnitConfig;
import com.customization.oaones.officeunit.common.OfficeUnitLog;
import com.customization.oaones.officeunit.common.OfficeUnitUtil;
import weaver.hrm.User;

import javax.crypto.Mac;
import javax.crypto.spec.SecretKeySpec;
import java.io.File;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;
import java.nio.file.attribute.PosixFilePermission;
import java.security.MessageDigest;
import java.security.SecureRandom;
import java.util.Base64;
import java.util.EnumSet;
import java.util.Set;
import java.util.UUID;

/**
 * Creates short-lived, user/file/version-bound download tickets.
 */
public final class OfficePreviewTicketService {

    public static final String PURPOSE = "office-preview";
    private static final String KEY_FILE = ".oaones.officeunit.ticket.key";
    private static final String HMAC_ALGORITHM = "HmacSHA256";
    private static final SecureRandom RANDOM = new SecureRandom();
    private static volatile byte[] volatileFallbackKey;

    public Ticket issue(OfficeFileAccessSupport.FileInfo file, User user,
                        OfficeUnitConfig.Config config) {
        long now = System.currentTimeMillis();
        long expiresAt = now + Math.max(60, config.ticketTtlSeconds) * 1000L;
        JSONObject claims = new JSONObject();
        claims.put("v", Integer.valueOf(1));
        claims.put("p", PURPOSE);
        claims.put("u", Integer.valueOf(user.getUID()));
        claims.put("d", Integer.valueOf(file.docId));
        claims.put("f", Integer.valueOf(file.imageFileId));
        claims.put("r", Integer.valueOf(file.versionId));
        claims.put("s", Long.valueOf(file.fileSize));
        claims.put("x", file.extension);
        claims.put("n", UUID.randomUUID().toString());
        claims.put("t", Long.valueOf(now));
        claims.put("e", Long.valueOf(expiresAt));
        byte[] data = claims.toJSONString().getBytes(StandardCharsets.UTF_8);
        String encodedData = Base64.getUrlEncoder().withoutPadding().encodeToString(data);
        String encodedSignature = Base64.getUrlEncoder().withoutPadding()
                .encodeToString(sign(encodedData, key()));
        return new Ticket(encodedData + "." + encodedSignature, expiresAt);
    }

    public Claims verify(String token, int currentUserId) {
        String text = OfficeUnitUtil.trim(token);
        if (text.length() == 0 || text.length() > 4096) {
            throw expired();
        }
        int dot = text.indexOf('.');
        if (dot <= 0 || dot != text.lastIndexOf('.') || dot == text.length() - 1) {
            throw expired();
        }
        String encodedData = text.substring(0, dot);
        String encodedSignature = text.substring(dot + 1);
        byte[] signature;
        byte[] expected;
        try {
            signature = Base64.getUrlDecoder().decode(encodedSignature);
            expected = sign(encodedData, key());
        } catch (Exception e) {
            throw expired();
        }
        if (!MessageDigest.isEqual(signature, expected)) {
            throw expired();
        }

        JSONObject data;
        try {
            data = JSONObject.parseObject(new String(
                    Base64.getUrlDecoder().decode(encodedData), StandardCharsets.UTF_8));
        } catch (Exception e) {
            throw expired();
        }
        if (data == null) {
            throw expired();
        }
        Claims claims = new Claims();
        claims.userId = intValue(data, "u");
        claims.docId = intValue(data, "d");
        claims.imageFileId = intValue(data, "f");
        claims.versionId = intValue(data, "r");
        claims.fileSize = longValue(data, "s");
        claims.extension = OfficeUnitUtil.lower(data.getString("x"));
        claims.issuedAt = longValue(data, "t");
        claims.expiresAt = longValue(data, "e");
        claims.purpose = OfficeUnitUtil.trim(data.getString("p"));
        claims.nonce = OfficeUnitUtil.trim(data.getString("n"));

        long now = System.currentTimeMillis();
        if (!PURPOSE.equals(claims.purpose)
                || claims.userId != currentUserId
                || claims.docId <= 0
                || claims.imageFileId <= 0
                || claims.fileSize < 0L
                || claims.issuedAt > now + 60000L
                || claims.expiresAt <= now
                || claims.expiresAt > claims.issuedAt + 3600000L
                || claims.extension.length() == 0
                || claims.nonce.length() < 16) {
            throw expired();
        }
        return claims;
    }

    private static byte[] sign(String data, byte[] key) {
        try {
            Mac mac = Mac.getInstance(HMAC_ALGORITHM);
            mac.init(new SecretKeySpec(key, HMAC_ALGORITHM));
            return mac.doFinal(data.getBytes(StandardCharsets.UTF_8));
        } catch (Exception e) {
            throw new IllegalStateException("无法签发预览票据", e);
        }
    }

    private static byte[] key() {
        File directory = OfficeUnitConfig.configFile().getParentFile();
        File file = new File(directory, KEY_FILE);
        if (!file.exists()) {
            synchronized (OfficePreviewTicketService.class) {
                if (!file.exists()) {
                    byte[] generated = new byte[32];
                    RANDOM.nextBytes(generated);
                    try {
                        if (directory != null && !directory.exists() && !directory.mkdirs()) {
                            throw new IllegalStateException("配置目录不可用");
                        }
                        Path path = file.toPath();
                        Files.write(path, Base64.getEncoder().encode(generated),
                                StandardOpenOption.CREATE_NEW, StandardOpenOption.WRITE);
                        restrictPermissions(path);
                    } catch (Exception e) {
                        OfficeUnitLog.error("无法持久化预览票据密钥，使用进程内密钥", e);
                        if (volatileFallbackKey == null) {
                            volatileFallbackKey = generated;
                        }
                    }
                }
            }
        }
        if (file.exists() && file.isFile()) {
            try {
                byte[] encoded = Files.readAllBytes(file.toPath());
                byte[] decoded = Base64.getDecoder().decode(encoded);
                if (decoded.length >= 32) {
                    return decoded;
                }
            } catch (Exception e) {
                OfficeUnitLog.error("读取预览票据密钥失败", e);
            }
        }
        synchronized (OfficePreviewTicketService.class) {
            if (volatileFallbackKey == null) {
                byte[] generated = new byte[32];
                RANDOM.nextBytes(generated);
                volatileFallbackKey = generated;
            }
            return volatileFallbackKey.clone();
        }
    }

    private static void restrictPermissions(Path path) {
        try {
            Set<PosixFilePermission> permissions = EnumSet.of(
                    PosixFilePermission.OWNER_READ,
                    PosixFilePermission.OWNER_WRITE);
            Files.setPosixFilePermissions(path, permissions);
        } catch (Exception ignore) {
            // Windows and non-POSIX filesystems are protected by WEB-INF placement.
        }
    }

    private static int intValue(JSONObject data, String name) {
        Integer value = data.getInteger(name);
        return value == null ? 0 : value.intValue();
    }

    private static long longValue(JSONObject data, String name) {
        Long value = data.getLong(name);
        return value == null ? 0L : value.longValue();
    }

    private static OfficePreviewException expired() {
        return new OfficePreviewException("OFFICE_PREVIEW_TICKET_EXPIRED",
                OfficePreviewException.defaultMessage("OFFICE_PREVIEW_TICKET_EXPIRED"), 403);
    }

    public static final class Ticket {
        public final String token;
        public final long expiresAt;

        private Ticket(String token, long expiresAt) {
            this.token = token;
            this.expiresAt = expiresAt;
        }
    }

    public static final class Claims {
        public int userId;
        public int docId;
        public int imageFileId;
        public int versionId;
        public long fileSize;
        public String extension;
        public String purpose;
        public String nonce;
        public long issuedAt;
        public long expiresAt;
    }
}
