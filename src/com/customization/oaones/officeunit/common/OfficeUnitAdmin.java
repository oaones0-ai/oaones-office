package com.customization.oaones.officeunit.common;

/**
 * Administrator policy for the OfficeUnit settings page.
 */
public final class OfficeUnitAdmin {

    private OfficeUnitAdmin() {
    }

    public static boolean isAdmin(int userId) {
        if (userId <= 0) {
            return false;
        }
        if (userId == 1) {
            return true;
        }
        return OfficeUnitConfig.load().adminIds().contains(Integer.valueOf(userId));
    }

    public static void requireAdmin(int userId) {
        if (!isAdmin(userId)) {
            throw new SecurityException("只有 Office 预览管理员可以执行此操作");
        }
    }
}
