package com.customization.oaones.officeunit.common;

import weaver.general.BaseBean;

/**
 * Controlled audit and failure logging without file contents or ticket text.
 */
public final class OfficeUnitLog {

    private OfficeUnitLog() {
    }

    public static void info(String message) {
        write("[oaones.officeunit] " + OfficeUnitUtil.s(message));
    }

    public static void error(String message, Throwable throwable) {
        StringBuilder text = new StringBuilder(OfficeUnitUtil.s(message));
        if (throwable != null) {
            text.append(" | ").append(throwable.getClass().getName());
            if (OfficeUnitUtil.trim(throwable.getMessage()).length() > 0) {
                text.append(": ").append(OfficeUnitUtil.trim(throwable.getMessage()));
            }
        }
        write("[oaones.officeunit][ERROR] " + text.toString());
    }

    private static void write(String message) {
        try {
            new BaseBean().writeLog(message);
        } catch (Throwable ignore) {
            // Logging is non-critical.
        }
    }
}
