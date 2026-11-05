package com.customization.oaones.officeunit.common;

import weaver.conn.RecordSet;

import java.util.Set;

/**
 * Matches a document's exact knowledge-directory ID against include/exclude IDs.
 */
public final class KnowledgeDirectoryScope {

    private KnowledgeDirectoryScope() {
    }

    public static boolean allows(int docId, OfficeUnitConfig.Config config) {
        Set<Integer> includes = config.includeSet();
        Set<Integer> excludes = config.excludeSet();
        if (includes.isEmpty() && excludes.isEmpty()) {
            return true;
        }
        int categoryId = documentCategory(docId);
        if (categoryId <= 0) {
            return false;
        }
        if (excludes.contains(Integer.valueOf(categoryId))) {
            return false;
        }
        return includes.isEmpty() || includes.contains(Integer.valueOf(categoryId));
    }

    private static int documentCategory(int docId) {
        RecordSet rs = new RecordSet();
        rs.executeQuery("select seccategory from DocDetail where id = ?", Integer.valueOf(docId));
        return rs.next() ? rs.getInt("seccategory") : 0;
    }
}
