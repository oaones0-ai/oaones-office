package com.customization.oaones.officeunit.preview;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.customization.oaones.officeunit.common.OfficeFormatPolicy;
import com.customization.oaones.officeunit.common.OfficeUnitAdmin;
import com.customization.oaones.officeunit.common.OfficeUnitConfig;
import com.customization.oaones.officeunit.common.OfficeUnitUtil;
import weaver.conn.RecordSet;

import javax.servlet.http.HttpServletRequest;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Set;

/**
 * Settings actions used by the preserved admin.jsp page.
 */
public final class OfficeUnitAdminController {

    public JSONObject handle(String action, HttpServletRequest request, int userId) {
        OfficeUnitAdmin.requireAdmin(userId);
        String name = OfficeUnitUtil.trim(action);
        if ("admin.config".equals(name)) {
            return ok(configData(OfficeUnitConfig.load()));
        }
        if ("admin.directories".equals(name)) {
            return ok(directories());
        }
        if ("admin.save".equals(name)) {
            JSONObject payload = parsePayload(request.getParameter("payload"));
            OfficeUnitConfig.Config current = OfficeUnitConfig.load();
            String includeDirs = ids(payload.get("includeDirs"));
            String excludeDirs = ids(payload.get("excludeDirs"));
            String enabledFormats = payload.containsKey("enabledFormats")
                    ? formats(payload.get("enabledFormats")) : current.formats;
            boolean enabled = enabledFormats.length() > 0
                    && OfficeUnitUtil.b(payload.get("enabled"), true);
            return ok(configData(OfficeUnitConfig.save(enabled, includeDirs, excludeDirs,
                    enabledFormats)));
        }
        return fail("未知 action");
    }

    private JSONObject configData(OfficeUnitConfig.Config config) {
        JSONObject data = new JSONObject();
        data.put("enabled", Boolean.valueOf(config.enabled));
        data.put("includeDirs", config.includeDirs);
        data.put("excludeDirs", config.excludeDirs);
        data.put("oaDefaultPreviewEnabled",
                Boolean.valueOf(OfficeUnitConfig.isOaDefaultPreviewEnabled()));
        data.put("configFile", "WEB-INF/prop/" + OfficeUnitConfig.PROP_FILE);
        data.put("mode", config.mode);
        data.put("convertibleFormats", OfficeFormatPolicy.convertibleFormats());
        data.put("textFormats", OfficeFormatPolicy.textFormats());
        data.put("enabledFormats", new java.util.ArrayList<String>(config.formatSet()));
        data.put("nativeFormats", OfficeFormatPolicy.fallbackFormats(config));
        data.put("defaultNativeFormats", new java.util.ArrayList<String>(config.nativeSet()));
        data.put("warning",
                "浏览器需要读取原始 Office 文件，泛微“不可下载”设置不能阻止具备阅读权限的"
                        + "用户通过浏览器开发工具取得文件。请严格控制目录范围。");
        return data;
    }

    private JSONArray directories() {
        JSONArray data = new JSONArray();
        RecordSet rs = new RecordSet();
        rs.executeQuery("select id, parentid, categoryname from DocSecCategory"
                + " order by parentid, id");
        while (rs.next()) {
            JSONObject item = new JSONObject();
            item.put("id", Integer.valueOf(rs.getInt("id")));
            item.put("parentId", Integer.valueOf(rs.getInt("parentid")));
            item.put("name", OfficeUnitUtil.trim(rs.getString("categoryname")));
            data.add(item);
        }
        return data;
    }

    private String ids(Object value) {
        if (value instanceof List) {
            Set<Integer> ids = new LinkedHashSet<Integer>();
            for (Object item : (List<?>) value) {
                int id = OfficeUnitUtil.i(item, 0);
                if (id > 0) {
                    ids.add(Integer.valueOf(id));
                }
            }
            return OfficeUnitUtil.joinIds(ids);
        }
        return OfficeUnitUtil.joinIds(new LinkedHashSet<Integer>(
                OfficeUnitUtil.parseIds(String.valueOf(value))));
    }

    private String formats(Object value) {
        Set<String> allowed = new LinkedHashSet<String>(
                OfficeFormatPolicy.convertibleFormats());
        Set<String> selected = new LinkedHashSet<String>();
        if (value instanceof List) {
            for (Object item : (List<?>) value) {
                addFormat(selected, allowed, item);
            }
        } else if (value != null) {
            String[] values = String.valueOf(value).split("[,，\\s]+");
            for (String item : values) {
                addFormat(selected, allowed, item);
            }
        }
        return OfficeUnitUtil.join(new java.util.ArrayList<String>(selected), ",");
    }

    private void addFormat(Set<String> selected, Set<String> allowed, Object value) {
        String extension = OfficeUnitUtil.lower(String.valueOf(value));
        if (extension.startsWith(".")) {
            extension = extension.substring(1);
        }
        if (allowed.contains(extension)) {
            selected.add(extension);
        }
    }

    private JSONObject parsePayload(String payload) {
        try {
            JSONObject value = JSONObject.parseObject(payload);
            return value == null ? new JSONObject() : value;
        } catch (Exception e) {
            throw new IllegalArgumentException("请求参数格式不正确");
        }
    }

    private JSONObject ok(Object data) {
        JSONObject result = new JSONObject();
        result.put("code", Integer.valueOf(0));
        result.put("msg", "success");
        result.put("data", data);
        return result;
    }

    private JSONObject fail(String message) {
        JSONObject result = new JSONObject();
        result.put("code", Integer.valueOf(-1));
        result.put("msg", message);
        return result;
    }
}
