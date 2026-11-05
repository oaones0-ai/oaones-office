<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8" %>
<%@ page import="com.customization.oaones.officeunit.common.OfficeUnitAdmin" %>
<%@ page import="weaver.hrm.HrmUserVarify" %>
<%@ page import="weaver.hrm.User" %>
<%@ include file="/oaones/officeunit/page/assets.jsp" %>
<%
    request.setCharacterEncoding("UTF-8");
    User user = HrmUserVarify.getUser(request, response);
    if (user == null) {
        response.sendRedirect("/wui/index.html");
        return;
    }
    boolean canAdmin = OfficeUnitAdmin.isAdmin(user.getUID());
%>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>文档在线预览设置</title>
    <link rel="stylesheet" href="/oaones/officeunit/css/admin.css?v=<%=aiAssetV%>">
</head>
<body>
<div class="oa-officeunit-admin">
    <% if (!canAdmin) { %>
    <section class="ou-card ou-forbidden">
        <h2>无权维护</h2>
        <p>只有系统管理员或 <code>officeunit.preview.admin_ids</code> 中的人员可以修改此配置。</p>
    </section>
    <% } else { %>
    <main id="officeUnitAdmin" data-api="/oaones/officeunit/page/api.jsp">
        <section class="ou-card ou-warning">
            <div class="ou-warning-icon">!</div>
            <div>
                <h2>下载权限提示</h2>
                <p>使用该功能，可能导致文件对于员工是可下载的（另存为），泛微的“不可下载”设置会失效。请确认知识文档目录范围与保密要求后再启用。</p>
            </div>
        </section>

        <section class="ou-card">
            <div class="ou-card-title">
                <div>
                    <p class="ou-eyebrow">文档即时预览</p>
                    <h2>预览设置</h2>
                </div>
                <button type="button" class="ou-save-button" id="ouSave">保存设置</button>
            </div>

            <section class="ou-settings-section" aria-labelledby="ouScopeTitle">
                <div class="ou-section-title">
                    <div>
                        <h3 id="ouScopeTitle">预览生效范围</h3>
                        <p>设置哪些知识文档目录中的文件可以使用即时预览。</p>
                    </div>
                </div>
                <p class="ou-help">“包含目录”不选择时代表全部知识文档目录；排除目录优先于包含目录。目录直接指向泛微知识文档目录（<code>DocSecCategory</code>），可选择任意目录。</p>
                <div class="ou-scope-fields">
                    <div class="ou-scope-field">
                        <div class="ou-scope-field-head">
                            <div>
                                <h3>包含目录</h3>
                                <small>不选 = 全部目录</small>
                            </div>
                        </div>
                        <button type="button" class="ou-scope-trigger" data-scope-trigger="include"
                                aria-haspopup="dialog" aria-expanded="false">
                            <span class="ou-scope-trigger-text" id="ouIncludeSummary">正在读取目录…</span>
                            <span class="ou-scope-trigger-action">选择 ›</span>
                        </button>
                    </div>
                    <div class="ou-scope-field">
                        <div class="ou-scope-field-head">
                            <div>
                                <h3>排除目录</h3>
                                <small>优先于包含目录</small>
                            </div>
                        </div>
                        <button type="button" class="ou-scope-trigger" data-scope-trigger="exclude"
                                aria-haspopup="dialog" aria-expanded="false">
                            <span class="ou-scope-trigger-text" id="ouExcludeSummary">正在读取目录…</span>
                            <span class="ou-scope-trigger-action">选择 ›</span>
                        </button>
                    </div>
                </div>
            </section>

            <section class="ou-settings-section ou-format-section" aria-labelledby="ouFormatsTitle">
                <div class="ou-section-title ou-format-title">
                    <div>
                        <div class="ou-title-with-tip">
                            <h3 id="ouFormatsTitle">预览文件格式</h3>
                            <button type="button" class="ou-help-tip"
                                    aria-label="格式清单说明：不在下方清单中的文件，将使用默认格式展示"
                                    data-tooltip="不在下方清单中的文件，将使用默认格式展示">?</button>
                        </div>
                        <p>按文件格式分别启用预览。</p>
                    </div>
                </div>
                <div class="ou-format-grid" id="ouFormatGrid" aria-live="polite">
                    <p class="ou-format-loading">正在读取格式配置…</p>
                </div>
            </section>
        </section>

        <footer class="ou-footer">
            <span>配置文件：<code>WEB-INF/prop/oaones.officeunit.properties</code></span>
        </footer>
        <div class="ou-toast" id="ouToast" role="status" aria-live="polite"></div>
        <div class="ou-picker-layer" id="ouDirectoryPicker" hidden>
            <section class="ou-picker-dialog" role="dialog" aria-modal="true"
                     aria-labelledby="ouPickerTitle">
                <div class="ou-picker-head">
                    <div>
                        <h3 id="ouPickerTitle">选择目录</h3>
                        <small id="ouPickerDescription"></small>
                    </div>
                    <button type="button" class="ou-picker-close" data-picker-close
                            aria-label="关闭">×</button>
                </div>
                <div class="ou-picker-toolbar">
                    <input type="search" class="ou-picker-search" id="ouPickerSearch"
                           placeholder="输入目录名称或 ID 筛选">
                    <label class="ou-selected-only">
                        <input type="checkbox" id="ouPickerSelectedOnly">
                        <span>仅显示已选择目录</span>
                    </label>
                    <button type="button" class="ou-picker-all" id="ouPickerAll">全选</button>
                    <span class="ou-selected-count" id="ouPickerToolbarCount">已选 0 项</span>
                </div>
                <div class="ou-picker-tree" id="ouDirectoryTree"></div>
                <div class="ou-picker-foot">
                    <span class="ou-selected-count" id="ouPickerFootCount">已选 0 项</span>
                    <div class="ou-picker-actions">
                        <button type="button" class="ou-secondary-button" id="ouPickerClear">清空已选</button>
                        <button type="button" class="ou-secondary-button" data-picker-close>取消</button>
                        <button type="button" class="ou-primary-button" id="ouPickerConfirm">确定</button>
                    </div>
                </div>
            </section>
        </div>
    </main>
    <script src="/oaones/scripts/js/tree.js?v=<%=aiAssetV%>"></script>
    <script src="/oaones/officeunit/js/admin.js?v=<%=aiAssetV%>"></script>
    <% } %>
</div>
</body>
</html>
