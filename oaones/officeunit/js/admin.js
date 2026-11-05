(function () {
    'use strict';

    var root = document.getElementById('officeUnitAdmin');
    if (!root) {
        return;
    }

    var apiUrl = root.getAttribute('data-api');
    var FORMAT_INFO = {
        doc: {
            label: 'DOC',
            level: '标准支持',
            technology: '服务端 Apache POI 转 HTML',
            tip: '旧版 Word 文档由服务端 Apache POI 转换为受控 HTML。'
        },
        docx: {
            label: 'DOCX',
            level: '标准支持',
            technology: '浏览器端 OOXML 渲染（docx-preview）',
            tip: 'Word OOXML 文档在浏览器内解析，并通过 docx-preview 渲染。'
        },
        xlsx: {
            label: 'XLSX',
            level: '标准支持',
            technology: '浏览器端 SheetJS 解析',
            tip: '工作簿在浏览器内通过 SheetJS 解析，再使用本地表格渲染器展示。'
        },
        xls: {
            label: 'XLS',
            level: '标准支持',
            technology: '浏览器端 SheetJS 解析',
            tip: '旧版工作簿在浏览器内通过 SheetJS 解析，再使用本地表格渲染器展示。'
        },
        csv: {
            label: 'CSV',
            level: '标准支持',
            technology: '浏览器端 SheetJS 解析',
            tip: 'CSV 文本在浏览器内通过 SheetJS 解析，再使用本地表格渲染器展示。'
        },
        tsv: {
            label: 'TSV',
            level: '标准支持',
            technology: '浏览器端 SheetJS 解析',
            tip: 'TSV 文本在浏览器内通过 SheetJS 解析，再使用本地表格渲染器展示。'
        },
        ppt: {
            label: 'PPT',
            level: '标准支持',
            technology: '服务端 Apache POI 转 PNG',
            tip: '旧版 PowerPoint 文档由服务端 Apache POI 逐页转换为 PNG。'
        },
        pptx: {
            label: 'PPTX',
            level: '标准支持',
            technology: '浏览器端 OOXML 解析渲染',
            tip: 'PowerPoint OOXML 文档在浏览器内解析，并通过本地幻灯片渲染器展示。'
        },
        dotx: {
            label: 'DOTX',
            level: '尽力支持',
            technology: '浏览器端 OOXML 渲染（docx-preview）',
            tip: 'Word 模板按 DOCX OOXML 方式在浏览器内解析渲染，复杂模板元素可能降级。'
        },
        xlsb: {
            label: 'XLSB',
            level: '尽力支持',
            technology: '浏览器端 SheetJS 解析',
            tip: '二进制工作簿在浏览器内通过 SheetJS 解析，复杂样式可能降级。'
        },
        ods: {
            label: 'ODS',
            level: '尽力支持',
            technology: '浏览器端 SheetJS 解析',
            tip: 'OpenDocument 表格在浏览器内通过 SheetJS 解析，复杂样式可能降级。'
        },
        potx: {
            label: 'POTX',
            level: '尽力支持',
            technology: '浏览器端 OOXML 解析渲染',
            tip: 'PowerPoint 模板按 PPTX OOXML 方式在浏览器内解析渲染，复杂模板元素可能降级。'
        },
        ppsx: {
            label: 'PPSX',
            level: '尽力支持',
            technology: '浏览器端 OOXML 解析渲染',
            tip: 'PowerPoint 放映文件按 PPTX OOXML 方式在浏览器内解析渲染，动画不会执行。'
        },
        ofd: {
            label: 'OFD',
            level: '标准支持',
            technology: '浏览器端 OFDRW Canvas 渲染',
            tip: '版式文档在浏览器内通过 OFDRW 阅读器渲染，支持发票、签章、缩放和页面导航。'
        }
    };
    var state = {
        config: null,
        directories: [],
        formats: {},
        textFormats: [],
        selected: {
            include: {},
            exclude: {}
        },
        picker: {
            scope: '',
            draft: {},
            keyword: '',
            selectedOnly: false,
            expanded: {},
            trigger: null
        },
        saving: false
    };
    var toastTimer = null;

    function request(action, payload, done) {
        var xhr = new XMLHttpRequest();
        var method = payload ? 'POST' : 'GET';
        var url = apiUrl + '?action=' + encodeURIComponent(action);
        xhr.open(method, url, true);
        xhr.setRequestHeader('X-Requested-With', 'XMLHttpRequest');
        if (payload) {
            xhr.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded;charset=UTF-8');
        }
        xhr.onreadystatechange = function () {
            if (xhr.readyState !== 4) {
                return;
            }
            var response;
            try {
                response = JSON.parse(xhr.responseText || '{}');
            } catch (ignore) {
                response = {code: -1, msg: '服务器返回格式不正确'};
            }
            if (xhr.status < 200 || xhr.status >= 300 || response.code !== 0) {
                done(new Error(response.msg || ('请求失败（HTTP ' + xhr.status + '）')));
                return;
            }
            done(null, response.data);
        };
        xhr.onerror = function () {
            done(new Error('网络请求失败，请刷新页面后重试'));
        };
        xhr.send(payload
            ? 'payload=' + encodeURIComponent(JSON.stringify(payload))
            : null);
    }

    function idSet(csv) {
        var result = {};
        String(csv || '').split(/[,，\s]+/).forEach(function (value) {
            var id = parseInt(value, 10);
            if (id > 0) {
                result[id] = true;
            }
        });
        return result;
    }

    function selectedIds(map) {
        return Object.keys(map).map(function (key) {
            return parseInt(key, 10);
        }).filter(function (id) {
            return id > 0;
        });
    }

    function text(value) {
        return value === null || value === undefined ? '' : String(value);
    }

    function showToast(message, isError) {
        var toast = document.getElementById('ouToast');
        if (!toast) {
            return;
        }
        toast.textContent = message;
        toast.className = 'ou-toast visible' + (isError ? ' error' : '');
        window.clearTimeout(toastTimer);
        toastTimer = window.setTimeout(function () {
            toast.className = 'ou-toast';
        }, isError ? 5000 : 2600);
    }

    function cloneMap(map) {
        var result = {};
        Object.keys(map || {}).forEach(function (key) {
            result[key] = true;
        });
        return result;
    }

    function selectedCount(map) {
        return selectedIds(map || {}).length;
    }

    function escapeHtml(value) {
        return text(value)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#39;');
    }

    function directoryRows() {
        var seen = {};
        var rows = [];
        state.directories.forEach(function (item) {
            var id = parseInt(item.id, 10);
            if (!(id > 0) || seen[id]) {
                return;
            }
            seen[id] = true;
            var parentId = parseInt(item.parentId, 10);
            rows.push({
                id: id,
                parentId: parentId > 0 ? parentId : 0,
                nodeId: String(id),
                parentNode: parentId > 0 ? String(parentId) : '',
                name: text(item.name) || ('目录 #' + id)
            });
        });
        return rows;
    }

    function directoryMap(rows) {
        var result = {};
        rows.forEach(function (row) {
            result[row.nodeId] = row;
        });
        return result;
    }

    function selectedAndAncestorIds(rows, selected) {
        var map = directoryMap(rows);
        var keep = {};
        Object.keys(selected || {}).forEach(function (key) {
            var current = String(key);
            var guard = {};
            while (map[current] && !guard[current]) {
                guard[current] = true;
                keep[current] = true;
                current = map[current].parentNode;
            }
        });
        return keep;
    }

    function pickerRows() {
        var rows = directoryRows();
        if (!state.picker.selectedOnly) {
            return rows;
        }
        var keep = selectedAndAncestorIds(rows, state.picker.draft);
        return rows.filter(function (row) {
            return keep[row.nodeId] === true;
        });
    }

    function initialExpanded(rows, selected) {
        var map = directoryMap(rows);
        var expanded = {};
        rows.forEach(function (row) {
            if (!row.parentNode || !map[row.parentNode]) {
                expanded[row.nodeId] = true;
            }
        });
        var keep = selectedAndAncestorIds(rows, selected);
        Object.keys(keep).forEach(function (nodeId) {
            var row = map[nodeId];
            if (row && row.parentNode) {
                expanded[row.parentNode] = true;
            }
        });
        return expanded;
    }

    function pickerSelectableIds() {
        var keyword = state.picker.keyword.trim().toLowerCase();
        return pickerRows().filter(function (row) {
            if (!keyword) {
                return true;
            }
            return row.name.toLowerCase().indexOf(keyword) >= 0
                || String(row.id).indexOf(keyword) >= 0;
        }).map(function (row) {
            return row.id;
        });
    }

    function allCurrentSelected(ids) {
        return ids.length > 0 && ids.every(function (id) {
            return state.picker.draft[String(id)] === true;
        });
    }

    function renderScopeSummary(scope) {
        var summary = document.getElementById(scope === 'include'
            ? 'ouIncludeSummary' : 'ouExcludeSummary');
        var trigger = document.querySelector('[data-scope-trigger="' + scope + '"]');
        var count = selectedCount(state.selected[scope]);
        if (summary) {
            summary.textContent = count
                ? '已选择 ' + count + ' 个目录'
                : scope === 'include' ? '未选择，使用全部目录' : '未选择排除目录';
        }
        if (trigger) {
            trigger.classList.toggle('has-value', count > 0);
        }
    }

    function renderPickerTree() {
        var tree = document.getElementById('ouDirectoryTree');
        var rows = pickerRows();
        var keyword = state.picker.keyword.trim();
        var previousScroll = tree.scrollTop;
        if (!window.OaonesTree) {
            tree.innerHTML = '<div class="ou-picker-empty">目录树组件加载失败，请刷新页面后重试。</div>';
            return;
        }
        tree.innerHTML = window.OaonesTree.render(rows, {
            key: state.picker.scope,
            state: state.picker.expanded,
            keyword: keyword,
            idOf: function (row) { return row.nodeId; },
            parentOf: function (row) { return row.parentNode; },
            textOf: function (row) { return row.name + ' ' + row.id; },
            isGroup: function () { return false; },
            optionHtml: function (row, isHit) {
                var id = String(row.id);
                var selected = state.picker.draft[id] === true;
                return '<label class="ou-directory-option' + (selected ? ' active' : '')
                    + (isHit ? ' hit' : '') + '" title="目录 #' + escapeHtml(row.id) + '">'
                    + '<input type="checkbox" data-directory-id="' + id + '"'
                    + (selected ? ' checked' : '') + '>'
                    + '<span class="ou-directory-name">' + escapeHtml(row.name) + '</span>'
                    + '<small>#' + escapeHtml(row.id) + '</small></label>';
            },
            emptyClass: 'ou-picker-empty',
            emptyText: state.picker.selectedOnly
                ? '暂无已选择目录'
                : keyword ? '没有匹配的目录' : '未读取到知识文档目录'
        });
        tree.scrollTop = previousScroll;
    }

    function renderPicker() {
        var scope = state.picker.scope;
        if (!scope) {
            return;
        }
        var count = selectedCount(state.picker.draft);
        var ids = pickerSelectableIds();
        document.getElementById('ouPickerTitle').textContent = scope === 'include'
            ? '选择包含目录' : '选择排除目录';
        document.getElementById('ouPickerDescription').textContent = scope === 'include'
            ? '不选择时代表全部目录。'
            : '排除目录优先于包含目录。';
        document.getElementById('ouPickerSearch').value = state.picker.keyword;
        document.getElementById('ouPickerSelectedOnly').checked = state.picker.selectedOnly;
        document.getElementById('ouPickerToolbarCount').textContent = '已选 ' + count + ' 项';
        document.getElementById('ouPickerFootCount').textContent = '已选 ' + count + ' 项';
        var selectAll = document.getElementById('ouPickerAll');
        selectAll.textContent = allCurrentSelected(ids) ? '全不选' : '全选';
        selectAll.disabled = ids.length === 0;
        renderPickerTree();
    }

    function openPicker(scope, trigger) {
        state.picker = {
            scope: scope,
            draft: cloneMap(state.selected[scope]),
            keyword: '',
            selectedOnly: false,
            expanded: {},
            trigger: trigger
        };
        var rows = directoryRows();
        state.picker.expanded = initialExpanded(rows, state.picker.draft);
        var layer = document.getElementById('ouDirectoryPicker');
        layer.hidden = false;
        trigger.setAttribute('aria-expanded', 'true');
        document.body.classList.add('ou-picker-open');
        renderPicker();
        window.setTimeout(function () {
            document.getElementById('ouPickerSearch').focus();
        }, 0);
    }

    function closePicker() {
        var layer = document.getElementById('ouDirectoryPicker');
        if (!layer || layer.hidden) {
            return;
        }
        var trigger = state.picker.trigger;
        layer.hidden = true;
        document.body.classList.remove('ou-picker-open');
        if (trigger) {
            trigger.setAttribute('aria-expanded', 'false');
            trigger.focus();
        }
        state.picker.scope = '';
    }

    function confirmPicker() {
        var scope = state.picker.scope;
        if (!scope) {
            return;
        }
        state.selected[scope] = cloneMap(state.picker.draft);
        renderScopeSummary(scope);
        closePicker();
    }

    function formatInfo(extension) {
        var ext = text(extension).toLowerCase().replace(/^\./, '');
        return FORMAT_INFO[ext] || {
            label: ext.toUpperCase(),
            level: '支持',
            technology: '浏览器端文档解析渲染',
            tip: '该格式在浏览器内解析渲染。'
        };
    }

    function formatExtensionList(values) {
        return (values || []).map(function (value) {
            return '.' + text(value).toLowerCase().replace(/^\./, '');
        }).join('  ');
    }

    function textFormatState() {
        var enabled = 0;
        state.textFormats.forEach(function (ext) {
            if (state.formats[ext] === true) {
                enabled += 1;
            }
        });
        return {
            enabled: enabled,
            total: state.textFormats.length,
            all: state.textFormats.length > 0 && enabled === state.textFormats.length,
            some: enabled > 0 && enabled < state.textFormats.length
        };
    }

    function updateTextFormatGroup() {
        var card = document.querySelector(
            '.ou-format-option[data-format-group="text"]');
        if (!card) {
            return;
        }
        var summary = textFormatState();
        var input = card.querySelector('input[data-format-group="text"]');
        var stateLabel = card.querySelector('.ou-format-group-state');
        if (input) {
            input.checked = summary.all;
            input.indeterminate = summary.some;
        }
        if (stateLabel) {
            stateLabel.textContent = summary.some ? '部分启用' : '启用预览';
        }
        card.classList.toggle('enabled', summary.enabled > 0);
        card.classList.toggle('partial', summary.some);
    }

    function renderTextFormatGroup(target) {
        if (!state.textFormats.length) {
            return;
        }
        var summary = textFormatState();
        var card = document.createElement('article');
        card.className = 'ou-format-option ou-format-group is-text-group'
            + (summary.enabled > 0 ? ' enabled' : '')
            + (summary.some ? ' partial' : '');
        card.setAttribute('data-format-group', 'text');

        var main = document.createElement('div');
        main.className = 'ou-format-group-main';
        var head = document.createElement('div');
        head.className = 'ou-format-option-head';
        var name = document.createElement('div');
        name.className = 'ou-format-name';
        var title = document.createElement('strong');
        title.textContent = '文本文档';
        var level = document.createElement('span');
        level.textContent = state.textFormats.length + ' 种格式';
        name.appendChild(title);
        name.appendChild(level);
        head.appendChild(name);

        var extensions = document.createElement('p');
        extensions.className = 'ou-format-extensions';
        extensions.textContent = formatExtensionList(state.textFormats);
        extensions.title = extensions.textContent;

        var toggle = document.createElement('label');
        toggle.className = 'ou-format-toggle';
        var stateLabel = document.createElement('span');
        stateLabel.className = 'ou-format-group-state';
        stateLabel.textContent = summary.some ? '部分启用' : '启用预览';
        var switchControl = document.createElement('span');
        switchControl.className = 'ou-switch';
        var input = document.createElement('input');
        input.type = 'checkbox';
        input.setAttribute('data-format-group', 'text');
        input.setAttribute('aria-label', '启用文本文档即时预览');
        input.checked = summary.all;
        input.indeterminate = summary.some;
        var track = document.createElement('i');
        switchControl.appendChild(input);
        switchControl.appendChild(track);
        toggle.appendChild(stateLabel);
        toggle.appendChild(switchControl);

        main.appendChild(head);
        main.appendChild(extensions);
        card.appendChild(main);
        card.appendChild(toggle);
        target.appendChild(card);
    }

    function renderFormats(values) {
        var target = document.getElementById('ouFormatGrid');
        target.innerHTML = '';
        var textFormats = {};
        state.textFormats.forEach(function (extension) {
            textFormats[extension] = true;
        });
        renderTextFormatGroup(target);
        (values || []).filter(function (value) {
            var ext = text(value).toLowerCase().replace(/^\./, '');
            return textFormats[ext] !== true;
        }).forEach(function (value) {
            var ext = text(value).toLowerCase().replace(/^\./, '');
            var info = formatInfo(ext);
            var card = document.createElement('article');
            card.className = 'ou-format-option'
                + (state.formats[ext] === true ? ' enabled' : '');
            card.setAttribute('data-format', ext);

            var head = document.createElement('div');
            head.className = 'ou-format-option-head';
            var name = document.createElement('div');
            name.className = 'ou-format-name';
            var extension = document.createElement('strong');
            extension.textContent = '.' + info.label;
            var level = document.createElement('span');
            level.textContent = info.level;
            name.appendChild(extension);
            name.appendChild(level);
            head.appendChild(name);

            var technology = document.createElement('p');
            technology.className = 'ou-format-technology';
            technology.textContent = info.technology;

            var detail = document.createElement('p');
            detail.className = 'ou-format-detail';
            detail.textContent = info.tip;

            var toggle = document.createElement('label');
            toggle.className = 'ou-format-toggle';
            var toggleText = document.createElement('span');
            toggleText.textContent = '启用预览';
            var switchControl = document.createElement('span');
            switchControl.className = 'ou-switch';
            var input = document.createElement('input');
            input.type = 'checkbox';
            input.setAttribute('data-format', ext);
            input.setAttribute('aria-label', '启用 .' + info.label + ' 即时预览');
            input.checked = state.formats[ext] === true;
            var track = document.createElement('i');
            switchControl.appendChild(input);
            switchControl.appendChild(track);
            toggle.appendChild(toggleText);
            toggle.appendChild(switchControl);

            card.appendChild(head);
            card.appendChild(technology);
            card.appendChild(detail);
            card.appendChild(toggle);
            target.appendChild(card);
        });
    }

    function applyConfig(config) {
        state.config = config;
        var formats = config.convertibleFormats || [];
        var enabledFormats = config.enabledFormats;
        if (!Array.isArray(enabledFormats)) {
            enabledFormats = config.enabled ? formats : [];
        }
        state.textFormats = (Array.isArray(config.textFormats)
            ? config.textFormats
            : []).map(function (value) {
                return text(value).toLowerCase().replace(/^\./, '');
            }).filter(Boolean);
        state.formats = {};
        enabledFormats.forEach(function (value) {
            var ext = text(value).toLowerCase().replace(/^\./, '');
            if (ext) {
                state.formats[ext] = true;
            }
        });
        state.selected.include = idSet(config.includeDirs);
        state.selected.exclude = idSet(config.excludeDirs);
        renderScopeSummary('include');
        renderScopeSummary('exclude');
        renderFormats(formats);
    }

    function save() {
        if (state.saving) {
            return;
        }
        state.saving = true;
        document.querySelectorAll('.ou-save-button').forEach(function (button) {
            button.disabled = true;
            button.textContent = '保存中…';
        });
        var enabledFormats = Object.keys(state.formats).filter(function (ext) {
            return state.formats[ext] === true;
        });
        request('admin.save', {
            enabled: enabledFormats.length > 0,
            enabledFormats: enabledFormats,
            includeDirs: selectedIds(state.selected.include),
            excludeDirs: selectedIds(state.selected.exclude)
        }, function (error, config) {
            state.saving = false;
            document.querySelectorAll('.ou-save-button').forEach(function (button) {
                button.disabled = false;
                button.textContent = '保存设置';
            });
            if (error) {
                showToast(error.message, true);
                return;
            }
            applyConfig(config);
            showToast('设置已保存，预览范围和文件格式立即生效', false);
        });
    }

    function load() {
        request('admin.directories', null, function (dirError, directories) {
            if (dirError) {
                showToast(dirError.message, true);
                return;
            }
            state.directories = directories || [];
            request('admin.config', null, function (configError, config) {
                if (configError) {
                    showToast(configError.message, true);
                    return;
                }
                applyConfig(config);
            });
        });
    }

    document.getElementById('ouSave').addEventListener('click', save);
    document.getElementById('ouFormatGrid').addEventListener('change', function (event) {
        var groupInput = event.target.closest('input[data-format-group="text"]');
        if (groupInput) {
            state.textFormats.forEach(function (ext) {
                if (groupInput.checked) {
                    state.formats[ext] = true;
                } else {
                    delete state.formats[ext];
                }
            });
            updateTextFormatGroup();
            return;
        }
        var input = event.target.closest('input[data-format]');
        if (!input) {
            return;
        }
        var ext = input.getAttribute('data-format');
        if (input.checked) {
            state.formats[ext] = true;
        } else {
            delete state.formats[ext];
        }
        var card = input.closest('.ou-format-option');
        if (card) {
            card.classList.toggle('enabled', input.checked);
        }
    });
    document.querySelectorAll('[data-scope-trigger]').forEach(function (trigger) {
        trigger.addEventListener('click', function () {
            openPicker(trigger.getAttribute('data-scope-trigger'), trigger);
        });
    });
    document.querySelectorAll('[data-picker-close]').forEach(function (button) {
        button.addEventListener('click', closePicker);
    });
    document.getElementById('ouDirectoryPicker').addEventListener('click', function (event) {
        if (event.target === event.currentTarget) {
            closePicker();
        }
    });
    document.getElementById('ouDirectoryTree').addEventListener('click', function (event) {
        var caret = event.target.closest('.tree-caret');
        if (!caret || caret.classList.contains('leaf')) {
            return;
        }
        event.preventDefault();
        var nodeId = caret.getAttribute('data-node');
        state.picker.expanded[nodeId] = state.picker.expanded[nodeId] !== true;
        renderPickerTree();
    });
    document.getElementById('ouDirectoryTree').addEventListener('change', function (event) {
        var input = event.target.closest('input[data-directory-id]');
        if (!input) {
            return;
        }
        var id = input.getAttribute('data-directory-id');
        if (input.checked) {
            state.picker.draft[id] = true;
        } else {
            delete state.picker.draft[id];
        }
        renderPicker();
    });
    document.getElementById('ouPickerSearch').addEventListener('input', function (event) {
        state.picker.keyword = event.target.value;
        renderPicker();
    });
    document.getElementById('ouPickerSelectedOnly').addEventListener('change', function (event) {
        state.picker.selectedOnly = event.target.checked;
        if (state.picker.selectedOnly) {
            var expanded = initialExpanded(directoryRows(), state.picker.draft);
            Object.keys(expanded).forEach(function (nodeId) {
                state.picker.expanded[nodeId] = true;
            });
        }
        renderPicker();
    });
    document.getElementById('ouPickerAll').addEventListener('click', function () {
        var ids = pickerSelectableIds();
        var allOn = allCurrentSelected(ids);
        ids.forEach(function (id) {
            if (allOn) {
                delete state.picker.draft[String(id)];
            } else {
                state.picker.draft[String(id)] = true;
            }
        });
        renderPicker();
    });
    document.getElementById('ouPickerClear').addEventListener('click', function () {
        state.picker.draft = {};
        renderPicker();
    });
    document.getElementById('ouPickerConfirm').addEventListener('click', confirmPicker);
    document.addEventListener('keydown', function (event) {
        if (event.key === 'Escape') {
            closePicker();
        }
    });

    load();
})();
