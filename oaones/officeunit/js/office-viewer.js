(function (global) {
    'use strict';

    var sharedModels = new Map();
    var viewerSequence = 0;
    var LOADING_TEXT = '文档加载中ing';
    var LOADING_COMPLETE_TEXT = '文档已加载完毕';
    var LOADING_COMPLETE_DELAY = 1000;
    var DOCUMENT_READY_MESSAGE = 'oaones.officeunit.document-ready';
    var DOCUMENT_ERROR_MESSAGE = 'oaones.officeunit.document-error';
    var ERROR_MESSAGES = {
        OFFICE_PREVIEW_FORBIDDEN: '无权查看该文档',
        OFFICE_PREVIEW_FILE_NOT_FOUND: '未找到文档正文文件',
        OFFICE_PREVIEW_FILE_TOO_LARGE: '文件过大，暂不支持在线预览',
        OFFICE_PREVIEW_UNSUPPORTED: '暂不支持预览此文件格式',
        OFFICE_PREVIEW_FORMAT_DISABLED: '该文件格式未启用在线预览，请下载查看',
        OFFICE_PREVIEW_LEGACY_UNSUPPORTED: '旧版 Office 格式暂不支持在线预览，请下载查看',
        OFFICE_PREVIEW_ENCRYPTED: '文档已加密，请下载后输入密码打开',
        OFFICE_PREVIEW_MACRO_BLOCKED: '出于安全原因，宏格式文档已被禁止预览',
        OFFICE_PREVIEW_FILE_FETCH_FAILED: '文档加载失败，请稍后重试',
        OFFICE_PREVIEW_ARCHIVE_INVALID: '文档结构异常，无法安全预览',
        OFFICE_PREVIEW_ARCHIVE_TOO_LARGE: '文档解压后内容过大，暂不支持预览',
        OFFICE_PREVIEW_PARSE_TIMEOUT: '文档解析超时，请稍后重试',
        OFFICE_PREVIEW_PARSE_FAILED: '文档解析失败，请下载后查看',
        OFFICE_PREVIEW_RENDER_FAILED: '文档显示失败，请重新打开',
        OFFICE_PREVIEW_OUT_OF_MEMORY: '浏览器可用内存不足，请关闭其他预览后重试',
        OFFICE_PREVIEW_TICKET_EXPIRED: '预览会话已失效，请重新打开'
    };

    function OfficeViewer(options) {
        if (!options || !options.container) {
            throw new Error('OfficeViewer.create 缺少 container');
        }
        this.id = 'oa-office-viewer-' + (++viewerSequence);
        this.container = options.container;
        this.apiUrl = options.apiUrl || '/oaones/officeunit/page/api.jsp';
        this.workerUrl = options.workerUrl || deriveWorkerUrl(this.apiUrl);
        this.cssUrl = options.cssUrl
            || '/oaones/officeunit/css/office-viewer.css?v=2026110405';
        this.docId = intValue(options.docId);
        this.imageFileId = intValue(options.imageFileId);
        this.fileVersion = intValue(options.fileVersion);
        this.fileName = String(options.fileName || 'Office 文档');
        this.loadingKey = String(options.loadingKey
            || this.id);
        this.loadingText = String(options.loadingText || LOADING_TEXT);
        this.onStateChange = typeof options.onStateChange === 'function'
            ? options.onStateChange : null;
        this.onError = typeof options.onError === 'function' ? options.onError : null;
        this.destroyed = false;
        this.opening = null;
        this.prepared = null;
        this.renderer = null;
        this.modelLease = null;
        this.abortController = null;
        this.activeWorker = null;
        this.cancelActiveWorker = null;
        this.sectionIndex = 0;
        this.zoom = 1;
        this.ui = null;
        this.shadowRoot = null;
    }

    OfficeViewer.prototype.open = function () {
        var self = this;
        if (this.destroyed) {
            return Promise.reject(publicError('OFFICE_PREVIEW_RENDER_FAILED'));
        }
        if (this.opening) {
            return this.opening;
        }
        this.abortController = new AbortController();
        this.createShell();
        this.opening = this.openInternal().catch(function (error) {
            if (!self.destroyed) {
                var normalized = normalizeError(error);
                self.showError(normalized);
                if (self.onError) {
                    self.onError(normalized);
                }
            }
            if (self.modelLease) {
                self.modelLease.release();
                self.modelLease = null;
            }
            if (self.renderer) {
                try {
                    self.renderer.destroy();
                } catch (ignore) {
                    // Error cleanup is best effort.
                }
                self.renderer = null;
            }
            throw error;
        });
        return this.opening;
    };

    OfficeViewer.prototype.openInternal = async function () {
        this.emit('preparing', '正在校验文档权限');
        this.updateProgress(12, '正在校验登录与阅读权限');
        var prepared = await this.prepare();
        if (this.destroyed) {
            return;
        }
        this.prepared = prepared;
        this.updateFileMeta(prepared);
        this.ui.content.dataset.viewer = prepared.viewer;
        this.ui.shell.dataset.supportLevel = prepared.supportLevel;
        this.configureViewerChrome(prepared);
        if (prepared.viewer === 'ofd') {
            this.ui.shell.classList.add('is-ofd');
            this.ui.toolbar.hidden = true;
        }
        var legacy = isLegacyViewer(prepared.viewer);
        var entry = null;
        if (prepared.viewer === 'text') {
            this.emit('loading', '正在读取文本文档');
            this.updateProgress(42, '正在读取原始文本文件');
            entry = {
                buffer: await fetchFile(this, prepared)
            };
            if (entry.buffer.byteLength > prepared.limits.maxInputBytes) {
                throw publicError('OFFICE_PREVIEW_FILE_TOO_LARGE');
            }
            if (this.destroyed) {
                return;
            }
            this.updateProgress(80, '正在生成源码高亮视图');
        } else if (prepared.viewer === 'ofd') {
            this.emit('loading', '正在读取 OFD 文件');
            this.updateProgress(42, '正在读取原始 OFD 文件');
            entry = {
                buffer: await fetchFile(this, prepared)
            };
            if (entry.buffer.byteLength > prepared.limits.maxInputBytes) {
                throw publicError('OFFICE_PREVIEW_FILE_TOO_LARGE');
            }
            if (this.destroyed) {
                return;
            }
            this.updateProgress(82, '正在渲染 OFD 页面');
        } else if (!legacy) {
            this.emit('loading', '正在读取文档文件');
            this.updateProgress(30, '正在读取原始 Office 文件');
            this.modelLease = acquireModel(this, prepared);
            try {
                entry = await this.modelLease.promise;
            } catch (error) {
                this.modelLease.release();
                this.modelLease = null;
                throw error;
            }
            if (this.destroyed) {
                this.modelLease.release();
                this.modelLease = null;
                return;
            }
            this.updateProgress(82, '正在构建文档视图');
        } else {
            this.emit('loading', prepared.viewer === 'legacy-doc'
                ? '正在生成 Word HTML' : '正在生成幻灯片图片');
            this.updateProgress(42, prepared.viewer === 'legacy-doc'
                ? '正在将旧版 Word 转换为 HTML'
                : '正在将旧版 PowerPoint 转换为 PNG');
        }
        var rendererFactory = global.OfficeRenderers
            && global.OfficeRenderers[prepared.viewer];
        if (!rendererFactory) {
            throw publicError('OFFICE_PREVIEW_RENDER_FAILED');
        }
        var renderModel = prepared.viewer === 'docx'
            ? entry.docxModel : entry ? entry.parsed : null;
        var renderOptions = {
            model: renderModel,
            container: this.ui.content,
            viewport: this.ui.content,
            shadowRoot: this.shadowRoot,
            destroyed: function () {
                return this.destroyed;
            }.bind(this),
            onSectionChange: this.onSectionChange.bind(this)
        };
        if (legacy) {
            renderOptions.token = prepared.fileToken;
            renderOptions.fileUrl = prepared.fileUrl;
            renderOptions.legacyUrl = prepared.legacyUrl;
            renderOptions.extension = prepared.extension;
            renderOptions.signal = this.abortController
                ? this.abortController.signal : undefined;
        }
        if (prepared.viewer === 'ofd') {
            renderOptions.source = new Blob([entry.buffer], {
                type: 'application/ofd'
            });
            renderOptions.timeoutSeconds = prepared.limits.parseTimeoutSeconds;
            renderOptions.maxPageCount = prepared.limits.maxPageCount;
            renderOptions.limits = prepared.limits;
        } else if (prepared.viewer === 'text') {
            renderOptions.buffer = entry.buffer;
            renderOptions.extension = prepared.extension;
        }
        this.renderer = await rendererFactory.render(renderOptions);
        if (this.destroyed) {
            if (this.renderer) {
                this.renderer.destroy();
            }
            return;
        }
        if (!this.renderer) {
            throw publicError('OFFICE_PREVIEW_RENDER_FAILED');
        }
        if (prepared.viewer === 'docx'
                && this.renderer.sections.length > prepared.limits.maxPageCount) {
            throw publicError('OFFICE_PREVIEW_ARCHIVE_TOO_LARGE');
        }
        this.setSections(this.renderer.sections || []);
        this.renderWarnings(this.renderer.warnings
            || (entry && entry.parsed && entry.parsed.warnings) || []);
        this.setReady(prepared);
        this.updateSectionLabel();
        this.zoom = prepared.viewer === 'spreadsheet'
                && typeof this.renderer.setZoom === 'function'
            ? this.renderer.setZoom(1)
            : prepared.viewer === 'ofd'
                && typeof this.renderer.fitPage === 'function'
                ? this.renderer.fitPage()
                : this.renderer.fitWidth();
        this.updateZoomLabel();
        this.showLoadingComplete();
        this.notifyParent(DOCUMENT_READY_MESSAGE);
        this.emit('ready', '预览已就绪', {
            fileName: prepared.fileName,
            extension: prepared.extension,
            viewer: prepared.viewer,
            supportLevel: prepared.supportLevel
        });
    };

    function isLegacyViewer(viewer) {
        return viewer === 'legacy-doc' || viewer === 'legacy-ppt';
    }

    OfficeViewer.prototype.prepare = async function () {
        var body = new URLSearchParams();
        body.set('docId', String(this.docId));
        body.set('imageFileId', String(this.imageFileId));
        body.set('fileVersion', String(this.fileVersion));
        var response;
        try {
            response = await fetch(this.apiUrl
                + '?action=public.officeunit.preview.prepare', {
                method: 'POST',
                credentials: 'same-origin',
                cache: 'no-store',
                redirect: 'error',
                signal: this.abortController ? this.abortController.signal : undefined,
                headers: {
                    'Accept': 'application/json',
                    'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8',
                    'X-Requested-With': 'XMLHttpRequest'
                },
                body: body.toString()
            });
        } catch (ignore) {
            throw publicError('OFFICE_PREVIEW_FILE_FETCH_FAILED');
        }
        var payload = await response.json().catch(function () {
            return null;
        });
        if (!response.ok || !payload || payload.code !== 0 || !payload.data) {
            throw publicError(payload && payload.errorCode,
                payload && payload.msg);
        }
        var data = payload.data;
        if (!data.fileToken || !data.viewer || !data.extension
                || Number(data.fileSize) < 0) {
            throw publicError('OFFICE_PREVIEW_PARSE_FAILED');
        }
        return data;
    };

    OfficeViewer.prototype.createShell = function () {
        if (this.ui) {
            return;
        }
        this.showLoading();
        this.container.replaceChildren();
        this.shadowRoot = this.container.attachShadow({mode: 'open'});
        var style = document.createElement('link');
        style.rel = 'stylesheet';
        style.href = this.cssUrl;
        this.shadowRoot.appendChild(style);

        var shell = element('div', 'oa-office-viewer-shell');
        shell.setAttribute('data-viewer-id', this.id);
        var toolbar = element('header', 'oa-office-viewer-toolbar');
        var identity = element('div', 'oa-office-viewer-identity');
        var mark = element('span', 'oa-office-viewer-mark', 'DOC');
        var name = element('span', 'oa-office-viewer-name', this.fileName);
        var meta = element('span', 'oa-office-viewer-meta', 'Office');
        identity.appendChild(mark);
        identity.appendChild(name);
        identity.appendChild(meta);

        var navigation = element('div', 'oa-office-viewer-navigation');
        var previous = iconButton('‹', '上一节');
        var sectionLabel = element('span', 'oa-office-viewer-section-label', '正在准备');
        var next = iconButton('›', '下一节');
        navigation.appendChild(previous);
        navigation.appendChild(sectionLabel);
        navigation.appendChild(next);

        var actions = element('div', 'oa-office-viewer-actions');
        var zoomOut = iconButton('−', '缩小');
        var zoomControl = element('label', 'oa-office-viewer-zoom-control');
        zoomControl.title = '输入缩放百分比（40%–250%）';
        var zoomInput = element('input', 'oa-office-viewer-zoom');
        zoomInput.type = 'number';
        zoomInput.min = '40';
        zoomInput.max = '250';
        zoomInput.step = '10';
        zoomInput.value = '100';
        zoomInput.setAttribute('aria-label', '缩放百分比');
        var zoomSuffix = element('span', 'oa-office-viewer-zoom-suffix', '%');
        zoomControl.appendChild(zoomInput);
        zoomControl.appendChild(zoomSuffix);
        var zoomIn = iconButton('+', '放大');
        var fit = iconButton('↔', '适应宽度');
        var fullscreen = iconButton('⛶', '全屏');
        actions.appendChild(zoomOut);
        actions.appendChild(zoomControl);
        actions.appendChild(zoomIn);
        actions.appendChild(fit);
        actions.appendChild(fullscreen);
        toolbar.appendChild(identity);
        toolbar.appendChild(navigation);
        toolbar.appendChild(actions);

        var warning = element('div', 'oa-office-viewer-warning');
        warning.hidden = true;
        var progressLayer = element('div', 'oa-office-viewer-progress-layer');
        var progressCard = element('div', 'oa-office-viewer-progress-card');
        var spinner = element('span', 'oa-office-viewer-spinner');
        var progressTitle = element('strong', '', '正在打开文档');
        var progressText = element('p', '', '正在校验文档访问权限');
        var progressTrack = element('span', 'oa-office-viewer-progress-track');
        var progressBar = element('i', 'oa-office-viewer-progress-bar');
        progressTrack.appendChild(progressBar);
        progressCard.appendChild(spinner);
        progressCard.appendChild(progressTitle);
        progressCard.appendChild(progressText);
        progressCard.appendChild(progressTrack);
        progressLayer.appendChild(progressCard);

        var content = element('main', 'oa-office-viewer-content');
        content.setAttribute('aria-live', 'polite');
        var errorLayer = element('div', 'oa-office-viewer-error');
        errorLayer.hidden = true;
        var errorCard = element('div', 'oa-office-viewer-error-card');
        var errorTitle = element('strong', '', '无法打开文档预览');
        var errorText = element('p', '');
        var errorCode = element('small', '');
        errorCard.appendChild(errorTitle);
        errorCard.appendChild(errorText);
        errorCard.appendChild(errorCode);
        errorLayer.appendChild(errorCard);
        shell.appendChild(toolbar);
        shell.appendChild(warning);
        shell.appendChild(progressLayer);
        shell.appendChild(content);
        shell.appendChild(errorLayer);
        this.shadowRoot.appendChild(shell);

        this.ui = {
            shell: shell,
            toolbar: toolbar,
            mark: mark,
            name: name,
            meta: meta,
            navigation: navigation,
            previous: previous,
            next: next,
            sectionLabel: sectionLabel,
            zoomOut: zoomOut,
            zoomIn: zoomIn,
            zoomControl: zoomControl,
            zoomInput: zoomInput,
            fit: fit,
            fullscreen: fullscreen,
            warning: warning,
            progressLayer: progressLayer,
            progressTitle: progressTitle,
            progressText: progressText,
            progressBar: progressBar,
            content: content,
            errorLayer: errorLayer,
            errorText: errorText,
            errorCode: errorCode
        };
        previous.addEventListener('click', this.goPrevious.bind(this));
        next.addEventListener('click', this.goNext.bind(this));
        zoomOut.addEventListener('click', function () {
            this.setZoom(this.zoom - 0.1, false);
        }.bind(this));
        zoomIn.addEventListener('click', function () {
            this.setZoom(this.zoom + 0.1, false);
        }.bind(this));
        zoomInput.addEventListener('focus', function () {
            zoomInput.select();
        });
        zoomInput.addEventListener('change', function () {
            this.commitZoomInput();
        }.bind(this));
        zoomInput.addEventListener('keydown', function (event) {
            if (event.key === 'Enter') {
                event.preventDefault();
                this.commitZoomInput();
                zoomInput.blur();
            } else if (event.key === 'Escape') {
                event.preventDefault();
                this.updateZoomLabel();
                zoomInput.blur();
            }
        }.bind(this));
        fit.addEventListener('click', function () {
            this.setZoom(1, true);
        }.bind(this));
        fullscreen.addEventListener('click', function () {
            this.toggleFullscreen();
        }.bind(this));
        var originalFocusTimer = null;
        this.shadowRoot.addEventListener('click', function (event) {
            this.emitOriginalFocus(event);
        }.bind(this), true);
        this.shadowRoot.addEventListener('focusin', function (event) {
            if (originalFocusTimer) {
                global.clearTimeout(originalFocusTimer);
            }
            originalFocusTimer = global.setTimeout(function () {
                originalFocusTimer = null;
                if (!this.destroyed) {
                    this.emitOriginalFocus(event);
                }
            }.bind(this), 40);
        }.bind(this), true);
    };

    OfficeViewer.prototype.emitOriginalFocus = function (event) {
        if (!event || !this.container || this.destroyed) {
            return;
        }
        var path = typeof event.composedPath === 'function'
            ? event.composedPath() : [event.target];
        var target = null;
        for (var i = 0; i < path.length; i += 1) {
            if (path[i] && path[i].nodeType === 1) {
                target = path[i];
                break;
            }
        }
        if (!target) {
            return;
        }
        if (!this.ui || !this.ui.content || !this.ui.content.contains(target)) {
            return;
        }
        var cell = closestWithin(target, '[data-oa-cell]');
        var anchorNode = closestWithin(target, '[data-oa-anchor]');
        var textNode = closestWithin(target,
            '.oa-pptx-paragraph,.CodeMirror-line,p,li,td,th,[data-oa-anchor]');
        var anchor = '';
        var text = '';
        var sheetIndex = 0;
        var rowNumber = 0;
        var cellAddress = '';
        if (cell) {
            var activeSheet = this.shadowRoot.querySelector('.oa-sheet-tab.active');
            var sheetAnchor = activeSheet
                ? activeSheet.getAttribute('data-oa-sheet') : '';
            var sheetMatch = /^sheet-(\d+)$/i.exec(String(sheetAnchor || ''));
            sheetIndex = sheetMatch ? parseInt(sheetMatch[1], 10) : 0;
            cellAddress = String(cell.getAttribute('data-oa-cell') || '');
            var rowMatch = /^\D+(\d+)$/.exec(cellAddress);
            rowNumber = rowMatch ? parseInt(rowMatch[1], 10) : 0;
            anchor = sheetAnchor ? sheetAnchor + '/cell-' + cell.getAttribute('data-oa-cell') : '';
            text = String(cell.textContent || '').trim();
            target = cell;
        } else {
            anchor = anchorNode ? String(anchorNode.getAttribute('data-oa-anchor') || '') : '';
            text = String((textNode || target).textContent || '').trim();
            if (textNode) {
                target = textNode;
            }
        }
        if (!anchor && !text) {
            return;
        }
        var detail = {
            anchor: anchor,
            text: text.substring(0, 4000),
            sheetIndex: sheetIndex,
            rowNumber: rowNumber,
            cellAddress: cellAddress,
            viewer: this.prepared ? this.prepared.viewer : ''
        };
        var eventKey = detail.anchor + '\u0001' + detail.text;
        if (this.lastOriginalFocusKey === eventKey
                && Date.now() - (this.lastOriginalFocusAt || 0) < 200) {
            return;
        }
        this.lastOriginalFocusKey = eventKey;
        this.lastOriginalFocusAt = Date.now();
        this.container.dispatchEvent(new global.CustomEvent('oa-office-focus', {
            bubbles: true,
            composed: true,
            detail: detail
        }));
    };

    OfficeViewer.prototype.setReady = function (prepared) {
        this.ui.progressLayer.hidden = true;
        this.ui.content.dataset.viewer = prepared.viewer;
        this.ui.shell.dataset.supportLevel = prepared.supportLevel;
        this.configureViewerChrome(prepared);
        if (prepared.viewer === 'ofd') {
            this.ui.shell.classList.add('is-ofd');
        }
        this.ui.previous.disabled = false;
        this.ui.next.disabled = false;
        this.ui.zoomOut.disabled = false;
        this.ui.zoomIn.disabled = false;
        this.ui.fit.disabled = false;
        this.ui.fullscreen.disabled = !this.container.requestFullscreen;
    };

    OfficeViewer.prototype.configureViewerChrome = function (prepared) {
        if (!this.ui) {
            return;
        }
        var text = prepared.viewer === 'text';
        this.ui.shell.classList.toggle('is-text', text);
        this.ui.navigation.hidden = text;
        this.ui.zoomOut.hidden = text;
        this.ui.zoomControl.hidden = text;
        this.ui.zoomInput.hidden = text;
        this.ui.zoomIn.hidden = text;
        this.ui.fit.hidden = text;
    };

    OfficeViewer.prototype.updateFileMeta = function (prepared) {
        if (!this.ui) {
            return;
        }
        var extension = String(prepared.extension || '').toUpperCase();
        this.ui.mark.textContent = extension || 'DOC';
        this.ui.name.textContent = prepared.fileName || this.fileName;
        this.ui.meta.textContent = extension + ' · ' + formatBytes(prepared.fileSize)
            + (prepared.supportLevel === 'best_effort' ? ' · 兼容模式' : '');
    };

    OfficeViewer.prototype.updateProgress = function (percent, message) {
        if (!this.ui) {
            return;
        }
        this.ui.progressBar.style.width = Math.max(4, Math.min(100, percent)) + '%';
        this.ui.progressText.textContent = message || '正在处理文档';
    };

    OfficeViewer.prototype.setSections = function (sections) {
        this.sections = sections;
        this.sectionIndex = 0;
        this.updateSectionLabel();
    };

    OfficeViewer.prototype.updateSectionLabel = function () {
        if (!this.ui) {
            return;
        }
        var section = this.sections && this.sections[this.sectionIndex];
        this.ui.sectionLabel.textContent = section ? section.label
            : (this.sections && this.sections.length ? '1 / ' + this.sections.length : '1 节');
        this.ui.previous.disabled = !section || this.sectionIndex <= 0;
        this.ui.next.disabled = !section || this.sectionIndex >= this.sections.length - 1;
    };

    OfficeViewer.prototype.onSectionChange = function (index) {
        if (!this.sections || index < 0 || index >= this.sections.length
                || index === this.sectionIndex) {
            return;
        }
        this.sectionIndex = index;
        this.updateSectionLabel();
        this.emit('section-change', this.sections[index].label, this.sections[index]);
    };

    OfficeViewer.prototype.goPrevious = function () {
        return this.goToIndex(this.sectionIndex - 1);
    };

    OfficeViewer.prototype.goNext = function () {
        return this.goToIndex(this.sectionIndex + 1);
    };

    OfficeViewer.prototype.goToIndex = function (index) {
        if (!this.renderer || !this.sections || index < 0
                || index >= this.sections.length) {
            return false;
        }
        var section = this.sections[index];
        var result = this.renderer.goTo(section.id);
        if (result) {
            this.sectionIndex = index;
            this.updateSectionLabel();
        }
        return result;
    };

    OfficeViewer.prototype.goTo = function (anchor) {
        if (!this.renderer) {
            return false;
        }
        var result = this.renderer.goTo(anchor);
        if (result) {
            var target = String(anchor || '');
            for (var i = 0; i < this.sections.length; i += 1) {
                if (target.indexOf(this.sections[i].id) === 0) {
                    this.sectionIndex = i;
                    break;
                }
            }
            this.updateSectionLabel();
        }
        return result;
    };

    OfficeViewer.prototype.getCurrentAnchor = function () {
        return this.renderer ? this.renderer.getCurrentAnchor() : null;
    };

    OfficeViewer.prototype.getSectionList = function () {
        return (this.sections || []).map(function (section) {
            return {
                id: section.id,
                label: section.label,
                kind: section.kind,
                index: section.index
            };
        });
    };

    OfficeViewer.prototype.reveal = function (anchor, text) {
        var targetAnchor = String(anchor || '').trim();
        var targetText = String(text || '').trim();
        var target = null;
        var navigated = false;
        if (targetAnchor) {
            navigated = this.goTo(targetAnchor);
            var cellMatch = /\/cell-([A-Z]+\d+)$/i.exec(targetAnchor);
            if (cellMatch) {
                target = this.shadowRoot.querySelector(
                    '[data-oa-cell="' + cellMatch[1].toUpperCase() + '"]'
                );
            }
            if (!target) {
                try {
                    target = this.shadowRoot.querySelector(
                        '[data-oa-anchor="' + cssAttributeValue(targetAnchor) + '"]'
                    );
                } catch (ignore) {
                    target = null;
                }
            }
        }
        if (!target && targetText) {
            target = findVisibleText(this.shadowRoot, targetText);
        }
        if (!target && navigated) {
            return true;
        }
        if (!target) {
            return false;
        }
        Array.prototype.forEach.call(
            this.shadowRoot.querySelectorAll('.oa-office-focus-target'),
            function (item) {
                item.classList.remove('oa-office-focus-target');
            }
        );
        target.classList.add('oa-office-focus-target');
        target.scrollIntoView({block: 'center', behavior: 'smooth'});
        global.setTimeout(function () {
            if (target && !this.destroyed) {
                if (!target.hasAttribute('tabindex')) {
                    target.setAttribute('tabindex', '-1');
                }
                target.focus({preventScroll: true});
            }
        }.bind(this), 80);
        return true;
    };

    OfficeViewer.prototype.setZoom = function (value, fit) {
        if (!this.renderer || !this.ui) {
            return;
        }
        this.zoom = fit ? this.renderer.fitWidth()
            : this.renderer.setZoom(value);
        this.updateZoomLabel();
    };

    OfficeViewer.prototype.commitZoomInput = function () {
        if (!this.ui || !this.ui.zoomInput) {
            return;
        }
        var percent = Number(this.ui.zoomInput.value);
        if (!Number.isFinite(percent)) {
            this.updateZoomLabel();
            return;
        }
        this.setZoom(Math.max(40, Math.min(250, percent)) / 100, false);
    };

    OfficeViewer.prototype.updateZoomLabel = function () {
        if (!this.ui || !this.ui.zoomInput) {
            return;
        }
        var percent = String(Math.round(this.zoom * 100));
        this.ui.zoomInput.setAttribute('aria-valuetext', percent + '%');
        this.ui.zoomOut.disabled = this.zoom <= 0.4001;
        this.ui.zoomIn.disabled = this.zoom >= 2.4999;
        if (this.shadowRoot.activeElement !== this.ui.zoomInput) {
            this.ui.zoomInput.value = percent;
        }
    };

    OfficeViewer.prototype.toggleFullscreen = function () {
        if (!document.fullscreenElement) {
            if (this.container.requestFullscreen) {
                this.container.requestFullscreen().catch(function () {
                    // User agents may reject fullscreen without a direct gesture.
                });
            }
        } else if (document.exitFullscreen) {
            document.exitFullscreen();
        }
    };

    OfficeViewer.prototype.renderWarnings = function (warnings) {
        if (!this.ui) {
            return;
        }
        var unique = Array.from(new Set((warnings || []).filter(Boolean)));
        if (!unique.length) {
            this.ui.warning.hidden = true;
            return;
        }
        this.ui.warning.textContent = '部分内容已降级：' + unique.join(' ');
        this.ui.warning.hidden = false;
    };

    OfficeViewer.prototype.showError = function (error) {
        this.notifyParent(DOCUMENT_ERROR_MESSAGE, error);
        if (!this.ui) {
            return;
        }
        this.hideLoading();
        this.ui.progressLayer.hidden = true;
        this.ui.content.replaceChildren();
        this.ui.errorLayer.hidden = false;
        this.ui.errorText.textContent = error.message;
        this.ui.errorCode.textContent = error.code || '';
        this.ui.previous.disabled = true;
        this.ui.next.disabled = true;
    };

    OfficeViewer.prototype.emit = function (phase, message, data) {
        if (this.onStateChange) {
            try {
                this.onStateChange({
                    phase: phase,
                    message: message,
                    viewer: this,
                    data: data || null
                });
            } catch (ignore) {
                // Host callback failures do not break the preview.
            }
        }
    };

    OfficeViewer.prototype.showLoading = function (message) {
        var loading = global.OaonesLoading;
        if (!loading || this.destroyed
                || typeof loading.show !== 'function') {
            return null;
        }
        return loading.show(message || this.loadingText, {
            key: this.loadingKey
        });
    };

    OfficeViewer.prototype.showLoadingComplete = function () {
        var self = this;
        this.clearLoadingTimer();
        this.showLoading(LOADING_COMPLETE_TEXT);
        this.loadingHideTimer = global.setTimeout(function () {
            self.loadingHideTimer = null;
            self.hideLoading();
        }, LOADING_COMPLETE_DELAY);
    };

    OfficeViewer.prototype.clearLoadingTimer = function () {
        if (this.loadingHideTimer) {
            global.clearTimeout(this.loadingHideTimer);
            this.loadingHideTimer = null;
        }
    };

    OfficeViewer.prototype.hideLoading = function () {
        this.clearLoadingTimer();
        var loading = global.OaonesLoading;
        if (loading && typeof loading.hide === 'function') {
            loading.hide(this.loadingKey);
        }
    };

    OfficeViewer.prototype.notifyParent = function (type, error) {
        var parentWindow = global.parent;
        if (!parentWindow || parentWindow === global
                || typeof parentWindow.postMessage !== 'function') {
            return;
        }
        try {
            parentWindow.postMessage({
                type: type,
                code: error && error.code ? String(error.code) : '',
                docId: this.docId || 0,
                imageFileId: this.imageFileId || 0
            }, global.location.origin);
        } catch (ignore) {
            // Parent notification is best effort; document content remains usable.
        }
    };

    OfficeViewer.prototype.destroy = async function () {
        if (this.destroyed) {
            return;
        }
        this.destroyed = true;
        this.hideLoading();
        if (this.abortController) {
            this.abortController.abort();
        }
        if (this.cancelActiveWorker
                && (!this.modelLease || this.modelLease.entry.refs <= 1)) {
            this.cancelActiveWorker();
        }
        this.cancelActiveWorker = null;
        this.activeWorker = null;
        if (this.renderer) {
            try {
                this.renderer.destroy();
            } catch (ignore) {
                // Idempotent cleanup.
            }
            this.renderer = null;
        }
        if (this.modelLease) {
            this.modelLease.release();
            this.modelLease = null;
        }
        this.ui = null;
        this.sections = [];
        if (this.shadowRoot) {
            this.shadowRoot.replaceChildren();
            this.shadowRoot = null;
        }
        this.container.replaceChildren();
        this.emit('destroyed', '预览已关闭');
    };

    function acquireModel(viewer, prepared) {
        var key = [viewer.docId, viewer.imageFileId, viewer.fileVersion,
            prepared.extension, prepared.fileSize].join(':');
        var entry = sharedModels.get(key);
        if (!entry) {
            entry = {
                key: key,
                refs: 0,
                promise: null,
                objectUrls: [],
                disposed: false
            };
            sharedModels.set(key, entry);
            entry.promise = loadModel(viewer, prepared)
                .then(function (value) {
                    entry.parsed = value.parsed;
                    entry.docxModel = value.docxModel;
                    entry.objectUrls = value.objectUrls || [];
                    return entry;
                })
                .catch(function (error) {
                    sharedModels.delete(key);
                    entry.disposed = true;
                    throw error;
                });
        }
        entry.refs += 1;
        var released = false;
        return {
            entry: entry,
            promise: entry.promise,
            release: function () {
                if (released) {
                    return;
                }
                released = true;
                entry.refs -= 1;
                if (entry.refs <= 0) {
                    sharedModels.delete(key);
                    disposeModel(entry);
                }
            }
        };
    }

    async function loadModel(viewer, prepared) {
        viewer.updateProgress(42, '正在读取原始 Office 文件');
        var buffer = await fetchFile(viewer, prepared);
        if (buffer.byteLength > prepared.limits.maxInputBytes) {
            throw publicError('OFFICE_PREVIEW_FILE_TOO_LARGE');
        }
        viewer.updateProgress(58, '正在安全解包文档');
        var workerBuffer = prepared.viewer === 'docx' ? buffer.slice(0) : buffer;
        var parsed = await runWorker(viewer, workerBuffer, prepared);
        if (prepared.viewer !== 'docx') {
            return {parsed: parsed, objectUrls: []};
        }
        viewer.updateProgress(72, '正在解析 Word 分页结构');
        if (!global.docx || !global.docx.parseAsync) {
            throw publicError('OFFICE_PREVIEW_RENDER_FAILED');
        }
        var docxModel;
        try {
            docxModel = await global.docx.parseAsync(buffer, {
                inWrapper: true,
                breakPages: true,
                ignoreWidth: false,
                ignoreHeight: false,
                ignoreFonts: false,
                renderHeaders: true,
                renderFooters: true,
                renderFootnotes: true,
                renderEndnotes: true,
                renderAltChunks: false,
                renderComments: false,
                renderChanges: false,
                useBase64URL: false,
                experimental: false
            });
        } catch (error) {
            throw publicError('OFFICE_PREVIEW_PARSE_FAILED');
        }
        return {
            parsed: parsed,
            docxModel: docxModel,
            objectUrls: global.OfficeSecurity
                ? global.OfficeSecurity.collectBlobUrls(docxModel) : []
        };
    }

    async function fetchFile(viewer, prepared) {
        var url = new URL(prepared.fileUrl,
            new URL(viewer.apiUrl, global.location.href));
        url.searchParams.set('token', prepared.fileToken);
        var response;
        try {
            response = await fetch(url.toString(), {
                method: 'GET',
                credentials: 'same-origin',
                cache: 'no-store',
                redirect: 'error',
                signal: viewer.abortController ? viewer.abortController.signal : undefined,
                headers: {'Accept': 'application/octet-stream'}
            });
        } catch (ignore) {
            throw publicError('OFFICE_PREVIEW_FILE_FETCH_FAILED');
        }
        if (!response.ok) {
            var payload = null;
            if ((response.headers.get('content-type') || '').indexOf('application/json') >= 0) {
                payload = await response.json().catch(function () {
                    return null;
                });
            }
            throw publicError(payload && payload.errorCode,
                payload && payload.msg, 'OFFICE_PREVIEW_FILE_FETCH_FAILED');
        }
        var buffer = await response.arrayBuffer();
        if (buffer.byteLength !== Number(prepared.fileSize)) {
            throw publicError('OFFICE_PREVIEW_FILE_FETCH_FAILED');
        }
        return buffer;
    }

    function runWorker(viewer, buffer, prepared) {
        return new Promise(function (resolve, reject) {
            var worker;
            try {
                worker = new Worker(viewer.workerUrl);
            } catch (error) {
                reject(publicError('OFFICE_PREVIEW_PARSE_FAILED'));
                return;
            }
            var requestId = viewer.id + '-' + Date.now();
            var settled = false;
            var timeout = global.setTimeout(function () {
                settle(function () {
                    worker.terminate();
                    reject(publicError('OFFICE_PREVIEW_PARSE_TIMEOUT'));
                });
            }, Math.max(1, Number(prepared.limits.parseTimeoutSeconds) || 30) * 1000);
            function settle(action) {
                if (settled) {
                    return;
                }
                settled = true;
                global.clearTimeout(timeout);
                if (viewer.activeWorker === worker) {
                    viewer.activeWorker = null;
                    viewer.cancelActiveWorker = null;
                }
                action();
            }
            viewer.activeWorker = worker;
            viewer.cancelActiveWorker = function () {
                settle(function () {
                    worker.terminate();
                    reject(publicError('OFFICE_PREVIEW_PARSE_FAILED'));
                });
            };
            worker.onmessage = function (event) {
                var message = event.data || {};
                if (message.requestId !== requestId) {
                    return;
                }
                settle(function () {
                    worker.terminate();
                    if (message.ok) {
                        viewer.updateProgress(70, '文档结构检查完成');
                        resolve(message.result);
                    } else {
                        reject(publicError(message.errorCode, message.message,
                            'OFFICE_PREVIEW_PARSE_FAILED'));
                    }
                });
            };
            worker.onerror = function () {
                settle(function () {
                    worker.terminate();
                    reject(publicError('OFFICE_PREVIEW_PARSE_FAILED'));
                });
            };
            try {
                worker.postMessage({
                    type: 'parse',
                    requestId: requestId,
                    buffer: buffer,
                    extension: prepared.extension,
                    viewer: prepared.viewer,
                    limits: prepared.limits
                }, [buffer]);
            } catch (error) {
                settle(function () {
                    worker.terminate();
                    reject(publicError('OFFICE_PREVIEW_PARSE_FAILED'));
                });
            }
        });
    }

    function disposeModel(entry) {
        if (!entry || entry.disposed) {
            return;
        }
        entry.disposed = true;
        (entry.objectUrls || []).forEach(function (url) {
            URL.revokeObjectURL(url);
        });
        entry.objectUrls = [];
        entry.docxModel = null;
        entry.parsed = null;
    }

    function deriveWorkerUrl(apiUrl) {
        try {
            var url = new URL(apiUrl, global.location.href);
            var moduleIndex = url.pathname.indexOf('/oaones/officeunit/');
            if (moduleIndex >= 0) {
                url.pathname = url.pathname.substring(0, moduleIndex)
                    + '/oaones/officeunit/js/worker/office-parser.worker.js';
            } else if (/\/page\/api\.jsp$/.test(url.pathname)) {
                url.pathname = url.pathname.replace(/\/page\/api\.jsp$/,
                    '/js/worker/office-parser.worker.js');
            } else {
                url.pathname = '/js/worker/office-parser.worker.js';
            }
            url.search = '?v=2026110402';
            return url.toString();
        } catch (ignore) {
            return '/oaones/officeunit/js/worker/office-parser.worker.js';
        }
    }

    function publicError(code, message, fallbackCode) {
        var normalizedCode = /^OFFICE_PREVIEW_[A-Z_]+$/.test(String(code || ''))
            ? code : (fallbackCode || 'OFFICE_PREVIEW_PARSE_FAILED');
        var error = new Error(message || ERROR_MESSAGES[normalizedCode]
            || ERROR_MESSAGES.OFFICE_PREVIEW_PARSE_FAILED);
        error.name = 'OfficePreviewError';
        error.code = normalizedCode;
        return error;
    }

    function normalizeError(error) {
        if (error && error.code && ERROR_MESSAGES[error.code]) {
            return error;
        }
        if (error && error.name === 'QuotaExceededError') {
            return publicError('OFFICE_PREVIEW_OUT_OF_MEMORY');
        }
        return publicError('OFFICE_PREVIEW_RENDER_FAILED');
    }

    function intValue(value) {
        var number = parseInt(value, 10);
        return isFinite(number) && number > 0 ? number : 0;
    }

    function formatBytes(value) {
        var bytes = Number(value) || 0;
        if (bytes < 1024) {
            return bytes + ' B';
        }
        if (bytes < 1024 * 1024) {
            return (bytes / 1024).toFixed(1) + ' KB';
        }
        return (bytes / 1024 / 1024).toFixed(1) + ' MB';
    }

    function closestWithin(element, selector) {
        return element && element.closest ? element.closest(selector) : null;
    }

    function cssAttributeValue(value) {
        return String(value || '').replace(/\\/g, '\\\\').replace(/"/g, '\\"');
    }

    function normalizedText(value) {
        return String(value || '').replace(/[\s\u00a0\u3000]+/g, ' ').trim().toLowerCase();
    }

    function findVisibleText(root, value) {
        var wanted = normalizedText(value);
        if (!root || !wanted) {
            return null;
        }
        var candidates = root.querySelectorAll(
            '[data-oa-cell],.oa-pptx-paragraph,.CodeMirror-line,p,li,td,th,'
                + '[data-oa-anchor]'
        );
        var partial = null;
        for (var i = 0; i < candidates.length; i += 1) {
            var current = normalizedText(candidates[i].textContent);
            if (!current) {
                continue;
            }
            if (current === wanted) {
                return candidates[i];
            }
            if (!partial && (current.indexOf(wanted) >= 0 || wanted.indexOf(current) >= 0)) {
                partial = candidates[i];
            }
        }
        var shadows = root.querySelectorAll('*');
        for (var j = 0; j < shadows.length; j += 1) {
            if (shadows[j].shadowRoot) {
                var nested = findVisibleText(shadows[j].shadowRoot, value);
                if (nested) {
                    return nested;
                }
            }
        }
        return partial;
    }

    function iconButton(text, title) {
        var button = element('button', 'oa-office-viewer-icon', text);
        button.type = 'button';
        button.title = title;
        button.setAttribute('aria-label', title);
        button.disabled = true;
        return button;
    }

    function element(tag, className, text) {
        var node = document.createElement(tag);
        if (className) {
            node.className = className;
        }
        if (text !== undefined) {
            node.textContent = text;
        }
        return node;
    }

    function create(options) {
        return new OfficeViewer(options);
    }

    global.OfficeViewer = {
        create: create,
        version: '2026.11.04'
    };

    function bootstrap() {
        var roots = document.querySelectorAll('[data-oa-office-viewer-auto]');
        Array.prototype.forEach.call(roots, function (root) {
            if (root.getAttribute('data-booted') === '1') {
                return;
            }
            root.setAttribute('data-booted', '1');
            var viewer = create({
                container: root,
                apiUrl: root.getAttribute('data-api'),
                docId: root.getAttribute('data-doc-id'),
                imageFileId: root.getAttribute('data-image-file-id'),
                fileVersion: root.getAttribute('data-file-version'),
                fileName: root.getAttribute('data-file-name'),
                loadingKey: root.getAttribute('data-loading-key'),
                loadingText: root.getAttribute('data-loading-text')
            });
            root.__oaOfficeViewer = viewer;
            viewer.open().catch(function () {
                // The viewer renders its own controlled error state.
            });
        });
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', bootstrap, {once: true});
    } else {
        bootstrap();
    }
})(self);
