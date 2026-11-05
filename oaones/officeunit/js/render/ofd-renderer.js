(function (global) {
    'use strict';

    var scriptUrl = document.currentScript && document.currentScript.src
        ? document.currentScript.src
        : global.location.href;

    function rendererError(code, message) {
        var error = new Error(message);
        error.name = 'OfficePreviewError';
        error.code = code;
        return error;
    }

    function stylesheetUrl() {
        try {
            return new URL('../vendor/ofdrw/ofdrw-viewer-ui.css?v=0.1.7',
                scriptUrl).toString();
        } catch (ignore) {
            return '/oaones/officeunit/js/vendor/ofdrw/ofdrw-viewer-ui.css?v=0.1.7';
        }
    }

    function ensureStyles(shadowRoot) {
        if (!shadowRoot
                || shadowRoot.querySelector('link[data-oa-ofdrw-viewer-style]')) {
            return Promise.resolve();
        }
        var link = document.createElement('link');
        link.rel = 'stylesheet';
        link.href = stylesheetUrl();
        link.setAttribute('data-oa-ofdrw-viewer-style', '1');
        return new Promise(function (resolve) {
            link.onload = resolve;
            link.onerror = resolve;
            shadowRoot.appendChild(link);
        });
    }

    function normalizedWarning(warning) {
        if (!warning) {
            return '';
        }
        var code = String(warning.code || '').trim();
        var message = String(warning.message || '').trim();
        if (code === 'W_FONT_MISSING' || code === 'W_GLYPH_MISSING') {
            return message || 'OFD 字体或字形缺失';
        }
        return '';
    }

    function pageSections(pageCount) {
        var sections = [];
        for (var index = 1; index <= pageCount; index += 1) {
            sections.push({
                id: 'page-' + index,
                label: '第 ' + index + ' 页',
                kind: 'page',
                index: index - 1
            });
        }
        return sections;
    }

    function pageNumberFromAnchor(anchor) {
        var match = /^page-(\d+)$/.exec(String(anchor || ''));
        return match ? parseInt(match[1], 10) : 0;
    }

    async function render(options) {
        await ensureStyles(options.shadowRoot);
        var api = global.OFDRWViewerUI;
        if (!api || typeof api.OFDViewer !== 'function') {
            throw rendererError('OFFICE_PREVIEW_RENDER_FAILED',
                'OFD 预览组件加载失败');
        }
        if (!options.source) {
            throw rendererError('OFFICE_PREVIEW_FILE_FETCH_FAILED',
                'OFD 文件读取失败');
        }
        var limits = options.limits || {};
        if (typeof api.setResourceLimits === 'function') {
            var resourceLimits = {};
            var mappings = {
                maxPackageBytes: limits.maxInputBytes,
                maxEntries: limits.maxArchiveEntries,
                maxEntryBytes: limits.maxAssetBytes,
                maxTotalUncompressedBytes: limits.maxUncompressedBytes,
                maxCompressionRatio: limits.maxCompressionRatio,
                maxXmlBytes: limits.maxXmlBytes,
                maxFontBytes: limits.maxAssetBytes
            };
            Object.keys(mappings).forEach(function (key) {
                var value = Number(mappings[key]);
                if (Number.isFinite(value) && value > 0) {
                    resourceLimits[key] = value;
                }
            });
            api.setResourceLimits(resourceLimits);
        }

        var warnings = [];
        var viewer = new api.OFDViewer({
            container: options.container,
            source: options.source,
            initialScale: 'page-fit',
            theme: 'light',
            locale: 'zh-CN',
            updateHash: false,
            features: {
                attachments: false,
                export: false,
                print: false
            },
            onWarning: function (warning) {
                var message = normalizedWarning(warning);
                if (message && warnings.indexOf(message) < 0) {
                    warnings.push(message);
                }
            }
        });
        var timeoutMs = Math.max(1, Number(options.timeoutSeconds) || 30) * 1000;
        var timeout = null;
        try {
            await Promise.race([
                viewer.ready,
                new Promise(function (resolve, reject) {
                    timeout = global.setTimeout(function () {
                        reject(rendererError('OFFICE_PREVIEW_PARSE_TIMEOUT',
                            'OFD 文档解析超时，请稍后重试'));
                    }, timeoutMs);
                })
            ]);
            if (options.destroyed()) {
                viewer.destroy();
                throw rendererError('OFFICE_PREVIEW_RENDER_FAILED',
                    'OFD 预览已关闭');
            }
            var pageCount = viewer.getPageCount();
            if (pageCount <= 0) {
                throw rendererError('OFFICE_PREVIEW_PARSE_FAILED',
                    'OFD 文档不包含可预览页面');
            }
            if (pageCount > Math.max(1, Number(options.maxPageCount) || 2000)) {
                throw rendererError('OFFICE_PREVIEW_ARCHIVE_TOO_LARGE',
                    'OFD 文档页数超过预览上限');
            }
            var sections = pageSections(pageCount);
            viewer.on('pagechanging', function (event) {
                var page = Number(event && event.pageNumber);
                if (Number.isFinite(page) && page > 0) {
                    options.onSectionChange(page - 1);
                }
            });
            return {
                sections: sections,
                warnings: warnings,
                goTo: function (anchor) {
                    var page = pageNumberFromAnchor(anchor);
                    if (!page || page > pageCount) {
                        return false;
                    }
                    viewer.goToPage(page);
                    return true;
                },
                getCurrentAnchor: function () {
                    return 'page-' + viewer.getPageNumber();
                },
                fitWidth: function () {
                    viewer.setZoomMode('page-width');
                    return viewer.getScale();
                },
                fitPage: function () {
                    viewer.setZoomMode('page-fit');
                    return viewer.getScale();
                },
                setZoom: function (value) {
                    viewer.setScale(value);
                    return viewer.getScale();
                },
                destroy: function () {
                    viewer.destroy();
                }
            };
        } catch (error) {
            viewer.destroy();
            if (error && /^OFFICE_PREVIEW_[A-Z_]+$/.test(String(error.code || ''))) {
                throw error;
            }
            throw rendererError('OFFICE_PREVIEW_PARSE_FAILED',
                'OFD 文档解析失败，请下载后查看');
        } finally {
            if (timeout) {
                global.clearTimeout(timeout);
            }
        }
    }

    global.OfficeRenderers = global.OfficeRenderers || {};
    global.OfficeRenderers.ofd = {render: render};
})(self);
