(function (global) {
    'use strict';

    function renderDoc(options) {
        var container = options.container;
        var viewport = options.viewport;
        var shadowRoot = options.shadowRoot;
        var url = legacyUrl(options, 'html');
        return fetchText(url, options.signal).then(function (html) {
            if (options.destroyed()) {
                return null;
            }
            var parsed = new DOMParser().parseFromString(html, 'text/html');
            if (!parsed.body) {
                throw previewError('OFFICE_PREVIEW_RENDER_FAILED');
            }

            var styleHost = shadowRoot.querySelector('[data-oa-legacy-styles]');
            if (!styleHost) {
                styleHost = document.createElement('div');
                styleHost.setAttribute('data-oa-legacy-styles', '');
                shadowRoot.appendChild(styleHost);
            }
            styleHost.replaceChildren();
            Array.prototype.forEach.call(parsed.querySelectorAll('style'),
                function (style) {
                    var clean = document.createElement('style');
                    clean.textContent = global.OfficeSecurity
                        ? global.OfficeSecurity.cleanCss(style.textContent || '')
                        : '';
                    styleHost.appendChild(clean);
                });

            var page = element('article', 'oa-legacy-doc');
            if (parsed.body.className) {
                page.classList.add.apply(page.classList,
                    parsed.body.className.split(/\s+/).filter(Boolean));
            }
            if (parsed.body.getAttribute('style')) {
                page.setAttribute('style', global.OfficeSecurity
                    ? global.OfficeSecurity.cleanCss(parsed.body.getAttribute('style'))
                    : '');
            }
            global.OfficeSecurity.sanitizeDom(parsed.body, shadowRoot);
            while (parsed.body.firstChild) {
                page.appendChild(parsed.body.firstChild);
            }
            page.setAttribute('data-oa-page', '1');
            page.setAttribute('data-oa-anchor', 'document');
            container.replaceChildren(page);

            var scale = 1;
            function setZoom(value, fitWidth) {
                var available = Math.max(320, viewport.clientWidth - 48);
                var width = Math.max(1, page.scrollWidth || page.offsetWidth || 816);
                var next = fitWidth ? available / width : value;
                scale = Math.max(0.4, Math.min(2.5, next || 1));
                page.style.zoom = String(scale);
                return scale;
            }
            function activeIndex() {
                return 0;
            }
            viewport.addEventListener('scroll', function () {
                if (typeof options.onSectionChange === 'function'
                        && !options.destroyed()) {
                    options.onSectionChange(0);
                }
            }, {passive: true});
            return {
                sections: [{
                    id: 'document',
                    label: '文档',
                    kind: 'document',
                    index: 0
                }],
                warnings: [],
                setZoom: setZoom,
                fitWidth: function () {
                    return setZoom(scale, true);
                },
                goTo: function () {
                    page.scrollIntoView({block: 'start'});
                    return true;
                },
                getCurrentIndex: activeIndex,
                getCurrentAnchor: function () {
                    return 'document';
                },
                destroy: function () {
                    container.replaceChildren();
                    styleHost.replaceChildren();
                }
            };
        });
    }

    function renderPpt(options) {
        var container = options.container;
        var viewport = options.viewport;
        var manifestUrl = legacyUrl(options, 'manifest');
        return fetchJson(manifestUrl, options.signal).then(function (payload) {
            if (!payload || payload.code !== 0 || !payload.data) {
                throw previewError(payload && payload.errorCode, payload && payload.msg);
            }
            if (options.destroyed()) {
                return null;
            }
            var data = payload.data;
            var count = Number(data.slideCount);
            var width = Number(data.width);
            var height = Number(data.height);
            if (!isFinite(count) || count <= 0 || !isFinite(width)
                    || width <= 0 || !isFinite(height) || height <= 0) {
                throw previewError('OFFICE_PREVIEW_PARSE_FAILED');
            }

            var root = element('div', 'oa-legacy-ppt');
            var scroller = element('div', 'oa-legacy-ppt-scroll');
            root.appendChild(scroller);
            container.replaceChildren(root);
            var slides = [];
            for (var index = 0; index < count; index += 1) {
                var section = element('section', 'oa-legacy-slide');
                section.setAttribute('data-oa-page', String(index + 1));
                section.setAttribute('data-oa-slide', String(index + 1));
                section.setAttribute('data-oa-anchor', 'slide-' + (index + 1));
                section.id = 'oa-legacy-slide-' + (index + 1);
                var stage = element('div', 'oa-legacy-slide-stage');
                stage.style.width = width + 'px';
                stage.style.aspectRatio = width + ' / ' + height;
                var image = document.createElement('img');
                image.alt = '第 ' + (index + 1) + ' 页';
                image.loading = 'lazy';
                image.decoding = 'async';
                image.setAttribute('referrerpolicy', 'no-referrer');
                image.src = legacyUrl(options, 'slide', index);
                image.addEventListener('error', function () {
                    this.classList.add('failed');
                    this.alt = '幻灯片加载失败';
                });
                stage.appendChild(image);
                section.appendChild(stage);
                scroller.appendChild(section);
                slides.push(section);
            }

            var scale = 1;
            function setZoom(value, fitWidth) {
                var available = Math.max(320, viewport.clientWidth - 48);
                var next = fitWidth ? available / width : value;
                scale = Math.max(0.25, Math.min(3, next || 1));
                scroller.style.zoom = String(scale);
                return scale;
            }
            function currentIndex() {
                var target = viewport.getBoundingClientRect().top + 24;
                var result = 0;
                slides.forEach(function (slide, index) {
                    if (slide.getBoundingClientRect().top <= target + 80) {
                        result = index;
                    }
                });
                return result;
            }
            viewport.addEventListener('scroll', function () {
                if (typeof options.onSectionChange === 'function'
                        && !options.destroyed()) {
                    options.onSectionChange(currentIndex());
                }
            }, {passive: true});

            return {
                sections: slides.map(function (slide, index) {
                    return {
                        id: 'slide-' + (index + 1),
                        label: '第 ' + (index + 1) + ' 页',
                        kind: 'page',
                        index: index
                    };
                }),
                warnings: [],
                setZoom: setZoom,
                fitWidth: function () {
                    return setZoom(scale, true);
                },
                goTo: function (anchor) {
                    var index = sectionIndex(anchor);
                    if (index < 0 || !slides[index]) {
                        return false;
                    }
                    slides[index].scrollIntoView({block: 'start', behavior: 'smooth'});
                    return true;
                },
                getCurrentIndex: currentIndex,
                getCurrentAnchor: function () {
                    return 'slide-' + (currentIndex() + 1);
                },
                destroy: function () {
                    Array.prototype.forEach.call(
                        container.querySelectorAll('img'), function (image) {
                            image.removeAttribute('src');
                        });
                    container.replaceChildren();
                }
            };
        });
    }

    function legacyUrl(options, part, index) {
        var base = options.legacyUrl;
        if (!base) {
            base = String(options.fileUrl || '/oaones/officeunit/preview/file.jsp')
                .replace(/\/file\.jsp(?:[?#].*)?$/,
                    '/legacy.jsp');
        }
        var url = new URL(base, global.location.href);
        if (url.origin !== global.location.origin) {
            throw previewError('OFFICE_PREVIEW_FILE_FETCH_FAILED');
        }
        url.searchParams.set('token', options.token || '');
        url.searchParams.set('part', part);
        if (index !== undefined) {
            url.searchParams.set('index', String(index));
        }
        return url.toString();
    }

    function fetchText(url, signal) {
        return fetch(url, {
            method: 'GET',
            credentials: 'same-origin',
            cache: 'no-store',
            redirect: 'error',
            signal: signal,
            headers: {'Accept': 'text/html'}
        }).catch(function () {
            throw previewError('OFFICE_PREVIEW_FILE_FETCH_FAILED');
        }).then(function (response) {
            if (!response.ok) {
                return response.text().then(function (text) {
                    throw parsedError(text);
                });
            }
            return response.text();
        });
    }

    function fetchJson(url, signal) {
        return fetch(url, {
            method: 'GET',
            credentials: 'same-origin',
            cache: 'no-store',
            redirect: 'error',
            signal: signal,
            headers: {'Accept': 'application/json'}
        }).catch(function () {
            throw previewError('OFFICE_PREVIEW_FILE_FETCH_FAILED');
        }).then(function (response) {
            return response.json().catch(function () {
                return null;
            }).then(function (payload) {
                if (!response.ok || !payload || payload.code !== 0) {
                    throw previewError(payload && payload.errorCode,
                        payload && payload.msg);
                }
                return payload;
            });
        });
    }

    function parsedError(text) {
        try {
            var payload = JSON.parse(String(text || ''));
            return previewError(payload.errorCode, payload.msg);
        } catch (ignore) {
            return previewError('OFFICE_PREVIEW_FILE_FETCH_FAILED');
        }
    }

    function sectionIndex(anchor) {
        var match = /^slide-(\d+)$/.exec(String(anchor || ''));
        return match ? Number(match[1]) - 1 : -1;
    }

    function previewError(code, message) {
        var error = new Error(message || '旧版 Office 文档显示失败，请稍后重试');
        error.name = 'OfficePreviewError';
        error.code = code || 'OFFICE_PREVIEW_RENDER_FAILED';
        return error;
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

    global.OfficeRenderers = global.OfficeRenderers || {};
    global.OfficeRenderers['legacy-doc'] = {render: renderDoc};
    global.OfficeRenderers['legacy-ppt'] = {render: renderPpt};
})(self);
