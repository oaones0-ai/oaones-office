(function (global) {
    'use strict';

    async function render(options) {
        var container = options.container;
        var viewport = options.viewport;
        var shadowRoot = options.shadowRoot;
        var model = options.model;
        if (!global.docx || !global.docx.renderDocument) {
            throw new Error('Word 渲染组件未加载');
        }
        container.replaceChildren();
        var nodes = await global.docx.renderDocument(model, {
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
        var styleHost = shadowRoot.querySelector('[data-oa-docx-styles]');
        if (!styleHost) {
            styleHost = document.createElement('div');
            styleHost.setAttribute('data-oa-docx-styles', '');
            shadowRoot.appendChild(styleHost);
        }
        styleHost.replaceChildren();
        nodes.forEach(function (node) {
            if (node.nodeName === 'STYLE') {
                node.textContent = global.OfficeSecurity.cleanCss(node.textContent || '');
                styleHost.appendChild(node);
            } else {
                global.OfficeSecurity.sanitizeDom(node, shadowRoot);
                container.appendChild(node);
            }
        });

        var pages = Array.prototype.slice.call(container.querySelectorAll('section'));
        if (!pages.length) {
            pages = [container];
            pages[0].setAttribute('data-oa-generated-page', 'true');
        }
        pages.forEach(function (page, pageIndex) {
            var pageNumber = pageIndex + 1;
            page.setAttribute('data-oa-page', String(pageNumber));
            page.setAttribute('data-oa-anchor', 'page-' + pageNumber);
            page.id = 'oa-page-' + pageNumber;
            var blocks = page.querySelectorAll(
                'h1,h2,h3,h4,h5,h6,p,table,blockquote,pre,figure');
            Array.prototype.forEach.call(blocks, function (block, blockIndex) {
                var anchor = 'page-' + pageNumber + '/paragraph-' + (blockIndex + 1);
                block.setAttribute('data-oa-block-id',
                    'page-' + pageNumber + '-block-' + (blockIndex + 1));
                block.setAttribute('data-oa-anchor', anchor);
                block.setAttribute('data-oa-order', String(blockIndex + 1));
                block.setAttribute('data-oa-role',
                    /^H[1-6]$/.test(block.tagName) ? 'heading'
                        : block.tagName === 'TABLE' ? 'table' : 'paragraph');
            });
        });

        var baseWidth = 0;
        pages.forEach(function (page) {
            baseWidth = Math.max(baseWidth, page.offsetWidth || 0);
        });
        var currentScale = 1;
        var wrapper = container.querySelector('.docx-wrapper') || container;
        function setZoom(value, fitWidth) {
            var available = Math.max(320, viewport.clientWidth - 48);
            var next = fitWidth && baseWidth > 0
                ? Math.min(1.35, available / baseWidth) : value;
            currentScale = Math.max(0.4, Math.min(2.5, next || 1));
            wrapper.style.zoom = String(currentScale);
            return currentScale;
        }
        function sectionIndexAtViewport() {
            var viewportTop = viewport.getBoundingClientRect().top + 24;
            var result = 0;
            pages.forEach(function (page, index) {
                var rect = page.getBoundingClientRect();
                if (rect.top <= viewportTop + 80) {
                    result = index;
                }
            });
            return result;
        }
        viewport.addEventListener('scroll', function () {
            if (typeof options.onSectionChange === 'function' && !options.destroyed()) {
                options.onSectionChange(sectionIndexAtViewport());
            }
        }, {passive: true});

        return {
            sections: pages.map(function (page, index) {
                return {
                    id: 'page-' + (index + 1),
                    label: '第 ' + (index + 1) + ' 页',
                    kind: 'page',
                    index: index
                };
            }),
            setZoom: setZoom,
            fitWidth: function () {
                return setZoom(currentScale, true);
            },
            goTo: function (anchor) {
                var target = container.querySelector('[data-oa-anchor="' + cssEscape(anchor) + '"]');
                if (target) {
                    target.scrollIntoView({block: 'start', behavior: 'smooth'});
                    return true;
                }
                return false;
            },
            getCurrentIndex: sectionIndexAtViewport,
            getCurrentAnchor: function () {
                return 'page-' + (sectionIndexAtViewport() + 1);
            },
            destroy: function () {
                wrapper.style.zoom = '';
                container.replaceChildren();
                styleHost.replaceChildren();
            }
        };
    }

    function cssEscape(value) {
        if (global.CSS && typeof global.CSS.escape === 'function') {
            return global.CSS.escape(value);
        }
        return String(value).replace(/["\\]/g, '\\$&');
    }

    global.OfficeRenderers = global.OfficeRenderers || {};
    global.OfficeRenderers.docx = {
        render: render
    };
})(self);
