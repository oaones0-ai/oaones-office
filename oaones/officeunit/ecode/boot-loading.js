(function (window, document) {
    'use strict';
    if (window.__oaonesOfficeUnitBootLoading) {
        return;
    }
    var fallbackTimer = null;
    var finishTimer = null;
    var completedSource = null;
    var positionedFrame = null;
    function root() {
        return document.getElementById('oaOfficeUnitBootLoading');
    }
    function create() {
        var existing = root();
        if (existing) {
            return existing;
        }
        var node = document.createElement('div');
        node.id = 'oaOfficeUnitBootLoading';
        node.className = 'oa-office-ecode-loading';
        node.setAttribute('role', 'status');
        node.setAttribute('aria-live', 'polite');
        node.style.display = 'none';
        var panel = document.createElement('div');
        panel.className = 'oa-office-ecode-loading-panel';
        var spinner = document.createElement('span');
        spinner.className = 'oa-office-ecode-loading-spinner';
        spinner.setAttribute('aria-hidden', 'true');
        var text = document.createElement('span');
        text.className = 'oa-office-ecode-loading-text';
        text.textContent = '文档加载中ing';
        panel.appendChild(spinner);
        panel.appendChild(text);
        node.appendChild(panel);
        (document.body || document.documentElement).appendChild(node);
        return node;
    }
    function position(target) {
        var node = root();
        if (!node) {
            return;
        }
        positionedFrame = target || null;
        node.style.display = 'none';
        if (!target) {
            return;
        }
        var rect = typeof target.getBoundingClientRect === 'function'
            ? target.getBoundingClientRect() : null;
        if (rect && rect.width > 0 && rect.height > 0) {
            node.style.left = rect.left + 'px';
            node.style.top = rect.top + 'px';
            node.style.width = rect.width + 'px';
            node.style.height = rect.height + 'px';
            node.style.display = 'flex';
        } else if (window.requestAnimationFrame) {
            window.requestAnimationFrame(function () {
                if (root() === node && positionedFrame === target) {
                    position(target);
                }
            });
        }
    }
    function clearTimers() {
        if (fallbackTimer) {
            window.clearTimeout(fallbackTimer);
            fallbackTimer = null;
        }
        if (finishTimer) {
            window.clearTimeout(finishTimer);
            finishTimer = null;
        }
    }
    function showAtTarget(node, target) {
        node.style.display = 'none';
        if (window.requestAnimationFrame) {
            window.requestAnimationFrame(function () {
                if (root() === node) {
                    position(target);
                }
            });
        } else {
            position(target);
        }
    }
    function show(message, target) {
        clearTimers();
        if (!target && message !== '文档已加载完毕') {
            return null;
        }
        var node = create();
        var text = node.querySelector('.oa-office-ecode-loading-text');
        text.textContent = message || '文档加载中ing';
        node.setAttribute('aria-busy', message === '文档已加载完毕' ? 'false' : 'true');
        showAtTarget(node, target);
        if (message === '文档已加载完毕') {
            finishTimer = window.setTimeout(hide, 1000);
        } else {
            fallbackTimer = window.setTimeout(hide, 120000);
        }
        return node;
    }
    function hide() {
        clearTimers();
        var node = root();
        if (node && node.parentNode) {
            node.parentNode.removeChild(node);
        }
    }
    window.__oaonesOfficeUnitBootLoading = {
        show: show,
        hide: hide,
        sync: syncFrame,
        position: function (frame) {
            position(frame);
        },
        root: root,
        message: function () {
            var node = root();
            var text = node && node.querySelector('.oa-office-ecode-loading-text');
            return text ? text.textContent : '';
        }
    };
    function frameSource(frame) {
        if (!frame || typeof frame.getAttribute !== 'function') {
            return '';
        }
        var source = frame.getAttribute('src');
        return source === null || source === undefined ? '' : String(source);
    }
    window.addEventListener('message', function (event) {
        if (!event || event.origin !== window.location.origin) {
            return;
        }
        var frame = document.getElementById('mainiframe');
        if (!frame || event.source !== frame.contentWindow || !event.data) {
            return;
        }
        if (event.data.type === 'oaones.officeunit.document-ready') {
            completedSource = frameSource(frame);
            show('文档已加载完毕', frame);
        } else if (event.data.type === 'oaones.officeunit.document-error') {
            completedSource = frameSource(frame);
            hide();
        }
    }, false);
    var observedFrame = null;
    var observedSource = '';
    function isOfficeFrameSource(source) {
        try {
            var url = new URL(source, window.location.origin);
            return url.origin === window.location.origin
                && (url.pathname === '/docs/docpreview/main.jsp'
                    || url.pathname === '/docs/view/main.jsp')
                && !!url.searchParams.get('docId')
                && !!url.searchParams.get('imageFileId')
                && url.searchParams.get('officeunit.original') !== '1';
        } catch (ignore) {
            return false;
        }
    }
    function isAttachmentTabActive() {
        var tab = document.querySelector(
            '[role="tab"][aria-selected="true"]');
        var text = tab && tab.textContent
            ? tab.textContent.replace(/\s/g, '') : '';
        return text.indexOf('文档附件') === 0;
    }
    function attachmentTarget(frame) {
        if (!isAttachmentTabActive()) {
            return null;
        }
        return document.querySelector('#wea_doc_detail_content .list-content')
            || (frame && frame.parentNode)
            || document.querySelector(
                '.wea-doc-detail-content-inner-tab-content');
    }
    function syncFrame() {
        var frame = document.getElementById('mainiframe');
        var target = attachmentTarget(frame);
        if (!target) {
            hide();
            return;
        }
        var source = frame
            ? frameSource(frame) : '';
        if (!root()) {
            if (!frame || (source !== completedSource
                    && (!source || source === 'about:blank'
                        || isOfficeFrameSource(source)))) {
                show('文档加载中ing', target);
            }
        } else {
            position(target);
        }
        if (!frame) {
            observedFrame = null;
            observedSource = '';
            return;
        }
        if (frame === observedFrame && source === observedSource) {
            return;
        }
        observedFrame = frame;
        observedSource = source;
        frame.addEventListener('load', function () {
            var loadedSource = frameSource(frame);
            if (loadedSource && loadedSource !== 'about:blank'
                    && !isOfficeFrameSource(loadedSource)) {
                hide();
            }
        }, false);
    }
    if (window.MutationObserver) {
        new MutationObserver(syncFrame).observe(document.body, {
            childList: true,
            subtree: true,
            attributes: true,
                attributeFilter: ['src', 'class', 'aria-selected']
        });
    }
    window.addEventListener('resize', function () {
        if (root() && positionedFrame) {
            position(positionedFrame);
        }
    }, false);
    window.addEventListener('scroll', function () {
        if (root() && positionedFrame) {
            position(positionedFrame);
        }
    }, true);
    if (window.requestAnimationFrame) {
        window.requestAnimationFrame(syncFrame);
    } else {
        window.setTimeout(syncFrame, 0);
    }
    fallbackTimer = window.setTimeout(hide, 120000);
})(window, document);
