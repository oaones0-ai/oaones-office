/**
 * OAones OfficeUnit E-code adapter.
 * Rewrites E9 getAccFileConvert before the SPA sets sysRemind.jsp.
 */
(function (window, document) {
    'use strict';

    if (window.__oaonesOfficeUnitEcodeInstalled) {
        return;
    }
    window.__oaonesOfficeUnitEcodeInstalled = true;

    var API_PATH = '/api/doc/acc/getAccFileConvert';
    var PREVIEW_PATH = '/docs/docpreview/main.jsp';
    var LOADING_TEXT = '文档加载中ing';
    var LOADING_COMPLETE_TEXT = '文档已加载完毕';
    var LOADING_COMPLETE_DELAY = 1000;
    var DOCUMENT_READY_MESSAGE = 'oaones.officeunit.document-ready';
    var DOCUMENT_ERROR_MESSAGE = 'oaones.officeunit.document-error';
    var DIRECT_LOADING_STYLE = [
        '.oa-office-ecode-loading{position:fixed;left:0;top:0;width:100%;height:100%;',
        'z-index:2147483000;display:flex;align-items:center;justify-content:center;',
        'background:rgba(255,255,255,.9);box-sizing:border-box;padding:24px;',
        'font-family:-apple-system,BlinkMacSystemFont,"Segoe UI","Microsoft YaHei",sans-serif;}',
        '.oa-office-ecode-loading-panel{display:flex;align-items:center;gap:11px;',
        'padding:15px 20px;border:1px solid rgba(0,85,204,.12);border-radius:10px;',
        'background:#fff;box-shadow:0 8px 28px rgba(24,39,75,.14);color:#24324a;',
        'font-size:14px;line-height:20px;}',
        '.oa-office-ecode-loading-spinner{width:20px;height:20px;border:2px solid #dce8fb;',
        'border-top-color:#1a73e8;border-radius:50%;animation:oa-office-ecode-spin .75s linear infinite;}',
        '@keyframes oa-office-ecode-spin{to{transform:rotate(360deg)}}',
        '@media (prefers-reduced-motion:reduce){',
        '.oa-office-ecode-loading-spinner{animation:none;}}'
    ].join('');
    var AUTH_KEYS = [
        'authStr', 'authSignatureStr', 'isrequest', 'requestid', 'desrequestid',
        'f_weaver_belongto_userid', 'f_weaver_belongto_usertype',
        'fromFlowDoc', 'meetingid', 'workplanid'
    ];
    var CONVERTIBLE_EXTENSIONS = {
        doc: true,
        docx: true,
        dotx: true,
        xlsx: true,
        xls: true,
        xlsb: true,
        csv: true,
        tsv: true,
        ppt: true,
        pptx: true,
        potx: true,
        ppsx: true,
        ods: true,
        ofd: true,
        txt: true,
        log: true,
        text: true,
        py: true,
        pyw: true,
        js: true,
        jsx: true,
        ts: true,
        tsx: true,
        java: true,
        c: true,
        h: true,
        cpp: true,
        cc: true,
        cxx: true,
        hpp: true,
        cs: true,
        kt: true,
        scala: true,
        m: true,
        mm: true,
        go: true,
        rs: true,
        php: true,
        rb: true,
        lua: true,
        r: true,
        sql: true,
        json: true,
        xml: true,
        html: true,
        htm: true,
        css: true,
        scss: true,
        less: true,
        yaml: true,
        yml: true,
        md: true,
        markdown: true,
        sh: true,
        bash: true,
        zsh: true,
        bat: true,
        cmd: true,
        ps1: true,
        psd1: true,
        psm1: true,
        properties: true,
        ini: true,
        in: true,
        cfg: true,
        toml: true,
        swift: true,
        vue: true,
        diff: true,
        patch: true,
        vb: true,
        vbs: true,
        groovy: true,
        gradle: true,
        dart: true,
        coffee: true,
        pl: true,
        pm: true,
        conf: true
    };
    var attemptedFrames = Object.create(null);
    var observer = null;
    var scheduled = false;
    var loadingFrame = null;
    var pendingDocumentLoading = false;
    var directLoading = null;
    var directLoadingFrame = null;
    var directLoadingHideTimer = null;
    var directLoadingPositionBound = false;

    function contentPath() {
        return window.ecologyContentPath || '';
    }

    function isDocumentDetailPage() {
        return /\/spa\/document\/index\.jsp$/i.test(window.location.pathname);
    }

    function queryMap(value) {
        var result = Object.create(null);
        if (!value) {
            return result;
        }
        var source = String(value);
        var index = source.indexOf('?');
        if (index >= 0) {
            source = source.substring(index + 1);
        }
        var hash = source.indexOf('#');
        if (hash >= 0) {
            source = source.substring(0, hash);
        }
        source.split('&').forEach(function (part) {
            if (!part) {
                return;
            }
            var equal = part.indexOf('=');
            var key = equal < 0 ? part : part.substring(0, equal);
            var item = equal < 0 ? '' : part.substring(equal + 1);
            try {
                key = decodeURIComponent(key.replace(/\+/g, ' '));
                item = decodeURIComponent(item.replace(/\+/g, ' '));
            } catch (ignore) {
                return;
            }
            result[String(key).toLowerCase()] = item;
        });
        return result;
    }

    function mergeMap(target) {
        for (var i = 1; i < arguments.length; i += 1) {
            var source = arguments[i] || {};
            Object.keys(source).forEach(function (key) {
                var normalized = String(key).toLowerCase();
                if (target[normalized] === undefined || target[normalized] === '') {
                    target[normalized] = source[key];
                }
            });
        }
        return target;
    }

    function pick(map, names) {
        for (var i = 0; i < names.length; i += 1) {
            var value = map[String(names[i]).toLowerCase()];
            if (value !== undefined && value !== null && String(value) !== '') {
                return String(value);
            }
        }
        return '';
    }

    function currentParams() {
        return queryMap(window.location.search);
    }

    function likelyDocumentPreviewRequested() {
        var params = currentParams();
        return !!pick(params, ['imagefileid', 'imageFileId', 'fileid']);
    }

    function element(tagName, className, text) {
        var node = document.createElement(tagName);
        if (className) {
            node.className = className;
        }
        if (text !== undefined) {
            node.textContent = text;
        }
        return node;
    }

    function bootLoading() {
        var loading = window.__oaonesOfficeUnitBootLoading;
        return loading && typeof loading.show === 'function'
            && typeof loading.hide === 'function' ? loading : null;
    }

    function placeDirectLoading(frame) {
        if (!directLoading) {
            return;
        }
        directLoadingFrame = frame || null;
        var boot = bootLoading();
        if (boot && directLoading.boot) {
            boot.position(frame);
            return;
        }
        var rect = frame && typeof frame.getBoundingClientRect === 'function'
            ? frame.getBoundingClientRect() : null;
        var rootStyle = directLoading.root.style;
        if (rect && rect.width > 0 && rect.height > 0) {
            rootStyle.left = rect.left + 'px';
            rootStyle.top = rect.top + 'px';
            rootStyle.width = rect.width + 'px';
            rootStyle.height = rect.height + 'px';
        } else {
            rootStyle.left = '0px';
            rootStyle.top = '0px';
            rootStyle.width = '100%';
            rootStyle.height = '100%';
        }
    }

    function ensureDirectLoading() {
        if (directLoading) {
            return directLoading;
        }
        if (bootLoading()) {
            directLoading = {boot: true};
            return directLoading;
        }
        var host = document.body || document.documentElement;
        if (!host) {
            return null;
        }
        if (!document.querySelector('style[data-oa-office-ecode-loading]')) {
            var style = element('style');
            style.setAttribute('data-oa-office-ecode-loading', '1');
            style.textContent = DIRECT_LOADING_STYLE;
            (document.head || document.documentElement).appendChild(style);
        }
        var root = element('div', 'oa-office-ecode-loading');
        root.setAttribute('role', 'status');
        root.setAttribute('aria-live', 'polite');
        root.setAttribute('aria-busy', 'true');
        var panel = element('div', 'oa-office-ecode-loading-panel');
        var spinner = element('span', 'oa-office-ecode-loading-spinner');
        spinner.setAttribute('aria-hidden', 'true');
        var text = element('span', '', LOADING_TEXT);
        panel.appendChild(spinner);
        panel.appendChild(text);
        root.appendChild(panel);
        host.appendChild(root);
        directLoading = {root: root, text: text};
        if (!directLoadingPositionBound && window.addEventListener) {
            directLoadingPositionBound = true;
            window.addEventListener('resize', function () {
                placeDirectLoading(directLoadingFrame);
            }, false);
            window.addEventListener('scroll', function () {
                placeDirectLoading(directLoadingFrame);
            }, true);
        }
        return directLoading;
    }

    function showDirectLoading(frame, message) {
        if (message !== LOADING_COMPLETE_TEXT && directLoadingHideTimer) {
            if (window.clearTimeout) {
                window.clearTimeout(directLoadingHideTimer);
            }
            directLoadingHideTimer = null;
        }
        if (!frame) {
            return;
        }
        var loading = ensureDirectLoading();
        if (!loading) {
            return;
        }
        if (loading.boot) {
            bootLoading().show(message || LOADING_TEXT, frame);
            return;
        }
        loading.text.textContent = message || LOADING_TEXT;
        loading.root.setAttribute('aria-busy',
            message === LOADING_COMPLETE_TEXT ? 'false' : 'true');
        placeDirectLoading(frame);
    }

    function hideDirectLoading() {
        if (directLoadingHideTimer) {
            if (window.clearTimeout) {
                window.clearTimeout(directLoadingHideTimer);
            }
            directLoadingHideTimer = null;
        }
        var boot = bootLoading();
        if (boot && directLoading && directLoading.boot) {
            boot.hide();
            directLoading = null;
            directLoadingFrame = null;
            return;
        }
        if (directLoading && directLoading.root.parentNode) {
            directLoading.root.parentNode.removeChild(directLoading.root);
        }
        directLoading = null;
        directLoadingFrame = null;
    }

    function buildPreviewUrl(docId, imageFileId, extra) {
        var params = mergeMap(Object.create(null), extra || {}, currentParams());
        var url = new URL(contentPath() + PREVIEW_PATH, window.location.origin);
        url.searchParams.set('docId', String(docId));
        url.searchParams.set('imageFileId', String(imageFileId));
        AUTH_KEYS.forEach(function (key) {
            var value = pick(params, [key]);
            if (value) {
                url.searchParams.set(key, value);
            }
        });
        url.searchParams.set('officeunit.ecode', '1');
        return url.pathname + url.search;
    }

    function extensionOf(value) {
        var text = String(value || '').trim();
        if (!text) {
            return '';
        }
        try {
            var url = new URL(text, window.location.origin);
            var names = ['imagefilename', 'imageFileName', 'filename', 'fileName'];
            for (var i = 0; i < names.length; i += 1) {
                var named = url.searchParams.get(names[i]);
                if (named && named.lastIndexOf('.') >= 0) {
                    return named.substring(named.lastIndexOf('.') + 1).toLowerCase();
                }
            }
            text = url.pathname;
        } catch (ignore) {
            // Fall through to plain-text extension parsing.
        }
        text = text.split('?')[0].split('#')[0];
        var slash = text.lastIndexOf('/');
        var dot = text.lastIndexOf('.');
        return dot > slash ? text.substring(dot + 1).toLowerCase() : '';
    }

    function responseExtension(item) {
        var pathExtension = extensionOf(item && item.path);
        if (pathExtension) {
            return pathExtension;
        }
        var names = [
            'imagefilename', 'imageFileName', 'filename', 'fileName',
            'extension', 'fileext', 'suffix'
        ];
        for (var i = 0; i < names.length; i += 1) {
            var extension = extensionOf(item && item[names[i]]);
            if (extension) {
                return extension;
            }
        }
        var limitType = String(item && item.limitType || '').toLowerCase();
        var types = limitType.split(',');
        for (var j = 0; j < types.length; j += 1) {
            var type = types[j].trim();
            if (type && CONVERTIBLE_EXTENSIONS[type]) {
                return type;
            }
        }
        return types[0] && types[0].trim();
    }

    function shouldRewrite(data) {
        if (!data || typeof data !== 'object' || data.ispic === 1
                || data.ispic === '1') {
            return false;
        }
        var item = data.data && data.data[0];
        if (!item || typeof item !== 'object' || !item.path) {
            return false;
        }
        var path = String(item.path);
        if (/\/wui\/common\/page\/sysRemind(?:Docpreview)?\.jsp/i.test(path)
                || /\/notice\/noright\.jsp/i.test(path)) {
            return false;
        }
        var extension = responseExtension(item);
        if (!extension
                && /\/docs\/(?:pdfview[^/]*\/[^/?]*pdfViewer|e9\/iwebpdf)\.jsp/i
                    .test(path)) {
            return false;
        }
        return !extension || CONVERTIBLE_EXTENSIONS[extension] === true;
    }

    function rewriteApiData(url, params, data) {
        if (!isDocumentDetailPage() || String(url || '').indexOf(API_PATH) < 0) {
            return data;
        }
        if (!shouldRewrite(data)) {
            return data;
        }
        var apiParams = queryMap(url);
        if (params && typeof params === 'object') {
            apiParams = mergeMap(apiParams, params);
        }
        var current = currentParams();
        var docId = pick(apiParams, ['docid', 'docId', 'id'])
            || pick(current, ['id', 'docid', 'docId']);
        var imageFileId = pick(apiParams,
            ['fileid', 'fileId', 'imagefileid', 'imageFileId'])
            || pick(data || {}, ['fileid', 'fileId'])
            || pick(current, ['imagefileid', 'imageFileId']);
        if (!docId || !imageFileId) {
            return data;
        }
        var result = data && typeof data === 'object' ? data : {};
        var item = result.data && result.data[0] && typeof result.data[0] === 'object'
            ? result.data[0] : {};
        item.path = buildPreviewUrl(docId, imageFileId,
            mergeMap(apiParams, current));
        item.oaonesPreview = 'browser-js';
        result.result = 0;
        result.status = 1;
        result.fileid = Number(imageFileId);
        result.data = [item];
        showEarlyDocumentLoading();
        return result;
    }

    function isFailureFrame(src) {
        var value = String(src || '');
        var knownFailure = /\/docs\/pdfview\/web\/sysRemind\.jsp/i.test(value)
            || /\/wui\/common\/page\/sysRemind(?:Docpreview)?\.jsp/i.test(value)
            || /labelid=998/i.test(value);
        if (knownFailure) {
            return true;
        }
        try {
            var url = new URL(value, window.location.origin);
            return url.origin === window.location.origin
                && /\/docs\/pdfview\/(?:[^/]+\/)*pdfViewer\.jsp$/i.test(url.pathname)
                && extensionOf(value) === 'ofd';
        } catch (ignore) {
            return false;
        }
    }

    function isOfficePreviewFrame(frame) {
        if (!frame) {
            return false;
        }
        try {
            var url = new URL(frame.getAttribute('src') || frame.src,
                window.location.origin);
            var path = url.pathname;
            return (path === '/docs/docpreview/main.jsp'
                    || path === '/docs/view/main.jsp')
                && !!url.searchParams.get('docId')
                && !!url.searchParams.get('imageFileId')
                && url.searchParams.get('officeunit.original') !== '1';
        } catch (error) {
            return false;
        }
    }

    function hideFrameLoading(frame) {
        if (frame) {
            var state = frameLoadingState(frame);
            state.requested = false;
        }
        hideDirectLoading();
    }

    function isFrameConnected(frame) {
        return !!frame && (typeof frame.isConnected === 'boolean'
            ? frame.isConnected : document.documentElement.contains(frame));
    }

    function showEarlyDocumentLoading() {
        pendingDocumentLoading = true;
        var frame = document.getElementById('mainiframe');
        if (!isFrameConnected(frame)) {
            return;
        }
        showDirectLoading(frame, LOADING_TEXT);
    }

    function frameSource(frame) {
        if (!frame) {
            return '';
        }
        return String(frame.getAttribute('src') || frame.src || '');
    }

    function isBlankFrame(frame) {
        var source = frameSource(frame).trim();
        return !source || source === 'about:blank';
    }

    function frameLoadingState(frame) {
        var source = frameSource(frame);
        var state = frame.__oaonesLoadingState;
        if (!state || state.source !== source) {
            if (state && state.hideTimer && window.clearTimeout) {
                window.clearTimeout(state.hideTimer);
            }
            state = {
                source: source,
                requested: false,
                loaded: false,
                completed: false
            };
            frame.__oaonesLoadingState = state;
        }
        return state;
    }

    function showFrameLoading(frame) {
        if (!frame) {
            return;
        }
        var state = frameLoadingState(frame);
        if (state.requested || state.completed) {
            return;
        }
        state.requested = true;
        state.loaded = false;
        if (!frame.__oaonesLoadingBound) {
            frame.__oaonesLoadingBound = true;
            frame.addEventListener('load', function () {
                var loadedState = frameLoadingState(frame);
                loadedState.loaded = true;
            }, false);
        }
        showDirectLoading(frame, LOADING_TEXT);
    }

    function completeFrameLoading(frame, success) {
        if (!frame) {
            return;
        }
        var state = frameLoadingState(frame);
        if (state.completed) {
            return;
        }
        state.loaded = true;
        if (!state.requested && !directLoading) {
            state.completed = true;
            return;
        }
        state.completed = true;
        if (!success) {
            hideFrameLoading(frame);
            return;
        }
        showDirectLoading(frame, LOADING_COMPLETE_TEXT);
        pendingDocumentLoading = false;
        if (directLoadingHideTimer && window.clearTimeout) {
            window.clearTimeout(directLoadingHideTimer);
        }
        directLoadingHideTimer = window.setTimeout(function () {
            directLoadingHideTimer = null;
            hideFrameLoading(frame);
        }, LOADING_COMPLETE_DELAY);
    }

    function handleDocumentMessage(event) {
        if (!event || event.origin !== window.location.origin) {
            return;
        }
        var frame = document.getElementById('mainiframe');
        if (!frame || event.source !== frame.contentWindow) {
            return;
        }
        var type = event.data && event.data.type;
        if (type === DOCUMENT_READY_MESSAGE) {
            completeFrameLoading(frame, true);
        } else if (type === DOCUMENT_ERROR_MESSAGE) {
            completeFrameLoading(frame, false);
        }
    }

    function syncFrameLoading() {
        if (!isDocumentDetailPage()) {
            pendingDocumentLoading = false;
            hideFrameLoading(loadingFrame);
            loadingFrame = null;
            return;
        }
        var frame = document.getElementById('mainiframe');
        if (loadingFrame && loadingFrame !== frame) {
            hideFrameLoading(loadingFrame);
        }
        loadingFrame = frame;
        var frameConnected = isFrameConnected(frame);
        if (!frameConnected) {
            pendingDocumentLoading = false;
            hideFrameLoading(frame);
            return;
        }
        if (isBlankFrame(frame)) {
            showEarlyDocumentLoading();
            return;
        }
        if (!isOfficePreviewFrame(frame)) {
            pendingDocumentLoading = false;
            hideFrameLoading(frame);
            return;
        }
        pendingDocumentLoading = false;
        showFrameLoading(frame);
    }

    function forceMainFrame() {
        if (!isDocumentDetailPage()) {
            return;
        }
        var frame = document.getElementById('mainiframe');
        if (!frame || !isFailureFrame(frame.getAttribute('src') || frame.src)) {
            return;
        }
        var params = currentParams();
        var docId = pick(params, ['id', 'docid', 'docId']);
        var imageFileId = pick(params,
            ['imagefileid', 'imageFileId', 'fileid']);
        if (!docId || !imageFileId) {
            return;
        }
        var key = docId + ':' + imageFileId;
        if (attemptedFrames[key]) {
            return;
        }
        attemptedFrames[key] = true;
        var target = buildPreviewUrl(docId, imageFileId, params);
        frame.setAttribute('data-oaones-officeunit-target', target);
        frame.src = target;
    }

    function scheduleFrameCheck() {
        if (scheduled) {
            return;
        }
        scheduled = true;
        var run = function () {
            scheduled = false;
            forceMainFrame();
            syncFrameLoading();
        };
        if (window.requestAnimationFrame) {
            window.requestAnimationFrame(run);
        } else {
            window.setTimeout(run, 0);
        }
    }

    function installDomGuard() {
        if (!document.body || !window.MutationObserver) {
            return;
        }
        if (!observer) {
            observer = new MutationObserver(scheduleFrameCheck);
            observer.observe(document.body, {
                childList: true,
                subtree: true,
                attributes: true,
                attributeFilter: ['src']
            });
        }
        window.addEventListener('hashchange', scheduleFrameCheck, false);
        window.addEventListener('popstate', scheduleFrameCheck, false);
        window.setTimeout(scheduleFrameCheck, 0);
    }

    try {
        if (window.ecodeSDK
                && typeof window.ecodeSDK.rewriteApiDataQueueSet === 'function') {
            window.ecodeSDK.rewriteApiDataQueueSet({
                fn: rewriteApiData,
                desc: 'OAones OfficeUnit: route SPA document preview to browser viewer'
            });
        } else if (window.console && console.warn) {
            console.warn('[OAones OfficeUnit] ecodeSDK rewrite hook is unavailable');
        }
    } catch (error) {
        if (window.console && console.error) {
            console.error('[OAones OfficeUnit] failed to register preview rewrite', error);
        }
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', installDomGuard, {once: true});
    } else {
        installDomGuard();
    }
    if (window.addEventListener) {
        window.addEventListener('message', handleDocumentMessage, false);
    }

    window.oaonesOfficeUnitEcode = {
        rewriteApiData: rewriteApiData,
        previewUrl: buildPreviewUrl,
        forceCurrentFrame: forceMainFrame,
        syncCurrentFrameLoading: syncFrameLoading
    };
})(window, document);
