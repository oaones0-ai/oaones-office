(function (global) {
    'use strict';

    var FORBIDDEN = new Set([
        'script', 'iframe', 'frame', 'frameset', 'object', 'embed', 'applet',
        'form', 'input', 'button', 'textarea', 'select', 'option', 'link',
        'meta', 'base', 'audio', 'video', 'source', 'track', 'canvas',
        'portal', 'fencedframe'
    ]);
    var ALLOWED = new Set([
        'div', 'section', 'article', 'header', 'footer', 'main', 'aside', 'nav',
        'span', 'p', 'br', 'hr', 'a', 'img', 'figure', 'figcaption',
        'table', 'caption', 'colgroup', 'col', 'thead', 'tbody', 'tfoot', 'tr',
        'td', 'th', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'ul', 'ol', 'li',
        'strong', 'b', 'em', 'i', 'u', 's', 'strike', 'sub', 'sup', 'small',
        'blockquote', 'pre', 'code', 'dl', 'dt', 'dd', 'style',
        'svg', 'g', 'path', 'rect', 'circle', 'ellipse', 'line', 'polyline',
        'polygon', 'text', 'tspan', 'defs', 'clippath', 'lineargradient',
        'radialgradient', 'stop', 'mask'
    ]);
    var ATTRIBUTE_ALLOW = new Set([
        'class', 'style', 'title', 'dir', 'lang', 'role', 'scope',
        'colspan', 'rowspan', 'width', 'height', 'align', 'valign',
        'src', 'alt', 'viewbox', 'd', 'x', 'y', 'x1', 'x2', 'y1', 'y2',
        'cx', 'cy', 'r', 'rx', 'ry', 'points', 'fill', 'stroke',
        'stroke-width', 'transform', 'opacity', 'offset', 'stop-color',
        'stop-opacity', 'clip-path', 'mask', 'preserveaspectratio'
    ]);

    function cleanCss(value) {
        if (!value) {
            return '';
        }
        var css = String(value);
        css = css.replace(/@import[\s\S]*?(?:;|$)/gi, '')
            .replace(/expression\s*\(/gi, '')
            .replace(/behavior\s*:/gi, '')
            .replace(/-moz-binding\s*:/gi, '')
            .replace(/javascript\s*:/gi, '')
            .replace(/vbscript\s*:/gi, '');
        return css.replace(/url\s*\(\s*(['"]?)(.*?)\1\s*\)/gi,
            function (all, quote, url) {
                var target = String(url).trim();
                if (/^blob:/i.test(target)
                        || /^data:image\/(?:png|jpe?g|gif|webp|bmp);/i.test(target)
                        || /^data:font\/(?:woff2?|ttf|otf);/i.test(target)
                        || target.charAt(0) === '#') {
                    return 'url(' + quote + target + quote + ')';
                }
                return 'none';
            });
    }

    function safeImageUrl(url) {
        var value = String(url || '').trim();
        return /^blob:/i.test(value)
            || /^data:image\/(?:png|jpe?g|gif|webp|bmp|tiff);/i.test(value);
    }

    function safeLinkUrl(url) {
        var value = String(url || '').trim();
        if (!value || /[\u0000-\u0020\u007f]/.test(value)) {
            return '';
        }
        if (value.charAt(0) === '#') {
            return value;
        }
        if (!/^(?:https?|mailto):/i.test(value)
                || typeof global.URL !== 'function') {
            return '';
        }
        try {
            var parsed = new global.URL(value);
            var protocol = parsed.protocol.toLowerCase();
            if ((protocol === 'http:' || protocol === 'https:')
                    && parsed.hostname) {
                return value;
            }
            if (protocol === 'mailto:' && parsed.pathname) {
                return value;
            }
        } catch (ignore) {
            // Malformed URLs remain as plain text.
        }
        return '';
    }

    function sanitizeElement(element, shadowRoot) {
        if (!element || element.nodeType !== 1) {
            return;
        }
        var tag = element.localName;
        var linkTarget = tag === 'a' ? safeLinkUrl(
            element.getAttribute('href')
                || element.getAttribute('xlink:href') || '') : '';
        if (FORBIDDEN.has(tag)) {
            element.remove();
            return;
        }
        if (!ALLOWED.has(tag)) {
            var parent = element.parentNode;
            if (!parent) {
                return;
            }
            var unknownChild = element.firstChild;
            while (unknownChild) {
                var unknownNext = unknownChild.nextSibling;
                sanitizeElement(unknownChild, shadowRoot);
                unknownChild = unknownNext;
            }
            while (element.firstChild) {
                parent.insertBefore(element.firstChild, element);
            }
            parent.removeChild(element);
            return;
        }
        var attributes = Array.prototype.slice.call(element.attributes || []);
        attributes.forEach(function (attribute) {
            var name = attribute.name.toLowerCase();
            var value = attribute.value;
            if (name.indexOf('on') === 0 || name === 'srcdoc'
                    || name === 'formaction' || name === 'action') {
                element.removeAttribute(attribute.name);
                return;
            }
            if (name === 'href' || name === 'xlink:href') {
                element.removeAttribute(attribute.name);
                return;
            }
            if (name === 'src') {
                if (!safeImageUrl(value)) {
                    element.removeAttribute(attribute.name);
                }
                return;
            }
            if (name === 'style') {
                element.setAttribute('style', cleanCss(value));
                return;
            }
            if (name.indexOf('data-') === 0 || ATTRIBUTE_ALLOW.has(name)
                    || name.indexOf('aria-') === 0) {
                return;
            }
            element.removeAttribute(attribute.name);
        });
        if (tag === 'a') {
            element.removeAttribute('href');
            element.removeAttribute('target');
            element.removeAttribute('rel');
            if (linkTarget) {
                element.setAttribute('href', linkTarget);
                element.setAttribute('data-oa-link',
                    linkTarget.charAt(0) === '#' ? 'anchor' : 'external');
                if (linkTarget.charAt(0) !== '#') {
                    element.setAttribute('target', '_blank');
                    element.setAttribute('rel', 'noopener noreferrer');
                    element.setAttribute('referrerpolicy', 'no-referrer');
                }
            } else {
                element.setAttribute('role', 'text');
            }
        }
        if (tag === 'style') {
            element.textContent = cleanCss(element.textContent || '');
            if (shadowRoot && element.parentNode !== shadowRoot) {
                shadowRoot.appendChild(element);
                return;
            }
        }
        var child = element.firstChild;
        while (child) {
            var next = child.nextSibling;
            sanitizeElement(child, shadowRoot);
            child = next;
        }
    }

    function sanitizeDom(root, shadowRoot) {
        if (!root) {
            return root;
        }
        var child = root.firstChild;
        while (child) {
            var next = child.nextSibling;
            sanitizeElement(child, shadowRoot);
            child = next;
        }
        return root;
    }

    function collectBlobUrls(value) {
        var urls = new Set();
        var seen = new WeakSet();
        function walk(item) {
            if (typeof item === 'string') {
                if (/^blob:/i.test(item)) {
                    urls.add(item);
                }
                return;
            }
            if (!item || typeof item !== 'object' || seen.has(item)) {
                return;
            }
            seen.add(item);
            if (typeof item.url === 'string' && /^blob:/i.test(item.url)) {
                urls.add(item.url);
            }
            if (Array.isArray(item)) {
                item.forEach(walk);
                return;
            }
            Object.keys(item).forEach(function (key) {
                walk(item[key]);
            });
        }
        walk(value);
        return Array.from(urls);
    }

    global.OfficeSecurity = {
        sanitizeDom: sanitizeDom,
        cleanCss: cleanCss,
        safeImageUrl: safeImageUrl,
        safeLinkUrl: safeLinkUrl,
        collectBlobUrls: collectBlobUrls
    };
})(self);
