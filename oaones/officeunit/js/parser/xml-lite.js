(function (global) {
    'use strict';

    var MAX_NODES = 250000;
    var MAX_DEPTH = 256;

    function localName(name) {
        var index = name.indexOf(':');
        return (index >= 0 ? name.substring(index + 1) : name).toLowerCase();
    }

    function decodeEntities(text) {
        return text.replace(/&(#x[0-9a-f]+|#[0-9]+|amp|lt|gt|quot|apos);/gi,
            function (all, entity) {
                var value = entity.toLowerCase();
                if (value === 'amp') {
                    return '&';
                }
                if (value === 'lt') {
                    return '<';
                }
                if (value === 'gt') {
                    return '>';
                }
                if (value === 'quot') {
                    return '"';
                }
                if (value === 'apos') {
                    return "'";
                }
                var number = value.indexOf('#x') === 0
                    ? parseInt(value.substring(2), 16)
                    : parseInt(value.substring(1), 10);
                if (!isFinite(number) || number < 0 || number > 0x10ffff) {
                    throw new Error('XML 实体数值无效');
                }
                return String.fromCodePoint(number);
            });
    }

    function parseAttributes(source, node) {
        var re = /([A-Za-z_][\w:.-]*)\s*=\s*("([^"]*)"|'([^']*)')/g;
        var match;
        while ((match = re.exec(source))) {
            var name = match[1];
            var value = match[3] !== undefined ? match[3] : match[4];
            node.attrs[name] = decodeEntities(value);
            node.attrs[localName(name)] = node.attrs[name];
        }
    }

    function parse(text) {
        if (typeof text !== 'string') {
            throw new Error('XML 内容为空');
        }
        if (/<!doctype/i.test(text) || /<!entity/i.test(text)) {
            throw new Error('XML 包含不允许的 DTD 或实体');
        }
        var root = {
            name: '#document',
            local: '#document',
            attrs: {},
            children: [],
            text: '',
            parent: null
        };
        var stack = [root];
        var nodes = 0;
        var index = 0;
        var tagRe = /<\?[\s\S]*?\?>|<!--[\s\S]*?-->|<!\[CDATA\[([\s\S]*?)\]\]>|<\/\s*([A-Za-z_][\w:.-]*)\s*>|<([A-Za-z_][\w:.-]*)([\s\S]*?)(\/?)>/g;
        var match;
        while ((match = tagRe.exec(text))) {
            if (match.index > index) {
                appendText(stack[stack.length - 1], text.substring(index, match.index));
            }
            index = tagRe.lastIndex;
            if (match[1] !== undefined) {
                appendText(stack[stack.length - 1], match[1]);
                continue;
            }
            if (match[2]) {
                var closing = match[2];
                var expected = stack[stack.length - 1];
                if (stack.length <= 1 || expected.name !== closing) {
                    throw new Error('XML 标签不匹配');
                }
                stack.pop();
                continue;
            }
            if (match[3]) {
                nodes += 1;
                if (nodes > MAX_NODES || stack.length > MAX_DEPTH) {
                    throw new Error('XML 结构超过限制');
                }
                var name = match[3];
                var attributes = match[4] || '';
                var empty = match[5] === '/';
                var node = {
                    name: name,
                    local: localName(name),
                    attrs: {},
                    children: [],
                    text: '',
                    parent: stack[stack.length - 1]
                };
                parseAttributes(attributes, node);
                stack[stack.length - 1].children.push(node);
                if (!empty) {
                    stack.push(node);
                }
            }
        }
        if (tagRe.lastIndex < text.length) {
            appendText(stack[stack.length - 1], text.substring(tagRe.lastIndex));
        }
        if (stack.length !== 1 || root.children.length !== 1) {
            throw new Error('XML 根节点无效');
        }
        return root.children[0];
    }

    function appendText(node, raw) {
        if (!raw) {
            return;
        }
        if (node.text.length + raw.length > 4000000) {
            throw new Error('XML 文本超过限制');
        }
        node.text += decodeEntities(raw);
    }

    function child(node, name) {
        if (!node) {
            return null;
        }
        var wanted = name.toLowerCase();
        for (var i = 0; i < node.children.length; i += 1) {
            if (node.children[i].local === wanted) {
                return node.children[i];
            }
        }
        return null;
    }

    function children(node, name) {
        if (!node) {
            return [];
        }
        if (!name) {
            return node.children;
        }
        var wanted = name.toLowerCase();
        return node.children.filter(function (item) {
            return item.local === wanted;
        });
    }

    function descendants(node, name) {
        var wanted = name ? name.toLowerCase() : '';
        var result = [];
        var stack = node ? [node] : [];
        while (stack.length) {
            var current = stack.pop();
            if (!wanted || current.local === wanted) {
                result.push(current);
            }
            for (var i = current.children.length - 1; i >= 0; i -= 1) {
                stack.push(current.children[i]);
            }
        }
        return result;
    }

    function firstDescendant(node, name) {
        var values = descendants(node, name);
        return values.length ? values[0] : null;
    }

    function attr(node, name, defaultValue) {
        if (!node || !node.attrs) {
            return defaultValue;
        }
        var value = node.attrs[name];
        if (value === undefined) {
            value = node.attrs[localName(name)];
        }
        return value === undefined ? defaultValue : value;
    }

    function text(node) {
        if (!node) {
            return '';
        }
        var value = node.text || '';
        node.children.forEach(function (item) {
            value += text(item);
        });
        return value;
    }

    global.OfficeXmlLite = {
        parse: parse,
        child: child,
        children: children,
        descendants: descendants,
        firstDescendant: firstDescendant,
        attr: attr,
        text: text,
        localName: localName
    };
})(self);
