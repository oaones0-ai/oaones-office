(function (global) {
    'use strict';

    var VERSION = '5.65.21';
    var script = document.currentScript;
    var assetBase = deriveAssetBase(script && script.src);
    var scriptLoads = Object.create(null);
    var MODES = createModes();
    var ALIASES = {
        txt: 'plain',
        log: 'plain',
        text: 'plain',
        cfg: 'plain',
        py: 'python',
        pyw: 'python',
        js: 'javascript',
        jsx: 'jsx',
        ts: 'javascript',
        tsx: 'jsx',
        java: 'clike',
        c: 'clike',
        h: 'clike',
        cpp: 'clike',
        cc: 'clike',
        cxx: 'clike',
        hpp: 'clike',
        cs: 'clike',
        kt: 'clike',
        scala: 'clike',
        m: 'clike',
        mm: 'clike',
        go: 'go',
        rs: 'rust',
        php: 'php',
        rb: 'ruby',
        lua: 'lua',
        r: 'r',
        sql: 'sql',
        json: 'json',
        xml: 'xml',
        html: 'html',
        htm: 'html',
        css: 'css',
        scss: 'css',
        less: 'css',
        yaml: 'yaml',
        yml: 'yaml',
        md: 'markdown',
        markdown: 'markdown',
        sh: 'shell',
        bash: 'shell',
        zsh: 'shell',
        bat: 'shell',
        cmd: 'shell',
        ps1: 'powershell',
        psd1: 'powershell',
        psm1: 'powershell',
        properties: 'properties',
        ini: 'properties',
        in: 'properties',
        toml: 'toml',
        swift: 'swift',
        vue: 'vue',
        diff: 'diff',
        patch: 'diff',
        vb: 'vb',
        vbs: 'vb',
        groovy: 'groovy',
        gradle: 'groovy',
        dart: 'dart',
        coffee: 'coffeescript',
        pl: 'perl',
        pm: 'perl',
        conf: 'nginx'
    };
    var MIME_BY_EXTENSION = {
        ts: 'application/typescript',
        java: 'text/x-java',
        c: 'text/x-csrc',
        h: 'text/x-csrc',
        cpp: 'text/x-c++src',
        cc: 'text/x-c++src',
        cxx: 'text/x-c++src',
        hpp: 'text/x-c++src',
        cs: 'text/x-csharp',
        kt: 'text/x-kotlin',
        scala: 'text/x-scala',
        m: 'text/x-objectivec',
        mm: 'text/x-objectivec++',
        scss: 'text/x-scss',
        less: 'text/x-less',
        bat: 'text/x-sh',
        cmd: 'text/x-sh'
    };

    function createModes() {
        return {
            plain: {
                label: '纯文本',
                mode: null,
                files: []
            },
            python: {
                label: 'Python',
                mode: 'python',
                files: ['mode/python/python.js']
            },
            javascript: {
                label: 'JavaScript / TypeScript',
                mode: 'javascript',
                files: ['mode/javascript/javascript.js']
            },
            jsx: {
                label: 'JSX / TSX',
                mode: 'jsx',
                files: [
                    'mode/javascript/javascript.js',
                    'mode/xml/xml.js',
                    'mode/jsx/jsx.js'
                ]
            },
            clike: {
                label: 'C / C++ / Java',
                mode: 'clike',
                files: ['mode/clike/clike.js']
            },
            go: {
                label: 'Go',
                mode: 'go',
                files: ['mode/go/go.js']
            },
            rust: {
                label: 'Rust',
                mode: 'rust',
                files: [
                    'addon/mode/simple.js',
                    'mode/rust/rust.js'
                ]
            },
            php: {
                label: 'PHP',
                mode: 'php',
                files: [
                    'mode/xml/xml.js',
                    'mode/javascript/javascript.js',
                    'mode/css/css.js',
                    'mode/htmlmixed/htmlmixed.js',
                    'mode/clike/clike.js',
                    'mode/php/php.js'
                ]
            },
            ruby: {
                label: 'Ruby',
                mode: 'ruby',
                files: ['mode/ruby/ruby.js']
            },
            lua: {
                label: 'Lua',
                mode: 'lua',
                files: ['mode/lua/lua.js']
            },
            r: {
                label: 'R',
                mode: 'r',
                files: ['mode/r/r.js']
            },
            sql: {
                label: 'SQL',
                mode: 'sql',
                files: ['mode/sql/sql.js']
            },
            json: {
                label: 'JSON',
                mode: {name: 'javascript', json: true},
                files: ['mode/javascript/javascript.js']
            },
            xml: {
                label: 'XML',
                mode: {name: 'xml'},
                files: ['mode/xml/xml.js']
            },
            html: {
                label: 'HTML',
                mode: 'htmlmixed',
                files: [
                    'mode/xml/xml.js',
                    'mode/javascript/javascript.js',
                    'mode/css/css.js',
                    'mode/htmlmixed/htmlmixed.js'
                ]
            },
            css: {
                label: 'CSS / SCSS / LESS',
                mode: 'css',
                files: ['mode/css/css.js']
            },
            yaml: {
                label: 'YAML',
                mode: 'yaml',
                files: ['mode/yaml/yaml.js']
            },
            markdown: {
                label: 'Markdown',
                mode: 'markdown',
                files: [
                    'mode/meta.js',
                    'mode/xml/xml.js',
                    'mode/markdown/markdown.js'
                ]
            },
            shell: {
                label: 'Shell / Batch',
                mode: 'shell',
                files: ['mode/shell/shell.js']
            },
            powershell: {
                label: 'PowerShell',
                mode: 'powershell',
                files: ['mode/powershell/powershell.js']
            },
            properties: {
                label: 'Properties / INI',
                mode: 'properties',
                files: ['mode/properties/properties.js']
            },
            toml: {
                label: 'TOML',
                mode: 'toml',
                files: ['mode/toml/toml.js']
            },
            swift: {
                label: 'Swift',
                mode: 'swift',
                files: ['mode/swift/swift.js']
            },
            vue: {
                label: 'Vue',
                mode: 'vue',
                files: [
                    'addon/mode/simple.js',
                    'addon/mode/multiplex.js',
                    'addon/mode/overlay.js',
                    'mode/xml/xml.js',
                    'mode/javascript/javascript.js',
                    'mode/css/css.js',
                    'mode/htmlmixed/htmlmixed.js',
                    'mode/coffeescript/coffeescript.js',
                    'mode/sass/sass.js',
                    'mode/stylus/stylus.js',
                    'mode/pug/pug.js',
                    'mode/handlebars/handlebars.js',
                    'mode/vue/vue.js'
                ]
            },
            diff: {
                label: 'Diff / Patch',
                mode: 'diff',
                files: ['mode/diff/diff.js']
            },
            vb: {
                label: 'Visual Basic',
                mode: 'vb',
                files: ['mode/vb/vb.js']
            },
            groovy: {
                label: 'Groovy',
                mode: 'groovy',
                files: ['mode/groovy/groovy.js']
            },
            dart: {
                label: 'Dart',
                mode: 'dart',
                files: [
                    'mode/clike/clike.js',
                    'mode/dart/dart.js'
                ]
            },
            coffeescript: {
                label: 'CoffeeScript',
                mode: 'coffeescript',
                files: ['mode/coffeescript/coffeescript.js']
            },
            perl: {
                label: 'Perl',
                mode: 'perl',
                files: ['mode/perl/perl.js']
            },
            nginx: {
                label: 'Nginx 配置',
                mode: 'nginx',
                files: ['mode/nginx/nginx.js']
            }
        };
    }

    function render(options) {
        var container = options.container;
        var shadowRoot = options.shadowRoot;
        var extension = String(options.extension || '').toLowerCase();
        var modeId = ALIASES[extension] || 'plain';
        var definition = MODES[modeId] || MODES.plain;
        var mode = withMime(definition.mode, extension);
        container.replaceChildren();
        var root = element('div', 'oa-text-viewer');
        var status = element('div', 'oa-text-viewer-status');
        var language = element('span', 'oa-text-viewer-language');
        language.appendChild(element('i', 'oa-text-viewer-dot'));
        language.appendChild(document.createTextNode(definition.label));
        var encoding = element('span', 'oa-text-viewer-meta');
        var wrap = element('button', 'oa-text-viewer-wrap', '自动换行');
        wrap.type = 'button';
        wrap.setAttribute('aria-pressed', 'false');
        status.appendChild(language);
        status.appendChild(encoding);
        status.appendChild(wrap);
        var editorHost = element('div', 'oa-text-viewer-editor');
        root.appendChild(status);
        root.appendChild(editorHost);
        container.appendChild(root);

        return loadStyles(shadowRoot)
            .then(function () {
                return loadScripts(definition.files);
            })
            .then(function () {
                if (options.destroyed()) {
                    return null;
                }
                if (typeof global.CodeMirror !== 'function') {
                    throw previewError('OFFICE_PREVIEW_RENDER_FAILED',
                        '文本预览组件未加载');
                }
                var decoded = decodeText(options.buffer);
                encoding.textContent = decoded.encoding;
                var lineWrapping = /^(txt|log|text|md|markdown)$/.test(extension);
                var editor = global.CodeMirror(editorHost, {
                    value: decoded.text,
                    mode: mode,
                    theme: 'material-ocean',
                    lineNumbers: true,
                    lineWrapping: lineWrapping,
                    readOnly: true,
                    styleActiveLine: {nonEmpty: true},
                    matchBrackets: true,
                    foldGutter: true,
                    gutters: [
                        'CodeMirror-linenumbers',
                        'CodeMirror-foldgutter'
                    ],
                    indentUnit: 4,
                    tabSize: 4,
                    viewportMargin: 20,
                    scrollbarStyle: 'native'
                });
                editor.setSize('100%', '100%');
                wrap.setAttribute('aria-pressed', lineWrapping ? 'true' : 'false');
                wrap.classList.toggle('active', lineWrapping);
                wrap.addEventListener('click', function () {
                    lineWrapping = !lineWrapping;
                    editor.setOption('lineWrapping', lineWrapping);
                    wrap.setAttribute('aria-pressed',
                        lineWrapping ? 'true' : 'false');
                    wrap.classList.toggle('active', lineWrapping);
                    editor.refresh();
                });

                var refresh = function () {
                    if (!options.destroyed() && editor) {
                        editor.refresh();
                    }
                };
                var frame = global.requestAnimationFrame
                    ? global.requestAnimationFrame(refresh)
                    : global.setTimeout(refresh, 0);
                var resizeObserver = null;
                if (global.ResizeObserver) {
                    resizeObserver = new global.ResizeObserver(refresh);
                    resizeObserver.observe(editorHost);
                } else {
                    global.addEventListener('resize', refresh, false);
                }
                var lineCount = editor.lineCount();
                return {
                    sections: [{
                        id: 'document',
                        label: definition.label + ' · ' + lineCount + ' 行',
                        kind: 'document',
                        index: 0
                    }],
                    warnings: [],
                    setZoom: function () {
                        return 1;
                    },
                    fitWidth: function () {
                        return 1;
                    },
                    goTo: function (anchor) {
                        var match = /(?:^|\/)line-(\d+)$/i.exec(String(anchor || ''));
                        if (!match) {
                            return false;
                        }
                        var line = Math.max(0, Math.min(
                            editor.lineCount() - 1,
                            parseInt(match[1], 10) - 1
                        ));
                        editor.setCursor({line: line, ch: 0});
                        editor.scrollIntoView({line: line, ch: 0}, 80);
                        editor.focus();
                        return true;
                    },
                    getCurrentAnchor: function () {
                        return 'line-' + (editor.getCursor().line + 1);
                    },
                    getEditor: function () {
                        return editor;
                    },
                    destroy: function () {
                        if (resizeObserver) {
                            resizeObserver.disconnect();
                        } else {
                            global.removeEventListener('resize', refresh, false);
                        }
                        if (global.cancelAnimationFrame && typeof frame === 'number') {
                            global.cancelAnimationFrame(frame);
                        } else if (global.clearTimeout) {
                            global.clearTimeout(frame);
                        }
                        if (editor) {
                            editor.getWrapperElement().remove();
                            editor = null;
                        }
                        container.replaceChildren();
                    }
                };
            });
    }

    function withMime(mode, extension) {
        var mime = MIME_BY_EXTENSION[extension];
        if (!mime || !mode) {
            return mode;
        }
        if (typeof mode === 'string') {
            return mime;
        }
        return mode;
    }

    function loadStyles(shadowRoot) {
        var files = [
            'lib/codemirror.css',
            'theme/material-ocean.css',
            'addon/fold/foldgutter.css'
        ];
        return Promise.all(files.map(function (file) {
            var url = assetUrl(file);
            var existing = shadowRoot.querySelector(
                'link[data-oa-codemirror-style="' + cssAttributeValue(url) + '"]'
            );
            if (existing) {
                return Promise.resolve();
            }
            return new Promise(function (resolve, reject) {
                var link = document.createElement('link');
                link.rel = 'stylesheet';
                link.href = url;
                link.setAttribute('data-oa-codemirror-style', url);
                link.onload = resolve;
                link.onerror = function () {
                    reject(previewError('OFFICE_PREVIEW_RENDER_FAILED',
                        '文本预览样式加载失败'));
                };
                shadowRoot.appendChild(link);
            });
        }));
    }

    function loadScripts(files) {
        var chain = Promise.resolve();
        files.forEach(function (file) {
            chain = chain.then(function () {
                return loadScript(file);
            });
        });
        return chain;
    }

    function loadScript(file) {
        var url = assetUrl(file);
        if (scriptLoads[url]) {
            return scriptLoads[url];
        }
        scriptLoads[url] = new Promise(function (resolve, reject) {
            var target = document.querySelector(
                'script[data-oa-codemirror-script="' + cssAttributeValue(url) + '"]'
            );
            if (target) {
                resolve();
                return;
            }
            target = document.createElement('script');
            target.src = url;
            target.async = false;
            target.setAttribute('data-oa-codemirror-script', url);
            target.onload = resolve;
            target.onerror = function () {
                delete scriptLoads[url];
                reject(previewError('OFFICE_PREVIEW_RENDER_FAILED',
                    '文本语法着色组件加载失败'));
            };
            (document.head || document.documentElement).appendChild(target);
        });
        return scriptLoads[url];
    }

    function decodeText(buffer) {
        var bytes = new Uint8Array(buffer);
        var encoding = detectedEncoding(bytes);
        var candidates = encoding
            ? [encoding]
            : ['utf-8', 'gb18030'];
        var failure = null;
        for (var index = 0; index < candidates.length; index += 1) {
            try {
                var decoder = new TextDecoder(candidates[index], {fatal: true});
                var value = decoder.decode(bytes);
                if (value.charCodeAt(0) === 0xFEFF) {
                    value = value.substring(1);
                }
                return {
                    text: value,
                    encoding: browserEncodingName(candidates[index])
                };
            } catch (error) {
                failure = error;
            }
        }
        throw previewError('OFFICE_PREVIEW_PARSE_FAILED',
            failure ? '文本编码无法识别' : '无法读取文本内容');
    }

    function detectedEncoding(bytes) {
        if (bytes.length >= 3 && bytes[0] === 0xEF && bytes[1] === 0xBB
                && bytes[2] === 0xBF) {
            return 'utf-8';
        }
        if (bytes.length >= 2 && bytes[0] === 0xFF && bytes[1] === 0xFE) {
            return 'utf-16le';
        }
        if (bytes.length >= 2 && bytes[0] === 0xFE && bytes[1] === 0xFF) {
            return 'utf-16be';
        }
        var sample = Math.min(bytes.length, 200);
        var oddZero = 0;
        var evenZero = 0;
        for (var index = 0; index < sample; index += 1) {
            if (bytes[index] !== 0) {
                continue;
            }
            if (index % 2 === 0) {
                evenZero += 1;
            } else {
                oddZero += 1;
            }
        }
        if (sample >= 8 && oddZero > sample / 4 && evenZero < sample / 20) {
            return 'utf-16le';
        }
        if (sample >= 8 && evenZero > sample / 4 && oddZero < sample / 20) {
            return 'utf-16be';
        }
        return '';
    }

    function browserEncodingName(encoding) {
        if (encoding === 'utf-8') {
            return 'UTF-8';
        }
        if (encoding === 'utf-16le' || encoding === 'utf-16be') {
            return encoding.toUpperCase().replace('LE', ' LE').replace('BE', ' BE');
        }
        return 'GB18030';
    }

    function previewError(code, message) {
        var error = new Error(message || '文本预览失败');
        error.name = 'OfficePreviewError';
        error.code = code || 'OFFICE_PREVIEW_RENDER_FAILED';
        return error;
    }

    function deriveAssetBase(scriptUrl) {
        try {
            var url = new URL(scriptUrl, global.location.href);
            var marker = '/js/render/text-renderer.js';
            var index = url.pathname.indexOf(marker);
            if (index >= 0) {
                return url.pathname.substring(0, index)
                    + '/js/vendor/codemirror/';
            }
            return '/oaones/officeunit/js/vendor/codemirror/';
        } catch (ignore) {
            return '/oaones/officeunit/js/vendor/codemirror/';
        }
    }

    function assetUrl(file) {
        return assetBase + file + '?v=' + VERSION;
    }

    function cssAttributeValue(value) {
        return String(value || '').replace(/\\/g, '\\\\').replace(/"/g, '\\"');
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

    global.OfficeRenderers = global.OfficeRenderers || {};
    global.OfficeRenderers.text = {render: render};
})(window);
