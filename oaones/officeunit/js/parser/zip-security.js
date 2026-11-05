(function (global) {
    'use strict';

    var decoder = new TextDecoder('utf-8', {fatal: false});
    var XML_RE = /\.(xml|rels)$/i;

    function PreviewError(code, message) {
        this.name = 'OfficePreviewError';
        this.code = code;
        this.message = message || code;
    }
    PreviewError.prototype = Object.create(Error.prototype);
    PreviewError.prototype.constructor = PreviewError;

    function u16(view, offset, end) {
        if (offset + 2 > end) {
            throw new PreviewError('OFFICE_PREVIEW_ARCHIVE_INVALID', 'ZIP 结构不完整');
        }
        return view.getUint16(offset, true);
    }

    function u32(view, offset, end) {
        if (offset + 4 > end) {
            throw new PreviewError('OFFICE_PREVIEW_ARCHIVE_INVALID', 'ZIP 结构不完整');
        }
        return view.getUint32(offset, true);
    }

    function normalizeName(name) {
        if (!name || name.charAt(0) === '/' || name.indexOf('\\') >= 0
                || name.indexOf('\u0000') >= 0 || /^[a-z]:/i.test(name)) {
            throw new PreviewError('OFFICE_PREVIEW_ARCHIVE_INVALID', 'ZIP 条目路径不安全');
        }
        var parts = name.split('/');
        for (var i = 0; i < parts.length; i += 1) {
            if (parts[i] === '..') {
                throw new PreviewError('OFFICE_PREVIEW_ARCHIVE_INVALID',
                    'ZIP 条目包含目录穿越');
            }
        }
        return name;
    }

    function findEnd(view, length) {
        var minimum = Math.max(0, length - 65557);
        for (var offset = length - 22; offset >= minimum; offset -= 1) {
            if (view.getUint32(offset, true) === 0x06054b50) {
                return offset;
            }
        }
        throw new PreviewError('OFFICE_PREVIEW_ARCHIVE_INVALID', '未找到 ZIP 中央目录');
    }

    function inspect(arrayBuffer, limits) {
        var bytes = new Uint8Array(arrayBuffer);
        var view = new DataView(arrayBuffer);
        var length = bytes.length;
        var eocd = findEnd(view, length);
        var disk = u16(view, eocd + 4, length);
        var centralDisk = u16(view, eocd + 6, length);
        var diskEntries = u16(view, eocd + 8, length);
        var totalEntries = u16(view, eocd + 10, length);
        var centralSize = u32(view, eocd + 12, length);
        var centralOffset = u32(view, eocd + 16, length);
        if (totalEntries === 0xffff || centralSize === 0xffffffff
                || centralOffset === 0xffffffff) {
            throw new PreviewError('OFFICE_PREVIEW_ARCHIVE_INVALID',
                '暂不支持 ZIP64 文档');
        }
        if (disk !== 0 || centralDisk !== 0 || diskEntries !== totalEntries) {
            throw new PreviewError('OFFICE_PREVIEW_ARCHIVE_INVALID', '不支持分卷 ZIP');
        }
        if (totalEntries > limits.maxArchiveEntries) {
            throw new PreviewError('OFFICE_PREVIEW_ARCHIVE_TOO_LARGE', 'ZIP 条目过多');
        }
        if (centralOffset + centralSize > length) {
            throw new PreviewError('OFFICE_PREVIEW_ARCHIVE_INVALID', 'ZIP 中央目录越界');
        }

        var entries = [];
        var names = new Set();
        var totalUncompressed = 0;
        var offset = centralOffset;
        var end = centralOffset + centralSize;
        for (var index = 0; index < totalEntries; index += 1) {
            if (u32(view, offset, end) !== 0x02014b50) {
                throw new PreviewError('OFFICE_PREVIEW_ARCHIVE_INVALID', 'ZIP 中央目录损坏');
            }
            var flags = u16(view, offset + 8, end);
            var method = u16(view, offset + 10, end);
            var compressedSize = u32(view, offset + 20, end);
            var uncompressedSize = u32(view, offset + 24, end);
            var nameLength = u16(view, offset + 28, end);
            var extraLength = u16(view, offset + 30, end);
            var commentLength = u16(view, offset + 32, end);
            var localOffset = u32(view, offset + 42, end);
            if (offset + 46 + nameLength + extraLength + commentLength > end) {
                throw new PreviewError('OFFICE_PREVIEW_ARCHIVE_INVALID', 'ZIP 条目越界');
            }
            if ((flags & 0x1) !== 0) {
                throw new PreviewError('OFFICE_PREVIEW_ENCRYPTED', '文档已加密');
            }
            if (method !== 0 && method !== 8) {
                throw new PreviewError('OFFICE_PREVIEW_ARCHIVE_INVALID',
                    'ZIP 使用了不支持的压缩算法');
            }
            var name = normalizeName(decoder.decode(
                bytes.subarray(offset + 46, offset + 46 + nameLength)));
            if (names.has(name)) {
                throw new PreviewError('OFFICE_PREVIEW_ARCHIVE_INVALID', 'ZIP 条目重复');
            }
            names.add(name);
            if (localOffset + 30 > centralOffset) {
                throw new PreviewError('OFFICE_PREVIEW_ARCHIVE_INVALID', 'ZIP 本地头越界');
            }
            totalUncompressed += uncompressedSize;
            if (totalUncompressed > limits.maxUncompressedBytes) {
                throw new PreviewError('OFFICE_PREVIEW_ARCHIVE_TOO_LARGE',
                    'ZIP 解压总量超过限制');
            }
            if (compressedSize > 0 && uncompressedSize > 1048576
                    && uncompressedSize / compressedSize > limits.maxCompressionRatio) {
                throw new PreviewError('OFFICE_PREVIEW_ARCHIVE_TOO_LARGE',
                    'ZIP 压缩比异常');
            }
            entries.push({
                index: index,
                name: name,
                directory: name.charAt(name.length - 1) === '/',
                method: method,
                compressedSize: compressedSize,
                uncompressedSize: uncompressedSize,
                localOffset: localOffset
            });
            offset += 46 + nameLength + extraLength + commentLength;
        }

        var kind = 'unknown';
        if (names.has('word/document.xml')) {
            kind = 'docx';
        } else if (names.has('xl/workbook.xml')) {
            kind = 'spreadsheet';
        } else if (names.has('ppt/presentation.xml')) {
            kind = 'pptx';
        }
        var macro = false;
        names.forEach(function (name) {
            if (/vbaProject\.bin$/i.test(name)
                    || /^xl\/macrosheets\//i.test(name)
                    || /^ppt\/embeddings\//i.test(name)
                    || /^word\/embeddings\/oleObject/i.test(name)) {
                macro = true;
            }
        });
        return {
            kind: kind,
            macro: macro,
            entries: entries,
            names: names,
            totalUncompressed: totalUncompressed
        };
    }

    function unzipSelected(arrayBuffer, inspection, shouldInclude, maxSelectedBytes) {
        var selected = new Set();
        inspection.entries.forEach(function (entry) {
            if (!entry.directory && shouldInclude(entry.name)) {
                selected.add(entry.name);
            }
        });
        if (selected.size === 0) {
            return {};
        }
        if (typeof global.fflate === 'undefined') {
            throw new PreviewError('OFFICE_PREVIEW_PARSE_FAILED', 'ZIP 解包组件未加载');
        }
        var files = global.fflate.unzipSync(new Uint8Array(arrayBuffer), {
            filter: function (file) {
                return selected.has(file.name);
            }
        });
        var total = 0;
        Object.keys(files).forEach(function (name) {
            total += files[name].length;
        });
        if (total > maxSelectedBytes) {
            throw new PreviewError('OFFICE_PREVIEW_ARCHIVE_TOO_LARGE',
                '选定内容超过解析限制');
        }
        return files;
    }

    function xmlText(files, name, limits) {
        var bytes = files[name];
        if (!bytes) {
            throw new PreviewError('OFFICE_PREVIEW_ARCHIVE_INVALID',
                '缺少必要的文档结构文件');
        }
        if (bytes.length > limits.maxXmlBytes) {
            throw new PreviewError('OFFICE_PREVIEW_ARCHIVE_TOO_LARGE', 'XML 文件过大');
        }
        var text = decoder.decode(bytes);
        if (/<!doctype/i.test(text) || /<!entity/i.test(text)) {
            throw new PreviewError('OFFICE_PREVIEW_ARCHIVE_INVALID',
                'XML 包含不允许的 DTD 或实体声明');
        }
        return text;
    }

    function validateXmlFiles(files, limits) {
        var warnings = [];
        var externalResource = false;
        Object.keys(files).forEach(function (name) {
            if (!XML_RE.test(name)) {
                return;
            }
            var text = xmlText(files, name, limits);
            var relationship = /<(?:[A-Za-z_][\w.-]*:)?Relationship\b[^>]*>/gi;
            var match;
            while ((match = relationship.exec(text))) {
                if (!/TargetMode\s*=\s*["']External["']/i.test(match[0])) {
                    continue;
                }
                var typeMatch = match[0].match(
                    /\bType\s*=\s*(?:"([^"]*)"|'([^']*)')/i);
                var type = typeMatch
                    ? (typeMatch[1] !== undefined
                        ? typeMatch[1] : typeMatch[2]) : '';
                if (!/\/hyperlink$/i.test(type.trim())) {
                    externalResource = true;
                    break;
                }
            }
        });
        if (externalResource) {
            warnings.push('文档包含外部关系，相关资源不会加载。');
        }
        return warnings;
    }

    global.OfficeZipSecurity = {
        PreviewError: PreviewError,
        inspect: inspect,
        unzipSelected: unzipSelected,
        xmlText: xmlText,
        validateXmlFiles: validateXmlFiles
    };
})(self);
