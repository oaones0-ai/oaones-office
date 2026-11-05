(function (global) {
    'use strict';

    var Zip = global.OfficeZipSecurity;
    var Xml = global.OfficeXmlLite;

    function inspect(buffer, extension, limits) {
        if (typeof buffer === 'undefined' || !buffer.byteLength) {
            throw new Zip.PreviewError('OFFICE_PREVIEW_PARSE_FAILED', 'Word 文件为空');
        }
        var header = new Uint8Array(buffer, 0, Math.min(buffer.byteLength, 4));
        if (!(header[0] === 0x50 && header[1] === 0x4b)) {
            throw new Zip.PreviewError('OFFICE_PREVIEW_ARCHIVE_INVALID',
                'Word 文件不是有效的 OOXML 容器');
        }
        var inspection = Zip.inspect(buffer, limits);
        if (inspection.macro) {
            throw new Zip.PreviewError('OFFICE_PREVIEW_MACRO_BLOCKED', '宏内容已阻止');
        }
        if (inspection.kind !== 'docx') {
            throw new Zip.PreviewError('OFFICE_PREVIEW_ARCHIVE_INVALID',
                '文件内容与 Word 扩展名不匹配');
        }
        var selected = Zip.unzipSelected(buffer, inspection, function (name) {
            return /\.(xml|rels)$/i.test(name);
        }, limits.maxUncompressedBytes);
        var warnings = Zip.validateXmlFiles(selected, limits);
        var documentXml = Zip.xmlText(selected, 'word/document.xml', limits);
        var root = Xml.parse(documentXml);
        var breaks = Xml.descendants(root, 'br').filter(function (node) {
            return Xml.attr(node, 'type', '') === 'page';
        }).length;
        var pageHint = Math.max(1, breaks + 1);
        if (pageHint > limits.maxPageCount) {
            throw new Zip.PreviewError('OFFICE_PREVIEW_ARCHIVE_TOO_LARGE',
                'Word 页面数量超过限制');
        }
        var paragraphs = Xml.descendants(root, 'p').length;
        var tables = Xml.descendants(root, 'tbl').length;
        return {
            ok: true,
            format: 'docx',
            viewer: 'docx',
            supportLevel: extension === 'dotx' ? 'best_effort' : 'standard',
            model: {
                format: 'docx',
                pageCountHint: pageHint,
                paragraphCount: paragraphs,
                tableCount: tables
            },
            warnings: warnings
        };
    }

    global.OfficeDocxAdapter = {
        inspect: inspect
    };
})(self);
