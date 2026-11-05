(function (global) {
    'use strict';

    importScripts(
        '../vendor/fflate.min.js',
        '../vendor/xlsx.full.min.js',
        '../parser/zip-security.js',
        '../parser/xml-lite.js',
        '../parser/spreadsheet-ooxml.js',
        '../parser/docx-adapter.js',
        '../parser/spreadsheet-adapter.js',
        '../parser/pptx-adapter.js'
    );

    var handlers = {
        docx: function (buffer, extension, limits) {
            return global.OfficeDocxAdapter.inspect(buffer, extension, limits);
        },
        spreadsheet: function (buffer, extension, limits) {
            return global.OfficeSpreadsheetAdapter.parse(buffer, extension, limits);
        },
        pptx: function (buffer, extension, limits) {
            return global.OfficePptxAdapter.parse(buffer, extension, limits);
        }
    };

    global.onmessage = function (event) {
        var message = event.data || {};
        if (message.type !== 'parse') {
            return;
        }
        var requestId = message.requestId;
        try {
            var handler = handlers[message.viewer];
            if (!handler) {
                throw new Error('OFFICE_PREVIEW_UNSUPPORTED');
            }
            if (!message.buffer || message.buffer.byteLength > message.limits.maxInputBytes) {
                throw new Error('OFFICE_PREVIEW_FILE_TOO_LARGE');
            }
            var result = handler(message.buffer, message.extension, message.limits);
            var transfer = [];
            if (result && result.assets) {
                var seen = new Set();
                result.assets.forEach(function (asset) {
                    if (asset && asset.bytes && asset.bytes.buffer
                            && !seen.has(asset.bytes.buffer)) {
                        seen.add(asset.bytes.buffer);
                        transfer.push(asset.bytes.buffer);
                    }
                });
            }
            global.postMessage({
                type: 'parse-result',
                requestId: requestId,
                ok: true,
                result: result
            }, transfer);
        } catch (error) {
            global.postMessage({
                type: 'parse-result',
                requestId: requestId,
                ok: false,
                errorCode: safeErrorCode(error),
                message: publicMessage(error)
            });
        }
    };

    function safeErrorCode(error) {
        if (error && error.code && /^OFFICE_PREVIEW_[A-Z_]+$/.test(error.code)) {
            return error.code;
        }
        if (error && /^OFFICE_PREVIEW_[A-Z_]+$/.test(error.message || '')) {
            return error.message;
        }
        return 'OFFICE_PREVIEW_PARSE_FAILED';
    }

    function publicMessage(error) {
        if (error && error.code && error.message && error.message !== error.code) {
            return error.message;
        }
        var code = safeErrorCode(error);
        var messages = {
            OFFICE_PREVIEW_FILE_TOO_LARGE: '文件过大，暂不支持在线预览',
            OFFICE_PREVIEW_UNSUPPORTED: '暂不支持预览此文件格式',
            OFFICE_PREVIEW_ENCRYPTED: '文档已加密，请下载后输入密码打开',
            OFFICE_PREVIEW_MACRO_BLOCKED: '出于安全原因，宏格式文档已被禁止预览',
            OFFICE_PREVIEW_ARCHIVE_INVALID: '文档结构异常，无法安全预览',
            OFFICE_PREVIEW_ARCHIVE_TOO_LARGE: '文档解压后内容过大，暂不支持预览',
            OFFICE_PREVIEW_PARSE_FAILED: '文档解析失败，请下载后查看'
        };
        return messages[code] || messages.OFFICE_PREVIEW_PARSE_FAILED;
    }
})(self);
