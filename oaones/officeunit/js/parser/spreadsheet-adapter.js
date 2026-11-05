(function (global) {
    'use strict';

    var Zip = global.OfficeZipSecurity;

    function decodeText(buffer) {
        var bytes = new Uint8Array(buffer);
        try {
            return new TextDecoder('utf-8', {fatal: true}).decode(bytes);
        } catch (ignoreUtf8) {
            try {
                return new TextDecoder('gb18030', {fatal: true}).decode(bytes);
            } catch (ignoreGb) {
                throw new Zip.PreviewError('OFFICE_PREVIEW_PARSE_FAILED',
                    '表格文本编码无法识别');
            }
        }
    }

    function parse(buffer, extension, limits) {
        if (typeof global.XLSX === 'undefined') {
            throw new Zip.PreviewError('OFFICE_PREVIEW_PARSE_FAILED',
                '表格解析组件未加载');
        }
        var bytes = new Uint8Array(buffer);
        var zip = bytes.length >= 2 && bytes[0] === 0x50 && bytes[1] === 0x4b;
        var inspection = null;
        var ooxml = null;
        var warnings = [];
        if (zip) {
            inspection = Zip.inspect(buffer, limits);
            if (inspection.macro) {
                throw new Zip.PreviewError('OFFICE_PREVIEW_MACRO_BLOCKED', '宏内容已阻止');
            }
            if (inspection.kind !== 'unknown'
                    && inspection.kind !== 'spreadsheet') {
                throw new Zip.PreviewError('OFFICE_PREVIEW_ARCHIVE_INVALID',
                    '文件内容与表格扩展名不匹配');
            }
            var xmlFiles = Zip.unzipSelected(buffer, inspection, function (name) {
                return /\.(xml|rels)$/i.test(name);
            }, limits.maxUncompressedBytes);
            warnings = warnings.concat(Zip.validateXmlFiles(xmlFiles, limits));
            if (global.OfficeSpreadsheetOoxml) {
                try {
                    ooxml = global.OfficeSpreadsheetOoxml.parse(xmlFiles, limits);
                } catch (ignore) {
                    warnings.push('部分 Excel 样式未能读取，已使用基础表格样式。');
                }
            }
        }

        var workbook;
        try {
            if (extension === 'csv' || extension === 'tsv') {
                workbook = global.XLSX.read(decodeText(buffer), {
                    type: 'string',
                    FS: extension === 'tsv' ? '\t' : undefined,
                    cellDates: true,
                    cellNF: true,
                    cellStyles: true
                });
            } else {
                workbook = global.XLSX.read(bytes, {
                    type: 'array',
                    cellDates: true,
                    cellNF: true,
                    cellStyles: true,
                    bookVBA: false,
                    WTF: false
                });
            }
        } catch (error) {
            throw new Zip.PreviewError('OFFICE_PREVIEW_PARSE_FAILED',
                '表格文件解析失败');
        }
        if (!workbook || !workbook.SheetNames) {
            throw new Zip.PreviewError('OFFICE_PREVIEW_PARSE_FAILED',
                '表格文件没有可用工作簿');
        }
        if (workbook.SheetNames.length > limits.maxSheetCount) {
            throw new Zip.PreviewError('OFFICE_PREVIEW_ARCHIVE_TOO_LARGE',
                '工作表数量超过限制');
        }

        var result = {
            format: 'spreadsheet',
            metadata: {
                title: workbook.Props && workbook.Props.Title
                    ? String(workbook.Props.Title) : ''
            },
            sheets: [],
            warnings: warnings
        };
        workbook.SheetNames.forEach(function (name, index) {
            var sheet = workbook.Sheets[name];
            if (!sheet || (sheet.Hidden && Number(sheet.Hidden) !== 0)) {
                return;
            }
            result.sheets.push(parseSheet(
                sheet,
                name,
                index,
                limits,
                ooxml && ooxml.sheets[name],
                ooxml && ooxml.styles
            ));
        });
        if (!result.sheets.length) {
            throw new Zip.PreviewError('OFFICE_PREVIEW_PARSE_FAILED',
                '工作簿中没有可见工作表');
        }
        return result;
    }

    function parseSheet(sheet, name, index, limits, sheetMeta, styles) {
        var range = sheet['!ref'] ? global.XLSX.utils.decode_range(sheet['!ref']) : null;
        if (!range) {
            return {
                name: name,
                index: index,
                rowCount: 0,
                colCount: 0,
                rows: [],
                merges: [],
                cols: [],
                defaultRowHeight: 0,
                defaultColWidth: 0
            };
        }
        var rowCount = range.e.r - range.s.r + 1;
        var colCount = range.e.c - range.s.c + 1;
        var area = Math.max(1, rowCount * colCount);
        var actualCells = Object.keys(sheet).filter(function (key) {
            return key.charAt(0) !== '!';
        }).length;
        if (area > limits.maxCellCount
                || actualCells > limits.maxCellCount) {
            throw new Zip.PreviewError('OFFICE_PREVIEW_ARCHIVE_TOO_LARGE',
                '工作表单元格数量超过限制');
        }

        var rows = [];
        var rowMap = new Map();
        for (var rowIndex = range.s.r; rowIndex <= range.e.r; rowIndex += 1) {
            var rowStyle = sheet['!rows'] && sheet['!rows'][rowIndex];
            var rowMeta = sheetMeta && sheetMeta.rows
                && sheetMeta.rows[rowIndex];
            var rowHeight = rowMeta
                ? Number(rowMeta.height) || 0
                : rowHeightPx(rowStyle);
            var rowHidden = rowMeta ? !!rowMeta.hidden : !!(rowStyle && rowStyle.hidden);
            var rowModel = null;
            if (rowMeta || (rowStyle && (rowHeight > 0 || rowStyle.hidden))) {
                rowModel = {
                    index: rowIndex,
                    height: rowHidden ? 0 : rowHeight,
                    hidden: rowHidden,
                    cells: []
                };
                rowMap.set(rowIndex, rowModel);
                rows.push(rowModel);
            }
            for (var colIndex = range.s.c; colIndex <= range.e.c; colIndex += 1) {
                var address = global.XLSX.utils.encode_cell({r: rowIndex, c: colIndex});
                var cell = sheet[address];
                if (!cell) {
                    continue;
                }
                if (!rowModel) {
                    rowModel = {
                        index: rowIndex,
                        height: rowHeight,
                        hidden: false,
                        cells: []
                    };
                    rowMap.set(rowIndex, rowModel);
                    rows.push(rowModel);
                }
                var cellMeta = sheetMeta && sheetMeta.cells
                    && sheetMeta.cells[address];
                rowModel.cells.push(createCell(
                    cell,
                    address,
                    rowIndex,
                    colIndex,
                    cellMeta,
                    styles
                ));
            }
        }

        var merges = (sheet['!merges'] || []).map(function (merge) {
            return {
                s: {r: merge.s.r, c: merge.s.c},
                e: {r: merge.e.r, c: merge.e.c},
                ref: global.XLSX.utils.encode_range(merge)
            };
        });
        var cols = [];
        for (var column = range.s.c; column <= range.e.c; column += 1) {
            var metaColumn = sheetMeta && sheetMeta.columns
                && sheetMeta.columns[column];
            var legacyColumn = sheet['!cols'] && sheet['!cols'][column];
            cols.push(metaColumn || {
                width: legacyColumn && legacyColumn.wpx
                    ? legacyColumn.wpx
                    : legacyColumn && legacyColumn.wch
                        ? Math.floor((legacyColumn.wch * 256 + Math.floor(128 / 7))
                            / 256 * 7)
                        : 0,
                hidden: !!(legacyColumn && legacyColumn.hidden)
            });
        }
        return {
            name: name,
            index: index,
            rowCount: rowCount,
            colCount: colCount,
            startRow: range.s.r,
            startCol: range.s.c,
            rows: rows,
            merges: merges,
            cols: cols,
            freeze: sheet['!freeze'] || null,
            defaultRowHeight: sheetMeta
                ? sheetMeta.defaultRowHeight : 0,
            defaultColWidth: sheetMeta
                ? sheetMeta.defaultColumnWidth : 0
        };
    }

    function rowHeightPx(rowStyle) {
        if (!rowStyle) {
            return 0;
        }
        var height = Number(rowStyle.hpx);
        if (isFinite(height) && height > 0) {
            return height;
        }
        height = Number(rowStyle.hpt);
        return isFinite(height) && height > 0 ? height * 96 / 72 : 0;
    }

    function createCell(cell, address, row, col, cellMeta, styles) {
        var style = cellMeta && isFinite(cellMeta.styleId) && styles
            ? cloneStyle(styles[cellMeta.styleId])
            : null;
        if (!style) {
            style = extractStyle(cell.s);
        }
        var value = cell.v;
        if (value instanceof Date) {
            value = value.toISOString();
        } else if (value !== null && typeof value === 'object') {
            value = '';
        }
        return {
            address: address,
            row: row,
            col: col,
            type: cell.t || 'z',
            value: value === undefined ? '' : value,
            text: cell.w !== undefined ? String(cell.w)
                : value === undefined || value === null ? '' : String(value),
            formula: cell.f ? 'cached' : '',
            style: style,
            comment: cellMeta && cellMeta.comment ? cellMeta.comment : ''
        };
    }

    function cloneStyle(style) {
        if (!style) {
            return null;
        }
        var result = {};
        if (style.font) {
            result.font = {
                name: style.font.name,
                size: style.font.size,
                bold: !!style.font.bold,
                italic: !!style.font.italic,
                underline: !!style.font.underline,
                strike: !!style.font.strike,
                color: style.font.color
            };
        }
        if (style.fill) {
            result.fill = style.fill;
        }
        if (style.alignment) {
            result.align = style.alignment.horizontal || '';
            result.valign = style.alignment.vertical || '';
            result.wrap = !!style.alignment.wrapText;
            result.indent = style.alignment.indent || 0;
            result.textRotation = style.alignment.textRotation || 0;
            result.shrinkToFit = !!style.alignment.shrinkToFit;
        }
        if (style.border && Object.keys(style.border).length) {
            result.border = {
                top: cloneBorder(style.border.top),
                right: cloneBorder(style.border.right),
                bottom: cloneBorder(style.border.bottom),
                left: cloneBorder(style.border.left),
                diagonal: cloneBorder(style.border.diagonal)
            };
        }
        if (style.numFmt) {
            result.numFmt = style.numFmt;
        }
        return result;
    }

    function cloneBorder(border) {
        return border ? {
            color: border.color || '',
            style: border.style || ''
        } : null;
    }

    function extractStyle(style) {
        if (!style || typeof style !== 'object') {
            return null;
        }
        var result = {};
        if (style.font) {
            result.font = {
                name: style.font.name,
                size: style.font.sz,
                bold: !!style.font.bold,
                italic: !!style.font.italic,
                underline: !!style.font.underline,
                color: rgb(style.font.color)
            };
        }
        if (style.fill) {
            result.fill = rgb(style.fill);
        } else if (style.patternType === 'solid') {
            result.fill = rgb(style.fgColor);
        }
        if (style.alignment) {
            result.align = style.alignment.horizontal || '';
            result.valign = style.alignment.vertical || '';
            result.wrap = !!style.alignment.wrapText;
        }
        if (style.border) {
            result.border = {
                top: border(style.border.top),
                right: border(style.border.right),
                bottom: border(style.border.bottom),
                left: border(style.border.left)
            };
        }
        if (style.numFmt) {
            result.numFmt = style.numFmt;
        }
        return result;
    }

    function border(value) {
        if (!value) {
            return null;
        }
        return {
            color: rgb(value.color),
            style: value.style || ''
        };
    }

    function rgb(value) {
        if (!value) {
            return '';
        }
        if (value.rgb) {
            return '#' + String(value.rgb).replace(/^FF/i, '').slice(-6);
        }
        if (value.theme !== undefined) {
            return '';
        }
        return '';
    }

    global.OfficeSpreadsheetAdapter = {
        parse: parse
    };
})(self);
