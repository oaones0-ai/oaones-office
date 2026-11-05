(function (global) {
    'use strict';

    var DEFAULT_ROW_HEIGHT = 19;
    var DEFAULT_COL_WIDTH = 64;
    var ROW_HEADER_WIDTH = 46;
    var COL_HEADER_HEIGHT = 24;

    function render(options) {
        var container = options.container;
        var destroyed = options.destroyed;
        var model = options.model;
        container.replaceChildren();
        var root = element('div', 'oa-spreadsheet');
        var tabs = element('div', 'oa-sheet-tabs');
        tabs.setAttribute('role', 'tablist');
        var sheetViewport = element('div', 'oa-sheet-viewport');
        sheetViewport.tabIndex = 0;
        sheetViewport.style.overflow = 'auto';
        sheetViewport.style.scrollbarGutter = 'stable';
        var canvas = element('div', 'oa-sheet-canvas');
        sheetViewport.appendChild(canvas);
        root.appendChild(sheetViewport);
        root.appendChild(tabs);
        container.appendChild(root);

        var activeIndex = 0;
        var scale = 1;
        var rendered = null;
        var resizeObserver = null;
        var frame = 0;
        var activeCellAddress = '';

        model.sheets.forEach(function (sheet, index) {
            var button = element('button', 'oa-sheet-tab', sheet.name);
            button.type = 'button';
            button.setAttribute('role', 'tab');
            button.setAttribute('data-oa-sheet', 'sheet-' + (index + 1));
            button.addEventListener('click', function () {
                activate(index, true);
            });
            tabs.appendChild(button);
        });

        function activate(index, scrollToTop) {
            if (index < 0 || index >= model.sheets.length) {
                return;
            }
            activeIndex = index;
            activeCellAddress = '';
            sheetViewport.setAttribute('data-oa-sheet',
                'sheet-' + (activeIndex + 1));
            Array.prototype.forEach.call(tabs.children, function (button, buttonIndex) {
                var active = buttonIndex === index;
                button.classList.toggle('active', active);
                button.setAttribute('aria-selected', active ? 'true' : 'false');
            });
            var sheet = model.sheets[index];
            rendered = prepareSheet(sheet);
            if (scrollToTop) {
                sheetViewport.scrollTop = 0;
                sheetViewport.scrollLeft = 0;
            }
            renderVisible();
            if (typeof options.onSectionChange === 'function' && !destroyed()) {
                options.onSectionChange(index);
            }
        }

        function prepareSheet(sheet) {
            var rowOffsets = new Float64Array(sheet.rowCount + 1);
            var rowByIndex = new Map();
            (sheet.rows || []).forEach(function (row) {
                rowByIndex.set(row.index - sheet.startRow, row);
            });
            for (var row = 0; row < sheet.rowCount; row += 1) {
                var rowModel = rowByIndex.get(row);
                rowOffsets[row + 1] = rowOffsets[row]
                    + (rowModel
                        ? rowModel.hidden ? 0 : rowModel.height || defaultRowHeight(sheet)
                        : defaultRowHeight(sheet));
            }
            var colOffsets = new Float64Array(sheet.colCount + 1);
            for (var col = 0; col < sheet.colCount; col += 1) {
                var colModel = sheet.cols && sheet.cols[col];
                var width = colModel
                    ? colModel.hidden ? 0 : colModel.width || defaultColWidth(sheet)
                    : defaultColWidth(sheet);
                colOffsets[col + 1] = colOffsets[col] + width;
            }
            var mergeStarts = new Map();
            var mergeCovers = new Map();
            (sheet.merges || []).forEach(function (merge) {
                var rowStart = merge.s.r - sheet.startRow;
                var rowEnd = merge.e.r - sheet.startRow;
                var colStart = merge.s.c - sheet.startCol;
                var colEnd = merge.e.c - sheet.startCol;
                mergeStarts.set(key(rowStart, colStart), {
                    rowStart: rowStart,
                    rowEnd: rowEnd,
                    colStart: colStart,
                    colEnd: colEnd,
                    ref: merge.ref
                });
                for (var mRow = rowStart; mRow <= rowEnd; mRow += 1) {
                    for (var mCol = colStart; mCol <= colEnd; mCol += 1) {
                        if (mRow !== rowStart || mCol !== colStart) {
                            mergeCovers.set(key(mRow, mCol), true);
                        }
                    }
                }
            });
            var centerContinuous = centerContinuousRanges(
                sheet,
                rowByIndex,
                sheet.rowCount,
                sheet.colCount
            );
            return {
                sheet: sheet,
                rowOffsets: rowOffsets,
                colOffsets: colOffsets,
                rowByIndex: rowByIndex,
                mergeStarts: mergeStarts,
                mergeCovers: mergeCovers,
                centerStarts: centerContinuous.starts,
                centerCovers: centerContinuous.covers,
                width: colOffsets[colOffsets.length - 1] || 0,
                height: rowOffsets[rowOffsets.length - 1] || 0
            };
        }

        function renderVisible() {
            if (!rendered || destroyed()) {
                return;
            }
            var sheet = rendered.sheet;
            var baseWidth = Math.max(sheetViewport.clientWidth, rendered.width + ROW_HEADER_WIDTH);
            var baseHeight = rendered.height + COL_HEADER_HEIGHT;
            canvas.style.width = baseWidth + 'px';
            canvas.style.height = baseHeight + 'px';
            canvas.style.zoom = String(scale);
            var left = sheetViewport.scrollLeft / scale;
            var top = Math.max(0, sheetViewport.scrollTop / scale - COL_HEADER_HEIGHT);
            var right = left + sheetViewport.clientWidth / scale;
            var bottom = top + sheetViewport.clientHeight / scale;
            var rowStart = Math.max(0, findOffset(rendered.rowOffsets, top) - 5);
            var rowEnd = Math.min(sheet.rowCount - 1,
                findOffset(rendered.rowOffsets, bottom) + 7);
            var colStart = Math.max(0, findOffset(rendered.colOffsets, left) - 2);
            var colEnd = Math.min(sheet.colCount - 1,
                findOffset(rendered.colOffsets, right) + 3);

            canvas.replaceChildren();
            canvas.appendChild(renderColumnHeaders(colStart, colEnd));
            canvas.appendChild(renderRows(rowStart, rowEnd, colStart, colEnd));
            if (activeCellAddress) {
                var active = canvas.querySelector(
                    '[data-oa-cell="' + activeCellAddress + '"]'
                );
                if (active) {
                    active.classList.add('active');
                }
            }
        }

        function renderColumnHeaders(colStart, colEnd) {
            var header = element('div', 'oa-sheet-column-headers');
            header.style.height = COL_HEADER_HEIGHT + 'px';
            var corner = element('span', 'oa-sheet-corner', '#');
            corner.style.width = ROW_HEADER_WIDTH + 'px';
            header.appendChild(corner);
            for (var col = colStart; col <= colEnd; col += 1) {
                var item = element('span', 'oa-sheet-column-header',
                    columnName(col + rendered.sheet.startCol));
                item.style.left = (ROW_HEADER_WIDTH + rendered.colOffsets[col]) + 'px';
                item.style.width = (rendered.colOffsets[col + 1]
                    - rendered.colOffsets[col]) + 'px';
                header.appendChild(item);
            }
            return header;
        }

        function renderRows(rowStart, rowEnd, colStart, colEnd) {
            var layer = element('div', 'oa-sheet-rows');
            var sheet = rendered.sheet;
            for (var row = rowStart; row <= rowEnd; row += 1) {
                if (row < 0 || row >= sheet.rowCount) {
                    continue;
                }
                var rowHeight = rendered.rowOffsets[row + 1] - rendered.rowOffsets[row];
                var rowElement = element('div', 'oa-sheet-row');
                rowElement.style.top = (COL_HEADER_HEIGHT + rendered.rowOffsets[row]) + 'px';
                rowElement.style.height = rowHeight + 'px';
                var excelRowNumber = row + sheet.startRow + 1;
                rowElement.setAttribute('data-oa-order', String(excelRowNumber));
                var rowNumber = element('span', 'oa-sheet-row-number',
                    String(excelRowNumber));
                rowNumber.style.width = ROW_HEADER_WIDTH + 'px';
                rowNumber.title = '第 ' + excelRowNumber + ' 行';
                rowElement.appendChild(rowNumber);
                var rowModel = rendered.rowByIndex.get(row);
                var cellMap = new Map();
                if (rowModel) {
                    rowModel.cells.forEach(function (cell) {
                        cellMap.set(cell.col - sheet.startCol, cell);
                    });
                }
                for (var col = colStart; col <= colEnd; col += 1) {
                    if (rendered.mergeCovers.get(key(row, col))) {
                        continue;
                    }
                    if (rendered.centerCovers.get(key(row, col))) {
                        continue;
                    }
                    var cell = cellMap.get(col);
                    var merge = rendered.mergeStarts.get(key(row, col));
                    var center = rendered.centerStarts.get(key(row, col));
                    var cellWidth = rendered.colOffsets[col + 1] - rendered.colOffsets[col];
                    var cellHeight = rowHeight;
                    var range = '';
                    if (center) {
                        cellWidth = rendered.colOffsets[center.colEnd + 1]
                            - rendered.colOffsets[col];
                        range = center.ref;
                    } else if (merge) {
                        var maxCol = Math.min(merge.colEnd, colEnd);
                        var maxRow = Math.min(merge.rowEnd, rowEnd);
                        cellWidth = rendered.colOffsets[maxCol + 1] - rendered.colOffsets[col];
                        cellHeight = rendered.rowOffsets[maxRow + 1] - rendered.rowOffsets[row];
                        range = merge.ref;
                    }
                    var cellElement = renderCell(cell, row, col, sheet);
                    cellElement.style.left = (ROW_HEADER_WIDTH
                        + rendered.colOffsets[col]) + 'px';
                    cellElement.style.width = cellWidth + 'px';
                    cellElement.style.height = cellHeight + 'px';
                    if (range) {
                        cellElement.setAttribute('data-oa-range', range);
                    }
                    if (center) {
                        cellElement.setAttribute('data-oa-center', range);
                    }
                    rowElement.appendChild(cellElement);
                }
                layer.appendChild(rowElement);
            }
            return layer;
        }

        function renderCell(cell, row, col, sheet) {
            var address = columnName(col + sheet.startCol) + (row + sheet.startRow + 1);
            var value = cell ? cell.text : '';
            var item = element('span', 'oa-sheet-cell', value);
            item.setAttribute('data-oa-cell', address);
            item.setAttribute('data-oa-role', 'table-cell');
            item.setAttribute('data-oa-order',
                String(col + sheet.startCol + 1));
            item.tabIndex = -1;
            item.title = value ? address + '  ' + value : address;
            if (cell && cell.comment) {
                item.classList.add('has-comment');
                item.title += '\n批注：' + cell.comment;
            }
            if (!cell || !cell.style) {
                return item;
            }
            var style = cell.style;
            if (style.fill) {
                item.style.background = style.fill;
            }
            if (style.font) {
                if (style.font.name) {
                    item.style.fontFamily = style.font.name;
                }
                if (style.font.size) {
                    item.style.fontSize = (style.font.size * 96 / 72) + 'px';
                }
                if (style.font.bold) {
                    item.style.fontWeight = '700';
                }
                if (style.font.italic) {
                    item.style.fontStyle = 'italic';
                }
                if (style.font.underline) {
                    item.style.textDecoration = 'underline';
                }
                if (style.font.color) {
                    item.style.color = style.font.color;
                }
            }
            if (style.align) {
                item.style.justifyContent = style.align === 'center' ? 'center'
                    : style.align === 'right' ? 'flex-end'
                        : style.align === 'centerContinuous' ? 'center'
                            : style.align === 'justify' ? 'space-between'
                                : 'flex-start';
                item.style.textAlign = style.align === 'centerContinuous'
                    ? 'center' : style.align;
            }
            if (style.valign) {
                item.style.alignItems = style.valign === 'center' ? 'center'
                    : style.valign === 'bottom' ? 'flex-end'
                        : style.valign === 'justify' ? 'space-between'
                            : 'flex-start';
            }
            if (style.wrap) {
                item.style.whiteSpace = 'pre-wrap';
                item.style.overflowWrap = 'break-word';
            }
            if (style.indent) {
                item.style.paddingLeft = (5 + style.indent * 7) + 'px';
            }
            if (style.textRotation) {
                item.style.rotate = style.textRotation + 'deg';
            }
            if (style.border) {
                applyBorder(item, style.border);
            }
            if (cell.formula) {
                item.title += '（显示缓存计算结果）';
            }
            return item;
        }

        function applyBorder(item, border) {
            [
                ['top', border.top],
                ['right', border.right],
                ['bottom', border.bottom],
                ['left', border.left]
            ].forEach(function (entry) {
                if (!entry[1]) {
                    return;
                }
                var css = borderStyle(entry[1].style);
                var width = borderWidth(entry[1].style);
                item.style['border' + entry[0].charAt(0).toUpperCase()
                    + entry[0].substring(1)] = width + ' '
                    + css + ' ' + (entry[1].color || '#000000');
            });
        }

        function borderStyle(value) {
            var text = String(value || '');
            if (/double/i.test(text)) {
                return 'double';
            }
            if (/dash/i.test(text)) {
                return 'dashed';
            }
            if (/dot/i.test(text)) {
                return 'dotted';
            }
            return 'solid';
        }

        function borderWidth(value) {
            var text = String(value || '');
            if (/thick/i.test(text)) {
                return '3px';
            }
            if (/medium/i.test(text)) {
                return '2px';
            }
            return '1px';
        }

        function onScroll() {
            if (frame) {
                return;
            }
            frame = global.requestAnimationFrame(function () {
                frame = 0;
                renderVisible();
            });
        }

        sheetViewport.addEventListener('scroll', onScroll, {passive: true});
        resizeObserver = new ResizeObserver(function () {
            renderVisible();
        });
        resizeObserver.observe(sheetViewport);
        activate(0, false);

        function focusCell(address) {
            var match = /^([A-Z]+)(\d+)$/i.exec(String(address || ''));
            if (!match || !rendered) {
                return false;
            }
            var row = parseInt(match[2], 10) - 1 - rendered.sheet.startRow;
            var col = columnNumber(match[1]) - 1 - rendered.sheet.startCol;
            if (row < 0 || row >= rendered.sheet.rowCount
                    || col < 0 || col >= rendered.sheet.colCount) {
                return false;
            }
            var rowHeight = rendered.rowOffsets[row + 1] - rendered.rowOffsets[row];
            var colWidth = rendered.colOffsets[col + 1] - rendered.colOffsets[col];
            var addressText = match[1].toUpperCase() + match[2];
            activeCellAddress = addressText;
            sheetViewport.scrollTop = Math.max(0, COL_HEADER_HEIGHT
                + rendered.rowOffsets[row] + rowHeight / 2
                - sheetViewport.clientHeight / 2);
            sheetViewport.scrollLeft = Math.max(0, ROW_HEADER_WIDTH
                + rendered.colOffsets[col] + colWidth / 2
                - sheetViewport.clientWidth / 2);
            renderVisible();
            global.requestAnimationFrame(function () {
                var cell = sheetViewport.querySelector(
                    '[data-oa-cell="' + addressText + '"]'
                );
                if (!cell) {
                    return;
                }
                Array.prototype.forEach.call(
                    sheetViewport.querySelectorAll('.oa-sheet-cell.active'),
                    function (item) {
                        item.classList.remove('active');
                    }
                );
                cell.classList.add('active');
                cell.focus({preventScroll: true});
            });
            return true;
        }

        return {
            sections: model.sheets.map(function (sheet, index) {
                return {
                    id: 'sheet-' + (index + 1),
                    label: sheet.name,
                    kind: 'sheet',
                    index: index
                };
            }),
            setZoom: function (value) {
                scale = Math.max(0.4, Math.min(2.5, value || 1));
                renderVisible();
                return scale;
            },
            fitWidth: function () {
                if (!rendered) {
                    return scale;
                }
                scale = Math.max(0.4, Math.min(1.35,
                    (sheetViewport.clientWidth - 8)
                    / Math.max(1, rendered.width + ROW_HEADER_WIDTH)));
                renderVisible();
                return scale;
            },
            goTo: function (anchor) {
                var match = /^sheet-(\d+)(?:\/cell-([A-Z]+\d+))?$/i.exec(
                    String(anchor || '')
                );
                if (!match) {
                    return false;
                }
                var sheetIndex = parseInt(match[1], 10) - 1;
                if (sheetIndex < 0 || sheetIndex >= model.sheets.length) {
                    return false;
                }
                activate(sheetIndex, true);
                if (match[2]) {
                    global.requestAnimationFrame(function () {
                        focusCell(match[2]);
                    });
                }
                return true;
            },
            getCurrentIndex: function () {
                return activeIndex;
            },
            getCurrentAnchor: function () {
                return 'sheet-' + (activeIndex + 1);
            },
            destroy: function () {
                if (frame) {
                    global.cancelAnimationFrame(frame);
                }
                if (resizeObserver) {
                    resizeObserver.disconnect();
                }
                container.replaceChildren();
            }
        };
    }

    function findOffset(offsets, value) {
        var low = 0;
        var high = offsets.length - 1;
        while (low < high) {
            var middle = Math.floor((low + high) / 2);
            if (offsets[middle + 1] <= value) {
                low = middle + 1;
            } else {
                high = middle;
            }
        }
        return Math.max(0, Math.min(offsets.length - 2, low));
    }

    function columnName(index) {
        var name = '';
        var value = index + 1;
        while (value > 0) {
            var remainder = (value - 1) % 26;
            name = String.fromCharCode(65 + remainder) + name;
            value = Math.floor((value - 1) / 26);
        }
        return name;
    }

    function columnNumber(value) {
        var text = String(value || '').toUpperCase();
        var number = 0;
        for (var i = 0; i < text.length; i += 1) {
            number = number * 26 + text.charCodeAt(i) - 64;
        }
        return number;
    }

    function key(row, col) {
        return row + ':' + col;
    }

    function defaultRowHeight(sheet) {
        var value = Number(sheet && sheet.defaultRowHeight);
        return isFinite(value) && value > 0 ? value : DEFAULT_ROW_HEIGHT;
    }

    function defaultColWidth(sheet) {
        var value = Number(sheet && sheet.defaultColWidth);
        return isFinite(value) && value > 0 ? value : DEFAULT_COL_WIDTH;
    }

    function centerContinuousRanges(sheet, rowByIndex, rowCount, colCount) {
        var starts = new Map();
        var covers = new Map();
        for (var row = 0; row < rowCount; row += 1) {
            var rowModel = rowByIndex.get(row);
            var cellMap = new Map();
            if (rowModel) {
                rowModel.cells.forEach(function (cell) {
                    cellMap.set(cell.col - sheet.startCol, cell);
                });
            }
            var groupStart = -1;
            for (var col = 0; col <= colCount; col += 1) {
                var cell = col < colCount ? cellMap.get(col) : null;
                var alignment = cell && cell.style && cell.style.align;
                var continues = alignment === 'centerContinuous';
                var closes = !continues;
                if (continues && groupStart < 0) {
                    groupStart = col;
                }
                if (continues && cell.style.border && cell.style.border.right) {
                    closes = true;
                }
                if (closes && groupStart >= 0) {
                    var end = continues ? col : col - 1;
                    if (end > groupStart) {
                        var ref = columnName(groupStart + sheet.startCol)
                            + (row + sheet.startRow + 1) + ':'
                            + columnName(end + sheet.startCol)
                            + (row + sheet.startRow + 1);
                        starts.set(key(row, groupStart), {
                            colStart: groupStart,
                            colEnd: end,
                            ref: ref
                        });
                        for (var covered = groupStart + 1;
                            covered <= end; covered += 1) {
                            covers.set(key(row, covered), true);
                        }
                    }
                    groupStart = -1;
                }
            }
        }
        return {starts: starts, covers: covers};
    }

    function element(tag, className, text) {
        var node = document.createElement(tag);
        if (className) {
            node.className = className;
        }
        if (text !== undefined) {
            node.textContent = text;
        }
        return node;
    }

    global.OfficeRenderers = global.OfficeRenderers || {};
    global.OfficeRenderers.spreadsheet = {
        render: render
    };
})(self);
