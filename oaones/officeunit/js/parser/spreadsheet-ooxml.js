(function (global) {
    'use strict';

    var Xml = global.OfficeXmlLite;
    var Zip = global.OfficeZipSecurity;
    var POINTS_TO_PIXELS = 96 / 72;
    var DEFAULT_FONT_SIZE = 11;
    var DEFAULT_COLUMN_WIDTH = 8.43;
    var DEFAULT_ROW_HEIGHT = 15;
    var MAX_DIGIT_WIDTH = 7;

    // ECMA-376 indexed colours. Unsupported custom palettes fall back to black.
    var INDEXED_COLORS = [
        '#000000', '#FFFFFF', '#FF0000', '#00FF00', '#0000FF', '#FFFF00',
        '#FF00FF', '#00FFFF', '#000000', '#FFFFFF', '#FF0000', '#00FF00',
        '#0000FF', '#FFFF00', '#FF00FF', '#00FFFF', '#800000', '#008000',
        '#000080', '#808000', '#800080', '#008080', '#C0C0C0', '#808080',
        '#9999FF', '#993366', '#FFFFCC', '#CCFFFF', '#660066', '#FF8080',
        '#0066CC', '#CCCCFF', '#000080', '#FF00FF', '#FFFF00', '#00FFFF',
        '#800080', '#800000', '#008080', '#0000FF', '#00CCFF', '#CCFFFF',
        '#CCFFCC', '#FFFF99', '#99CCFF', '#FF99CC', '#CC99FF', '#FFCC99',
        '#3366FF', '#33CCCC', '#99CC00', '#FFCC00', '#FF9900', '#FF6600',
        '#666699', '#969696', '#003366', '#339966', '#003300', '#333300',
        '#993300', '#993366', '#333399', '#333333'
    ];

    function parseContext(files, limits) {
        if (!files || !files['xl/workbook.xml'] || !Xml) {
            return null;
        }
        var themeColors = parseTheme(files, limits);
        var styles = [];
        try {
            styles = parseStyles(files, limits, themeColors);
        } catch (ignore) {
            // Basic grid rendering remains available when styles are malformed.
        }
        var relationships = parseRelationships(
            files['xl/_rels/workbook.xml.rels'], limits);
        var workbook = Xml.parse(Zip.xmlText(files, 'xl/workbook.xml', limits));
        var sheets = {};
        Xml.descendants(workbook, 'sheet').forEach(function (sheetNode, index) {
            var name = Xml.attr(sheetNode, 'name', '');
            if (!name) {
                return;
            }
            var relationshipId = Xml.attr(sheetNode, 'id', '');
            var relationship = relationships[relationshipId];
            var path = relationship && relationship.type
                && /\/worksheet$/i.test(relationship.type)
                ? resolvePart('xl/workbook.xml', relationship.target)
                : 'xl/worksheets/sheet' + (index + 1) + '.xml';
            try {
                sheets[name] = parseSheetMeta(files, path, limits, themeColors);
            } catch (ignore) {
                sheets[name] = defaultSheetMeta();
            }
        });
        return {
            styles: styles,
            sheets: sheets,
            themeColors: themeColors
        };
    }

    function parseTheme(files, limits) {
        var path = files['xl/theme/theme1.xml']
            ? 'xl/theme/theme1.xml'
            : Object.keys(files).filter(function (name) {
                return /^xl\/theme\/theme\d+\.xml$/i.test(name);
            }).sort()[0];
        if (!path) {
            return [];
        }
        var root;
        try {
            root = Xml.parse(Zip.xmlText(files, path, limits));
        } catch (ignore) {
            return [];
        }
        var scheme = Xml.firstDescendant(root, 'clrScheme');
        if (!scheme) {
            return [];
        }
        var byName = {};
        scheme.children.forEach(function (node) {
            var colorNode = node.children[0];
            if (colorNode) {
                byName[node.local] = parseDirectColor(colorNode);
            }
        });
        return [
            byName.lt1 || '#FFFFFF',
            byName.dk1 || '#000000',
            byName.lt2 || '#E7E6E6',
            byName.dk2 || '#44546A',
            byName.accent1 || '#4472C4',
            byName.accent2 || '#ED7D31',
            byName.accent3 || '#A5A5A5',
            byName.accent4 || '#FFC000',
            byName.accent5 || '#5B9BD5',
            byName.accent6 || '#70AD47',
            byName.hlink || '#0563C1',
            byName.folhlink || '#954F72'
        ];
    }

    function parseDirectColor(node) {
        var rgb = Xml.attr(node, 'val', '');
        if (/^[0-9a-f]{6}$/i.test(rgb) || /^[0-9a-f]{8}$/i.test(rgb)) {
            return normalizeHex(rgb);
        }
        var last = Xml.attr(node, 'lastClr', '');
        if (/^[0-9a-f]{6}$/i.test(last) || /^[0-9a-f]{8}$/i.test(last)) {
            return normalizeHex(last);
        }
        return '';
    }

    function parseStyles(files, limits, themeColors) {
        if (!files['xl/styles.xml']) {
            return [];
        }
        var root = Xml.parse(Zip.xmlText(files, 'xl/styles.xml', limits));
        var customFormats = {};
        var numFmts = Xml.child(root, 'numFmts');
        Xml.children(numFmts, 'numFmt').forEach(function (node) {
            customFormats[Xml.attr(node, 'numFmtId', '')] =
                Xml.attr(node, 'formatCode', '');
        });

        var fonts = Xml.children(Xml.child(root, 'fonts'), 'font').map(function (node) {
            var size = parseFloat(Xml.attr(Xml.child(node, 'sz'), 'val', ''));
            var nameNode = Xml.child(node, 'name');
            var colorNode = Xml.child(node, 'color');
            return {
                name: Xml.attr(nameNode, 'val', ''),
                size: isFinite(size) && size > 0 ? size : DEFAULT_FONT_SIZE,
                bold: !!Xml.child(node, 'b'),
                italic: !!Xml.child(node, 'i'),
                underline: !!Xml.child(node, 'u'),
                strike: !!Xml.child(node, 'strike'),
                color: parseColor(colorNode, themeColors)
            };
        });
        var fills = Xml.children(Xml.child(root, 'fills'), 'fill').map(function (node) {
            var pattern = Xml.child(node, 'patternFill');
            if (!pattern) {
                return {};
            }
            var type = Xml.attr(pattern, 'patternType', 'none');
            if (type === 'none' || type === 'gray125') {
                return {pattern: type};
            }
            return {
                pattern: type,
                fgColor: parseColor(Xml.child(pattern, 'fgColor'), themeColors),
                bgColor: parseColor(Xml.child(pattern, 'bgColor'), themeColors)
            };
        });
        var borders = Xml.children(Xml.child(root, 'borders'), 'border').map(function (node) {
            var result = {};
            ['left', 'right', 'top', 'bottom', 'diagonal'].forEach(function (name) {
                var side = Xml.child(node, name);
                if (!side || !Xml.attr(side, 'style', '')) {
                    return;
                }
                result[name] = {
                    style: Xml.attr(side, 'style', ''),
                    color: parseColor(Xml.child(side, 'color'), themeColors)
                };
            });
            return result;
        });

        return Xml.children(Xml.child(root, 'cellXfs'), 'xf').map(function (node) {
            var fontId = parseInt(Xml.attr(node, 'fontId', '0'), 10);
            var fillId = parseInt(Xml.attr(node, 'fillId', '0'), 10);
            var borderId = parseInt(Xml.attr(node, 'borderId', '0'), 10);
            var numFmtId = Xml.attr(node, 'numFmtId', '0');
            var alignment = Xml.child(node, 'alignment');
            var fill = fills[fillId] || {};
            var result = {
                font: fonts[fontId] || fonts[0] || null,
                alignment: {
                    horizontal: Xml.attr(alignment, 'horizontal', ''),
                    vertical: Xml.attr(alignment, 'vertical', ''),
                    wrapText: Xml.attr(alignment, 'wrapText', '0') === '1',
                    indent: Math.max(0, parseInt(Xml.attr(alignment, 'indent', '0'), 10) || 0),
                    textRotation: parseInt(Xml.attr(alignment, 'textRotation', '0'), 10) || 0,
                    shrinkToFit: Xml.attr(alignment, 'shrinkToFit', '0') === '1'
                },
                border: borders[borderId] || {},
                numFmt: customFormats[numFmtId] || builtInFormat(numFmtId)
            };
            if (fill.pattern && fill.pattern !== 'none' && fill.pattern !== 'gray125') {
                result.fill = fill.fgColor || fill.bgColor || '';
                result.pattern = fill.pattern;
            }
            if (!result.font && !result.fill && !result.numFmt
                    && !Object.keys(result.border).length
                    && !result.alignment.horizontal && !result.alignment.vertical
                    && !result.alignment.wrapText) {
                return null;
            }
            return result;
        });
    }

    function parseSheetMeta(files, path, limits, themeColors) {
        var result = defaultSheetMeta();
        if (!files[path]) {
            return result;
        }
        var root = Xml.parse(Zip.xmlText(files, path, limits));
        var format = Xml.child(root, 'sheetFormatPr');
        var defaultRowHeight = parseFloat(Xml.attr(format, 'defaultRowHeight', ''));
        var defaultColumnWidth = parseFloat(Xml.attr(format, 'defaultColWidth', ''));
        if (isFinite(defaultRowHeight) && defaultRowHeight > 0) {
            result.defaultRowHeight = defaultRowHeight * POINTS_TO_PIXELS;
        }
        if (isFinite(defaultColumnWidth) && defaultColumnWidth > 0) {
            result.defaultColumnWidth = widthToPixels(defaultColumnWidth);
        }

        Xml.children(Xml.child(root, 'cols'), 'col').forEach(function (node) {
            var min = parseInt(Xml.attr(node, 'min', '0'), 10);
            var max = parseInt(Xml.attr(node, 'max', '0'), 10);
            var width = parseFloat(Xml.attr(node, 'width', ''));
            var hidden = Xml.attr(node, 'hidden', '0') === '1';
            var styleId = parseInt(Xml.attr(node, 'style', '0'), 10) || 0;
            if (min < 1 || max < min) {
                return;
            }
            for (var column = min; column <= max; column += 1) {
                result.columns[column - 1] = {
                    width: hidden ? 0 : isFinite(width) && width > 0
                        ? widthToPixels(width) : result.defaultColumnWidth,
                    hidden: hidden,
                    styleId: styleId
                };
            }
        });

        Xml.children(Xml.child(root, 'sheetData'), 'row').forEach(function (rowNode) {
            var row = parseInt(Xml.attr(rowNode, 'r', '0'), 10) - 1;
            if (row < 0) {
                return;
            }
            var heightPoints = parseFloat(Xml.attr(rowNode, 'ht', ''));
            var hidden = Xml.attr(rowNode, 'hidden', '0') === '1';
            var rowStyleId = parseInt(Xml.attr(rowNode, 's', '0'), 10) || 0;
            result.rows[row] = {
                height: hidden ? 0 : isFinite(heightPoints) && heightPoints > 0
                    ? heightPoints * POINTS_TO_PIXELS : 0,
                hidden: hidden,
                styleId: rowStyleId
            };
            Xml.children(rowNode, 'c').forEach(function (cellNode) {
                var address = Xml.attr(cellNode, 'r', '');
                if (!address) {
                    return;
                }
                var cellStyleId = parseInt(Xml.attr(cellNode, 's', ''), 10);
                if (!isFinite(cellStyleId)) {
                    var columnMeta = result.columns[columnNumber(address) - 1];
                    cellStyleId = rowStyleId
                        || columnMeta && columnMeta.styleId
                        || 0;
                }
                result.cells[address] = {
                    styleId: cellStyleId,
                    comment: ''
                };
            });
        });

        var relPath = relationshipsPath(path);
        var rels = parseRelationships(files[relPath], limits);
        Object.keys(rels).forEach(function (id) {
            var relationship = rels[id];
            if (!relationship.type || !/\/comments$/i.test(relationship.type)) {
                return;
            }
            var commentsPath = resolvePart(path, relationship.target);
            if (!files[commentsPath]) {
                return;
            }
            try {
                var commentsRoot = Xml.parse(
                    Zip.xmlText(files, commentsPath, limits));
                Xml.descendants(commentsRoot, 'comment').forEach(function (node) {
                    var ref = Xml.attr(node, 'ref', '');
                    if (!ref) {
                        return;
                    }
                    var commentText = Xml.text(Xml.child(node, 'text')).trim();
                    result.cells[ref] = result.cells[ref] || {
                        styleId: 0,
                        comment: ''
                    };
                    result.cells[ref].comment = commentText;
                    result.hasComments = true;
                });
            } catch (ignore) {
                // Comments are optional presentation content.
            }
        });
        return result;
    }

    function defaultSheetMeta() {
        return {
            rows: {},
            columns: {},
            cells: {},
            defaultRowHeight: DEFAULT_ROW_HEIGHT * POINTS_TO_PIXELS,
            defaultColumnWidth: widthToPixels(DEFAULT_COLUMN_WIDTH),
            hasComments: false
        };
    }

    function parseRelationships(xml, limits) {
        if (!xml) {
            return {};
        }
        var text = typeof xml === 'string' ? xml : Zip.xmlText(
            {'relationship.xml': xml}, 'relationship.xml', limits);
        var root;
        try {
            root = Xml.parse(text);
        } catch (ignore) {
            return {};
        }
        var result = {};
        Xml.children(root, 'Relationship').forEach(function (node) {
            var id = Xml.attr(node, 'Id', '');
            if (id) {
                result[id] = {
                    type: Xml.attr(node, 'Type', ''),
                    target: Xml.attr(node, 'Target', ''),
                    targetMode: Xml.attr(node, 'TargetMode', '')
                };
            }
        });
        return result;
    }

    function parseColor(node, themeColors) {
        if (!node) {
            return '';
        }
        if (Xml.attr(node, 'auto', '') === '1') {
            return '#000000';
        }
        var rgb = Xml.attr(node, 'rgb', '');
        if (rgb) {
            return normalizeHex(rgb);
        }
        var indexed = parseInt(Xml.attr(node, 'indexed', ''), 10);
        if (isFinite(indexed)) {
            return INDEXED_COLORS[indexed] || '#000000';
        }
        var theme = parseInt(Xml.attr(node, 'theme', ''), 10);
        if (isFinite(theme)) {
            var color = themeColors[theme] || '#000000';
            var tint = parseFloat(Xml.attr(node, 'tint', '0'));
            return applyTint(color, isFinite(tint) ? tint : 0);
        }
        return '';
    }

    function normalizeHex(value) {
        var text = String(value || '').replace(/^#/, '');
        if (text.length === 8) {
            text = text.slice(-6);
        }
        return /^[0-9a-f]{6}$/i.test(text) ? '#' + text.toUpperCase() : '';
    }

    function applyTint(hex, tint) {
        if (!tint) {
            return hex;
        }
        var value = normalizeHex(hex).substring(1);
        var channels = [0, 2, 4].map(function (offset) {
            var channel = parseInt(value.substring(offset, offset + 2), 16) / 255;
            var adjusted = tint < 0
                ? channel * (1 + tint)
                : channel + (1 - channel) * tint;
            return Math.max(0, Math.min(255, Math.round(adjusted * 255)));
        });
        return '#' + channels.map(function (channel) {
            return channel.toString(16).padStart(2, '0');
        }).join('').toUpperCase();
    }

    function widthToPixels(width) {
        if (!isFinite(width) || width <= 0) {
            return 0;
        }
        return Math.floor(
            ((width * 256 + Math.floor(128 / MAX_DIGIT_WIDTH)) / 256)
            * MAX_DIGIT_WIDTH
        );
    }

    function resolvePart(basePart, target) {
        var text = String(target || '').replace(/\\/g, '/');
        if (text.charAt(0) === '/') {
            return normalizePart(text.substring(1));
        }
        var base = basePart.substring(0, basePart.lastIndexOf('/'));
        return normalizePart(base + '/' + text);
    }

    function normalizePart(value) {
        var parts = String(value || '').split('/');
        var result = [];
        parts.forEach(function (part) {
            if (!part || part === '.') {
                return;
            }
            if (part === '..') {
                result.pop();
                return;
            }
            result.push(part);
        });
        return result.join('/');
    }

    function relationshipsPath(part) {
        var slash = part.lastIndexOf('/');
        return part.substring(0, slash + 1) + '_rels/'
            + part.substring(slash + 1) + '.rels';
    }

    function builtInFormat(id) {
        var formats = {
            '0': 'General',
            '1': '0',
            '2': '0.00',
            '3': '#,##0',
            '4': '#,##0.00',
            '9': '0%',
            '10': '0.00%',
            '11': '0.00E+00',
            '12': '# ?/?',
            '13': '# ??/??',
            '14': 'm/d/yy',
            '15': 'd-mmm-yy',
            '16': 'd-mmm',
            '17': 'mmm-yy',
            '18': 'h:mm AM/PM',
            '19': 'h:mm:ss AM/PM',
            '20': 'h:mm',
            '21': 'h:mm:ss',
            '22': 'm/d/yy h:mm',
            '37': '#,##0 ;(#,##0)',
            '38': '#,##0 ;[Red](#,##0)',
            '39': '#,##0.00;(#,##0.00)',
            '40': '#,##0.00;[Red](#,##0.00)',
            '45': 'mm:ss',
            '46': '[h]:mm:ss',
            '47': 'mmss.0',
            '48': '##0.0E+0',
            '49': '@'
        };
        return formats[id] || '';
    }

    function columnNumber(address) {
        var match = /^([A-Z]+)/i.exec(String(address || ''));
        if (!match) {
            return 0;
        }
        var value = 0;
        var text = match[1].toUpperCase();
        for (var i = 0; i < text.length; i += 1) {
            value = value * 26 + text.charCodeAt(i) - 64;
        }
        return value;
    }

    global.OfficeSpreadsheetOoxml = {
        parse: parseContext,
        widthToPixels: widthToPixels
    };
})(self);
