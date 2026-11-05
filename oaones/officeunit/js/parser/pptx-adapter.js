(function (global) {
    'use strict';

    var Zip = global.OfficeZipSecurity;
    var Xml = global.OfficeXmlLite;

    function parse(buffer, extension, limits) {
        var inspection = Zip.inspect(buffer, limits);
        if (inspection.macro) {
            throw new Zip.PreviewError('OFFICE_PREVIEW_MACRO_BLOCKED', '宏内容已阻止');
        }
        if (inspection.kind !== 'pptx') {
            throw new Zip.PreviewError('OFFICE_PREVIEW_ARCHIVE_INVALID',
                '文件内容与 PowerPoint 扩展名不匹配');
        }
        var files = Zip.unzipSelected(buffer, inspection, function (name) {
            return name === '[Content_Types].xml'
                || name.indexOf('docProps/') === 0
                || (name.indexOf('ppt/') === 0 && !/\.bin$/i.test(name));
        }, limits.maxUncompressedBytes);
        var warnings = Zip.validateXmlFiles(files, limits);
        var defaultFont = themeFont(files, limits);
        var presentation = Xml.parse(Zip.xmlText(
            files, 'ppt/presentation.xml', limits));
        var presentationRels = rels(files, 'ppt/_rels/presentation.xml.rels', limits);
        var sizeNode = Xml.firstDescendant(presentation, 'sldSz');
        var width = Number(Xml.attr(sizeNode, 'cx', 12192000));
        var height = Number(Xml.attr(sizeNode, 'cy', 6858000));
        var slideIds = Xml.descendants(presentation, 'sldId');
        if (slideIds.length > limits.maxSlideCount) {
            throw new Zip.PreviewError('OFFICE_PREVIEW_ARCHIVE_TOO_LARGE',
                '幻灯片数量超过限制');
        }

        var assets = [];
        var assetMap = new Map();
        var totalAssetBytes = 0;
        var totalShapes = 0;
        var slides = [];
        slideIds.forEach(function (slideId, index) {
            var relId = Xml.attr(slideId, 'r:id', '');
            var relation = presentationRels[relId];
            if (!relation) {
                warnings.push('第 ' + (index + 1) + ' 张幻灯片关系缺失。');
                return;
            }
            var slidePath = resolvePath('ppt/presentation.xml', relation.target);
            var slideText = files[slidePath]
                ? Zip.xmlText(files, slidePath, limits) : '';
            if (!slideText) {
                warnings.push('第 ' + (index + 1) + ' 张幻灯片文件缺失。');
                return;
            }
            var slideRoot = Xml.parse(slideText);
            var slideRels = rels(files, relsPath(slidePath), limits);
            var background = parseBackground(slideRoot);
            if (!background) {
                background = inheritedBackground(files, slidePath, slideRels, limits);
            }
            var shapeTree = Xml.firstDescendant(slideRoot, 'spTree');
            var shapes = [];
            if (shapeTree) {
                Xml.children(shapeTree).forEach(function (node) {
                    if (node.local === 'nvGrpSpPr' || node.local === 'grpSpPr') {
                        return;
                    }
                    totalShapes += 1;
                    if (totalShapes > limits.maxShapeCount) {
                        throw new Zip.PreviewError('OFFICE_PREVIEW_ARCHIVE_TOO_LARGE',
                            '图形数量超过限制');
                    }
                    var shape = parseShape(node, files, slideRels, width, height,
                        limits, assets, assetMap, warnings, defaultFont);
                    if (shape) {
                        shape.z = shapes.length;
                        shapes.push(shape);
                        if (shape.kind === 'image') {
                            totalAssetBytes += shape.assetSize || 0;
                        }
                    }
                });
            }
            if (totalAssetBytes > limits.maxAssetBytes) {
                throw new Zip.PreviewError('OFFICE_PREVIEW_ARCHIVE_TOO_LARGE',
                    '图片资源总量超过限制');
            }
            var title = '';
            for (var i = 0; i < shapes.length; i += 1) {
                if (shapes[i].kind === 'text' && shapes[i].text) {
                    title = shapes[i].text.replace(/\s+/g, ' ').trim().substring(0, 80);
                    break;
                }
            }
            slides.push({
                index: index,
                id: String(Xml.attr(slideId, 'id', index + 1)),
                width: width,
                height: height,
                background: background || '#ffffff',
                title: title || ('幻灯片 ' + (index + 1)),
                shapes: shapes
            });
        });
        if (!slides.length) {
            throw new Zip.PreviewError('OFFICE_PREVIEW_PARSE_FAILED',
                '演示文稿没有可显示的幻灯片');
        }
        var metadata = documentProperties(files, limits);
        return {
            format: 'pptx',
            supportLevel: extension === 'pptx' ? 'standard' : 'best_effort',
            metadata: metadata,
            width: width,
            height: height,
            slides: slides,
            assets: assets,
            warnings: warnings
        };
    }

    function parseShape(node, files, slideRels, width, height, limits,
                        assets, assetMap, warnings, defaultFont) {
        if (node.local === 'grpSp') {
            return placeholder(node, '组合图形暂不支持完整还原', width, height);
        }
        if (node.local === 'cxnSp') {
            return parseConnector(node);
        }
        if (node.local === 'sp') {
            return parseTextShape(node, width, height, defaultFont);
        }
        if (node.local === 'pic') {
            return parseImage(node, files, slideRels, limits, assets, assetMap, warnings);
        }
        if (node.local === 'graphicFrame') {
            var table = Xml.firstDescendant(node, 'tbl');
            if (table) {
                return parseTable(node, table, width, height, defaultFont, limits);
            }
            var chartReference = Xml.firstDescendant(node, 'chart');
            if (chartReference) {
                return parseChart(node, chartReference, files, slideRels, limits,
                    width, height, warnings);
            }
            return placeholder(node, '嵌入对象已降级显示', width, height);
        }
        return null;
    }

    function parseTextShape(node, fallbackWidth, fallbackHeight, defaultFont) {
        var properties = Xml.child(node, 'spPr');
        var transform = readTransform(properties);
        var textBody = Xml.child(node, 'txBody');
        if (!textBody && !transform) {
            return null;
        }
        var paragraphs = parseParagraphs(textBody, defaultFont);
        var box = transform || {
            x: 0,
            y: 0,
            width: Math.round(fallbackWidth * 0.4),
            height: Math.round(fallbackHeight * 0.25),
            rotation: 0
        };
        box.kind = 'text';
        box.text = paragraphs.map(function (paragraph) {
            return paragraph.runs.map(function (run) {
                return run.text;
            }).join('');
        }).join('\n');
        box.paragraphs = paragraphs;
        box.fill = shapeFill(properties);
        box.line = shapeLine(properties);
        box.geometry = geometry(properties);
        box.verticalAlign = bodyVerticalAlign(textBody);
        box.margins = bodyMargins(textBody);
        return box;
    }

    function parseImage(node, files, slideRels, limits, assets, assetMap, warnings) {
        var properties = Xml.child(node, 'spPr');
        var box = readTransform(properties);
        if (!box) {
            return placeholder(node, '图片位置信息缺失', 1, 1);
        }
        var blip = Xml.firstDescendant(node, 'blip');
        var relId = Xml.attr(blip, 'r:embed', '');
        var relation = slideRels[relId];
        if (!relation || relation.external) {
            warnings.push('外链图片已阻止。');
            box.kind = 'placeholder';
            box.text = '外部图片未加载';
            return box;
        }
        var target = resolvePath('ppt/slides/slide1.xml', relation.target);
        var bytes = files[target];
        if (!bytes) {
            warnings.push('内嵌图片资源缺失。');
            box.kind = 'placeholder';
            box.text = '图片资源缺失';
            return box;
        }
        var key = 'asset-' + hashName(target);
        var isNewAsset = !assetMap.has(key);
        if (isNewAsset) {
            var mime = mimeType(target);
            assetMap.set(key, {
                key: key,
                mime: mime,
                name: target.substring(target.lastIndexOf('/') + 1),
                bytes: bytes
            });
            assets.push(assetMap.get(key));
        }
        box.kind = 'image';
        box.assetKey = key;
        box.assetSize = isNewAsset ? bytes.length : 0;
        box.crop = parseCrop(node);
        return box;
    }

    function parseTable(frame, table, fallbackWidth, fallbackHeight, defaultFont, limits) {
        var box = readTransform(Xml.child(frame, 'xfrm')) || {
            x: 0,
            y: 0,
            width: fallbackWidth,
            height: fallbackHeight,
            rotation: 0
        };
        var columns = Xml.children(table, 'gridCol').map(function (col) {
            return Number(Xml.attr(col, 'w', 0));
        });
        var cellCount = 0;
        var rows = Xml.children(table, 'tr').map(function (row) {
            var cells = Xml.children(row, 'tc').map(function (cell) {
                cellCount += 1;
                if (cellCount > limits.maxShapeCount) {
                    throw new Zip.PreviewError('OFFICE_PREVIEW_ARCHIVE_TOO_LARGE',
                        '幻灯片表格单元格数量超过限制');
                }
                var properties = Xml.child(cell, 'tcPr');
                return {
                    paragraphs: parseParagraphs(Xml.child(cell, 'txBody'), defaultFont),
                    fill: shapeFill(properties),
                    colspan: Number(Xml.attr(cell, 'gridSpan', 1)),
                    rowspan: Number(Xml.attr(cell, 'rowSpan', 1))
                };
            });
            return {
                height: Number(Xml.attr(row, 'h', 0)),
                cells: cells
            };
        });
        box.kind = 'table';
        box.columns = columns;
        box.rows = rows;
        box.fill = shapeFill(Xml.child(frame, 'xfrm'));
        return box;
    }

    function parseChart(frame, chartReference, files, slideRels, limits,
                        fallbackWidth, fallbackHeight, warnings) {
        var box = readTransform(Xml.child(frame, 'xfrm')) || {
            x: 0,
            y: 0,
            width: fallbackWidth,
            height: fallbackHeight,
            rotation: 0
        };
        var relId = Xml.attr(chartReference, 'r:id', '');
        var relation = slideRels[relId];
        if (!relation || relation.external || !files[resolvePath(
            'ppt/slides/slide1.xml', relation.target)]) {
            warnings.push('一个图表资源无法读取，已显示占位。');
            box.kind = 'placeholder';
            box.text = '图表暂不支持';
            return box;
        }
        var chartPath = resolvePath('ppt/slides/slide1.xml', relation.target);
        var root = Xml.parse(Zip.xmlText(files, chartPath, limits));
        var type = chartType(root);
        var titleNode = Xml.firstDescendant(root, 'title');
        var title = titleNode ? Xml.text(titleNode).trim() : '';
        var categories = [];
        var series = Xml.descendants(root, 'ser').map(function (ser) {
            var localCategories = chartValues(ser, 'cat');
            if (!categories.length) {
                categories = localCategories;
            }
            return {
                name: chartTextValue(Xml.child(ser, 'tx')) || '',
                values: chartValues(ser, 'val').map(function (value) {
                    return Number(value);
                }).filter(function (value) {
                    return isFinite(value);
                })
            };
        });
        box.kind = 'chart';
        box.chart = {
            type: type,
            title: title,
            categories: categories,
            series: series
        };
        return box;
    }

    function chartType(root) {
        if (Xml.firstDescendant(root, 'pieChart')
                || Xml.firstDescendant(root, 'doughnutChart')
                || Xml.firstDescendant(root, 'pie3DChart')) {
            return 'pie';
        }
        if (Xml.firstDescendant(root, 'lineChart')
                || Xml.firstDescendant(root, 'line3DChart')) {
            return 'line';
        }
        if (Xml.firstDescendant(root, 'barChart')
                || Xml.firstDescendant(root, 'bar3DChart')) {
            return 'bar';
        }
        return 'unknown';
    }

    function chartValues(series, containerName) {
        var container = Xml.child(series, containerName);
        if (!container) {
            return [];
        }
        var points = Xml.descendants(container, 'pt');
        if (!points.length) {
            return [];
        }
        return points.map(function (point) {
            var value = Xml.child(point, 'v');
            return value ? Xml.text(value) : '';
        });
    }

    function chartTextValue(container) {
        if (!container) {
            return '';
        }
        var valueNodes = Xml.descendants(container, 'v');
        return valueNodes.length ? Xml.text(valueNodes[0]) : Xml.text(container);
    }

    function parseParagraphs(textBody, defaultFont) {
        if (!textBody) {
            return [];
        }
        var paragraphs = Xml.children(textBody, 'p');
        if (!paragraphs.length) {
            return [{
                align: 'left',
                bullet: false,
                runs: [{text: Xml.text(textBody), size: 18, bold: false,
                    italic: false, color: '#1f2937', font: defaultFont || ''}]
            }];
        }
        return paragraphs.map(function (paragraph) {
            var properties = Xml.child(paragraph, 'pPr');
            var runs = [];
            Xml.children(paragraph).forEach(function (node) {
                if (node.local === 'r' || node.local === 'fld') {
                    var value = Xml.child(node, 't');
                    var runProperties = Xml.child(node, 'rPr');
                    runs.push({
                        text: value ? Xml.text(value) : '',
                        size: Number(Xml.attr(runProperties, 'sz', 1800)) / 100,
                        bold: boolAttr(runProperties, 'b'),
                        italic: boolAttr(runProperties, 'i'),
                        underline: Xml.attr(runProperties, 'u', '') !== '',
                        color: color(runProperties) || '#1f2937',
                        font: Xml.attr(runProperties, 'typeface', defaultFont || '')
                    });
                } else if (node.local === 'br') {
                    runs.push({text: '\n', size: 18, bold: false, italic: false,
                        color: '#1f2937', font: defaultFont || ''});
                }
            });
            if (!runs.length) {
                runs.push({text: Xml.text(paragraph), size: 18, bold: false,
                    italic: false, color: '#1f2937', font: defaultFont || ''});
            }
            return {
                align: alignment(Xml.attr(properties, 'algn', 'l')),
                bullet: !!Xml.child(properties, 'buChar')
                    || !!Xml.child(properties, 'buAutoNum'),
                level: Number(Xml.attr(properties, 'lvl', 0)),
                runs: runs
            };
        });
    }

    function readTransform(properties) {
        if (!properties) {
            return null;
        }
        var transform = Xml.child(properties, 'xfrm');
        if (!transform) {
            return null;
        }
        var offset = Xml.child(transform, 'off');
        var extent = Xml.child(transform, 'ext');
        if (!offset || !extent) {
            return null;
        }
        return {
            x: Number(Xml.attr(offset, 'x', 0)),
            y: Number(Xml.attr(offset, 'y', 0)),
            width: Number(Xml.attr(extent, 'cx', 0)),
            height: Number(Xml.attr(extent, 'cy', 0)),
            rotation: Number(Xml.attr(transform, 'rot', 0)) / 60000,
            flipH: boolAttr(transform, 'flipH'),
            flipV: boolAttr(transform, 'flipV')
        };
    }

    function shapeFill(properties) {
        if (!properties) {
            return 'transparent';
        }
        if (Xml.child(properties, 'noFill')) {
            return 'transparent';
        }
        var solid = Xml.child(properties, 'solidFill');
        if (solid) {
            return color(solid) || 'transparent';
        }
        var gradient = Xml.child(properties, 'gradFill');
        if (gradient) {
            var stop = Xml.firstDescendant(gradient, 'srgbClr');
            if (stop) {
                return '#' + Xml.attr(stop, 'val', 'ffffff');
            }
        }
        return 'transparent';
    }

    function shapeLine(properties) {
        if (!properties) {
            return null;
        }
        var line = Xml.child(properties, 'ln');
        if (!line || Xml.child(line, 'noFill')) {
            return null;
        }
        var solid = Xml.child(line, 'solidFill');
        return {
            color: color(solid) || '#d1d5db',
            width: Number(Xml.attr(line, 'w', 9525)) / 9525
        };
    }

    function geometry(properties) {
        var preset = properties ? Xml.child(properties, 'prstGeom') : null;
        return preset ? Xml.attr(preset, 'prst', 'rect') : 'rect';
    }

    function color(fillNode) {
        if (!fillNode) {
            return '';
        }
        var srgb = Xml.firstDescendant(fillNode, 'srgbClr');
        if (srgb) {
            return '#' + Xml.attr(srgb, 'val', '000000');
        }
        var scheme = Xml.firstDescendant(fillNode, 'schemeClr');
        if (scheme) {
            var value = Xml.attr(scheme, 'val', '');
            var theme = {
                bg1: '#ffffff',
                lt1: '#ffffff',
                tx1: '#000000',
                dk1: '#000000',
                bg2: '#f3f4f6',
                lt2: '#f3f4f6',
                tx2: '#1f2937',
                dk2: '#1f2937'
            };
            return theme[value] || '';
        }
        return '';
    }

    function parseBackground(root) {
        var bg = Xml.child(root, 'bg');
        var properties = bg ? Xml.child(bg, 'bgPr') : null;
        return properties ? shapeFill(properties) : '';
    }

    function inheritedBackground(files, slidePath, slideRels, limits) {
        var layoutRelation = firstRelation(slideRels, 'slideLayout');
        if (!layoutRelation || layoutRelation.external) {
            return '';
        }
        var layoutPath = resolvePath(slidePath, layoutRelation.target);
        if (!files[layoutPath]) {
            return '';
        }
        var layout = Xml.parse(Zip.xmlText(files, layoutPath, limits));
        var value = parseBackground(layout);
        if (value) {
            return value;
        }
        var layoutRels = rels(files, relsPath(layoutPath), limits);
        var masterRelation = firstRelation(layoutRels, 'slideMaster');
        if (!masterRelation || masterRelation.external
                || !files[resolvePath(layoutPath, masterRelation.target)]) {
            return '';
        }
        var masterPath = resolvePath(layoutPath, masterRelation.target);
        return parseBackground(Xml.parse(Zip.xmlText(files, masterPath, limits)));
    }

    function rels(files, path, limits) {
        var result = {};
        if (!files[path]) {
            return result;
        }
        var root = Xml.parse(Zip.xmlText(files, path, limits));
        Xml.children(root, 'Relationship').forEach(function (node) {
            result[Xml.attr(node, 'Id', '')] = {
                id: Xml.attr(node, 'Id', ''),
                type: Xml.attr(node, 'Type', ''),
                target: Xml.attr(node, 'Target', ''),
                external: Xml.attr(node, 'TargetMode', '') === 'External'
            };
        });
        return result;
    }

    function firstRelation(map, suffix) {
        var keys = Object.keys(map);
        for (var i = 0; i < keys.length; i += 1) {
            var item = map[keys[i]];
            if (item.type && item.type.split('/').pop() === suffix) {
                return item;
            }
        }
        return null;
    }

    function relsPath(path) {
        var slash = path.lastIndexOf('/');
        var directory = slash < 0 ? '' : path.substring(0, slash + 1);
        return directory + '_rels/' + path.substring(slash + 1) + '.rels';
    }

    function resolvePath(basePath, target) {
        if (!target) {
            return '';
        }
        var targetText = target.replace(/\\/g, '/');
        var combined;
        if (targetText.charAt(0) === '/') {
            combined = targetText.substring(1);
        } else {
            var slash = basePath.lastIndexOf('/');
            combined = (slash < 0 ? '' : basePath.substring(0, slash + 1)) + targetText;
        }
        var parts = combined.split('/');
        var result = [];
        parts.forEach(function (part) {
            if (!part || part === '.') {
                return;
            }
            if (part === '..') {
                if (!result.length) {
                    throw new Zip.PreviewError('OFFICE_PREVIEW_ARCHIVE_INVALID',
                        '关系目标包含目录穿越');
                }
                result.pop();
            } else {
                result.push(part);
            }
        });
        return result.join('/');
    }

    function parseCrop(node) {
        var source = Xml.firstDescendant(node, 'srcRect');
        if (!source) {
            return null;
        }
        return {
            left: Number(Xml.attr(source, 'l', 0)) / 100000,
            top: Number(Xml.attr(source, 't', 0)) / 100000,
            right: Number(Xml.attr(source, 'r', 0)) / 100000,
            bottom: Number(Xml.attr(source, 'b', 0)) / 100000
        };
    }

    function placeholder(node, text, width, height) {
        var properties = Xml.child(node, 'spPr')
            || (Xml.child(node, 'grpSpPr') || Xml.child(node, 'xfrm'));
        var box = readTransform(properties) || {
            x: Math.round(width * 0.1),
            y: Math.round(height * 0.1),
            width: Math.round(width * 0.8),
            height: Math.round(height * 0.2),
            rotation: 0
        };
        box.kind = 'placeholder';
        box.text = text;
        return box;
    }

    function parseConnector(node) {
        var properties = Xml.child(node, 'spPr');
        var box = readTransform(properties);
        if (!box) {
            return null;
        }
        box.kind = 'connector';
        box.line = shapeLine(properties);
        return box;
    }

    function bodyVerticalAlign(textBody) {
        var body = textBody ? Xml.child(textBody, 'bodyPr') : null;
        var value = Xml.attr(body, 'anchor', 't');
        if (value === 'ctr') {
            return 'center';
        }
        if (value === 'b') {
            return 'bottom';
        }
        return 'top';
    }

    function bodyMargins(textBody) {
        var body = textBody ? Xml.child(textBody, 'bodyPr') : null;
        return {
            left: Number(Xml.attr(body, 'lIns', 91440)) / 9525,
            top: Number(Xml.attr(body, 'tIns', 45720)) / 9525,
            right: Number(Xml.attr(body, 'rIns', 91440)) / 9525,
            bottom: Number(Xml.attr(body, 'bIns', 45720)) / 9525
        };
    }

    function alignment(value) {
        if (value === 'ctr') {
            return 'center';
        }
        if (value === 'r') {
            return 'right';
        }
        if (value === 'just') {
            return 'justify';
        }
        return 'left';
    }

    function boolAttr(node, name) {
        var value = Xml.attr(node, name, '0');
        return value === '1' || value === 'true';
    }

    function documentProperties(files, limits) {
        var results = {};
        ['docProps/core.xml', 'docProps/app.xml'].forEach(function (name) {
            if (!files[name]) {
                return;
            }
            try {
                var root = Xml.parse(Zip.xmlText(files, name, limits));
                ['title', 'creator', 'lastModifiedBy', 'Application']
                    .forEach(function (key) {
                        var node = Xml.firstDescendant(root, key);
                        if (node) {
                            results[key] = Xml.text(node).trim();
                        }
                    });
            } catch (ignore) {
                // Optional metadata must not block preview.
            }
        });
        return results;
    }

    function themeFont(files, limits) {
        var names = Object.keys(files).filter(function (name) {
            return /^ppt\/theme\/theme\d+\.xml$/i.test(name);
        }).sort();
        if (!names.length) {
            return '';
        }
        try {
            var root = Xml.parse(Zip.xmlText(files, names[0], limits));
            var minor = Xml.firstDescendant(root, 'minorFont');
            var major = Xml.firstDescendant(root, 'majorFont');
            var fontNode = minor ? Xml.firstDescendant(minor, 'ea')
                || Xml.firstDescendant(minor, 'latin') : null;
            if (!fontNode && major) {
                fontNode = Xml.firstDescendant(major, 'ea')
                    || Xml.firstDescendant(major, 'latin');
            }
            return fontNode ? Xml.attr(fontNode, 'typeface', '') : '';
        } catch (ignore) {
            return '';
        }
    }

    function mimeType(path) {
        var ext = path.substring(path.lastIndexOf('.') + 1).toLowerCase();
        var values = {
            png: 'image/png',
            jpg: 'image/jpeg',
            jpeg: 'image/jpeg',
            gif: 'image/gif',
            bmp: 'image/bmp',
            webp: 'image/webp',
            svg: 'image/svg+xml',
            tif: 'image/tiff',
            tiff: 'image/tiff'
        };
        return values[ext] || 'application/octet-stream';
    }

    function hashName(value) {
        var hash = 2166136261;
        for (var i = 0; i < value.length; i += 1) {
            hash ^= value.charCodeAt(i);
            hash = Math.imul(hash, 16777619);
        }
        return (hash >>> 0).toString(16);
    }

    global.OfficePptxAdapter = {
        parse: parse
    };
})(self);
