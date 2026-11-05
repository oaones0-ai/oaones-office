(function (global) {
    'use strict';

    var EMU_PER_PIXEL = 9525;
    var SVG_NS = 'http://www.w3.org/2000/svg';

    function render(options) {
        var container = options.container;
        var destroyed = options.destroyed;
        var model = options.model;
        container.replaceChildren();
        var root = element('div', 'oa-pptx');
        var scroller = element('div', 'oa-pptx-scroll');
        root.appendChild(scroller);
        container.appendChild(root);
        var scale = 1;
        var activeIndex = 0;
        var frame = 0;
        var resizeObserver = null;
        var assetUrls = new Map();
        var slides = [];

        model.assets.forEach(function (asset) {
            if (!asset || !asset.bytes) {
                return;
            }
            var blob = new Blob([asset.bytes], {type: asset.mime || 'application/octet-stream'});
            assetUrls.set(asset.key, URL.createObjectURL(blob));
        });

        var observer = typeof IntersectionObserver === 'function'
            ? new IntersectionObserver(function (entries) {
                entries.forEach(function (entry) {
                    if (entry.isIntersecting) {
                        mountSlide(entry.target.__oaSlide);
                    }
                });
            }, {root: scroller, rootMargin: '1000px 0px'})
            : null;

        model.slides.forEach(function (slide, index) {
            var shell = element('section', 'oa-slide-shell');
            shell.setAttribute('data-oa-page', String(index + 1));
            shell.setAttribute('data-oa-slide', String(index + 1));
            shell.setAttribute('data-oa-anchor', 'slide-' + (index + 1));
            shell.id = 'oa-slide-' + (index + 1);
            var stage = element('div', 'oa-slide-stage');
            stage.style.width = toPx(slide.width) + 'px';
            stage.style.height = toPx(slide.height) + 'px';
            stage.style.background = slide.background || '#ffffff';
            shell.appendChild(stage);
            var state = {
                index: index,
                slide: slide,
                shell: shell,
                stage: stage,
                mounted: false
            };
            shell.__oaSlide = state;
            slides.push(state);
            scroller.appendChild(shell);
            if (observer) {
                observer.observe(shell);
            }
        });

        function mountSlide(state) {
            if (!state || state.mounted || destroyed()) {
                return;
            }
            state.mounted = true;
            state.slide.shapes.forEach(function (shape) {
                mountShape(state.stage, shape, state.index);
            });
        }

        function mountShape(stage, shape, slideIndex) {
            var node;
            if (shape.kind === 'text') {
                node = renderTextShape(shape);
            } else if (shape.kind === 'image') {
                node = renderImageShape(shape);
            } else if (shape.kind === 'table') {
                node = renderTableShape(shape);
            } else if (shape.kind === 'chart') {
                node = renderChartShape(shape);
            } else if (shape.kind === 'connector') {
                node = renderConnector(shape);
            } else {
                node = renderPlaceholder(shape);
            }
            if (!node) {
                return;
            }
            node.setAttribute('data-oa-order', String(shape.z || 0));
            node.setAttribute('data-oa-role',
                shape.kind === 'text' ? 'slide-text'
                    : shape.kind === 'table' ? 'table' : shape.kind);
            node.setAttribute('data-oa-block-id',
                'slide-' + (slideIndex + 1) + '-shape-' + (shape.z + 1));
            position(node, shape);
            stage.appendChild(node);
        }

        function renderTextShape(shape) {
            var node = element('div', 'oa-pptx-text');
            applyBoxStyle(node, shape);
            if (shape.fill && shape.fill !== 'transparent') {
                node.style.background = shape.fill;
            }
            if (shape.line) {
                node.style.border = Math.max(1, shape.line.width) + 'px solid '
                    + shape.line.color;
            }
            if (shape.geometry === 'ellipse') {
                node.style.borderRadius = '50%';
            } else if (shape.geometry === 'roundRect') {
                node.style.borderRadius = '8px';
            }
            var margins = shape.margins || {left: 9.6, top: 4.8, right: 9.6, bottom: 4.8};
            node.style.padding = [margins.top, margins.right,
                margins.bottom, margins.left].join('px ') + 'px';
            node.style.justifyContent = shape.verticalAlign === 'center' ? 'center'
                : shape.verticalAlign === 'bottom' ? 'flex-end' : 'flex-start';
            (shape.paragraphs || []).forEach(function (paragraph) {
                var paragraphNode = element('p', 'oa-pptx-paragraph');
                paragraphNode.style.textAlign = paragraph.align || 'left';
                paragraphNode.style.paddingLeft = paragraph.bullet
                    ? (16 + Math.max(0, paragraph.level || 0) * 18) + 'px' : '0';
                if (paragraph.bullet) {
                    var bullet = element('span', 'oa-pptx-bullet', '•');
                    paragraphNode.appendChild(bullet);
                }
                (paragraph.runs || []).forEach(function (run) {
                    var span = element('span', 'oa-pptx-run', run.text);
                    span.style.fontSize = Math.max(6, run.size || 18) * 1.333 + 'px';
                    span.style.fontWeight = run.bold ? '700' : '400';
                    span.style.fontStyle = run.italic ? 'italic' : 'normal';
                    span.style.textDecoration = run.underline ? 'underline' : 'none';
                    span.style.color = run.color || '#1f2937';
                    if (run.font) {
                        span.style.fontFamily = run.font;
                    }
                    paragraphNode.appendChild(span);
                });
                node.appendChild(paragraphNode);
            });
            return node;
        }

        function renderImageShape(shape) {
            var node = element('div', 'oa-pptx-image');
            applyBoxStyle(node, shape);
            var source = assetUrls.get(shape.assetKey);
            if (!source) {
                node.classList.add('oa-pptx-placeholder');
                node.textContent = '图片资源不可用';
                return node;
            }
            var image = document.createElement('img');
            image.src = source;
            image.alt = '';
            image.setAttribute('referrerpolicy', 'no-referrer');
            image.style.width = '100%';
            image.style.height = '100%';
            image.style.objectFit = 'cover';
            if (shape.crop) {
                image.style.clipPath = 'inset('
                    + (shape.crop.top * 100) + '% '
                    + (shape.crop.right * 100) + '% '
                    + (shape.crop.bottom * 100) + '% '
                    + (shape.crop.left * 100) + '%)';
            }
            node.appendChild(image);
            return node;
        }

        function renderTableShape(shape) {
            var wrapper = element('div', 'oa-pptx-table-wrap');
            applyBoxStyle(wrapper, shape);
            var table = element('table', 'oa-pptx-table');
            var colgroup = document.createElement('colgroup');
            var widths = shape.columns && shape.columns.length
                ? shape.columns : [shape.width / Math.max(1, shape.rows[0]
                    ? shape.rows[0].cells.length : 1)];
            widths.forEach(function (width) {
                var col = document.createElement('col');
                col.style.width = toPx(width) + 'px';
                colgroup.appendChild(col);
            });
            table.appendChild(colgroup);
            var body = document.createElement('tbody');
            (shape.rows || []).forEach(function (row) {
                var tr = document.createElement('tr');
                if (row.height > 0) {
                    tr.style.height = toPx(row.height) + 'px';
                }
                (row.cells || []).forEach(function (cell) {
                    var td = document.createElement('td');
                    if (cell.colspan > 1) {
                        td.colSpan = cell.colspan;
                    }
                    if (cell.rowspan > 1) {
                        td.rowSpan = cell.rowspan;
                    }
                    if (cell.fill && cell.fill !== 'transparent') {
                        td.style.background = cell.fill;
                    }
                    (cell.paragraphs || []).forEach(function (paragraph) {
                        var p = element('p', '', paragraph.runs.map(function (run) {
                            return run.text;
                        }).join(''));
                        p.style.textAlign = paragraph.align || 'left';
                        var first = paragraph.runs && paragraph.runs[0];
                        if (first) {
                            p.style.fontSize = Math.max(6, first.size || 14) * 1.333 + 'px';
                            p.style.color = first.color || '#1f2937';
                            p.style.fontWeight = first.bold ? '700' : '400';
                        }
                        td.appendChild(p);
                    });
                    tr.appendChild(td);
                });
                body.appendChild(tr);
            });
            table.appendChild(body);
            wrapper.appendChild(table);
            return wrapper;
        }

        function renderChartShape(shape) {
            var wrapper = element('div', 'oa-pptx-chart');
            applyBoxStyle(wrapper, shape);
            var chart = shape.chart || {};
            if (!chart.series || !chart.series.length || chart.type === 'unknown') {
                wrapper.classList.add('oa-pptx-placeholder');
                wrapper.textContent = '图表暂不支持重绘';
                return wrapper;
            }
            var svg = document.createElementNS(SVG_NS, 'svg');
            svg.setAttribute('viewBox', '0 0 800 450');
            svg.setAttribute('preserveAspectRatio', 'xMidYMid meet');
            svg.setAttribute('role', 'img');
            svg.setAttribute('aria-label', chart.title || '图表');
            svg.style.width = '100%';
            svg.style.height = '100%';
            if (chart.title) {
                svg.appendChild(svgText(400, 28, chart.title, 20, 'middle', '#111827'));
            }
            if (chart.type === 'pie') {
                renderPie(svg, chart);
            } else if (chart.type === 'line') {
                renderLine(svg, chart);
            } else {
                renderBar(svg, chart);
            }
            wrapper.appendChild(svg);
            return wrapper;
        }

        function renderBar(svg, chart) {
            var values = flattenValues(chart);
            var max = Math.max.apply(Math, values.concat([1]));
            var categories = chart.categories.length
                ? chart.categories : chart.series[0].values.map(function (_, i) {
                    return String(i + 1);
                });
            var left = 64;
            var top = 58;
            var width = 690;
            var height = 330;
            svg.appendChild(svgLine(left, top + height, left + width, top + height,
                '#94a3b8', 1));
            svg.appendChild(svgLine(left, top, left, top + height, '#94a3b8', 1));
            var slot = width / Math.max(1, categories.length);
            var seriesWidth = Math.max(4, slot * 0.7 / chart.series.length);
            categories.forEach(function (category, categoryIndex) {
                var barX = left + categoryIndex * slot + slot * 0.15;
                svg.appendChild(svgText(barX + slot * 0.35, top + height + 22,
                    String(category).substring(0, 18), 12, 'middle', '#475569'));
                chart.series.forEach(function (series, seriesIndex) {
                    var value = Number(series.values[categoryIndex] || 0);
                    var barHeight = Math.max(0, value / max * height);
                    var rect = document.createElementNS(SVG_NS, 'rect');
                    rect.setAttribute('x', String(barX + seriesIndex * seriesWidth));
                    rect.setAttribute('y', String(top + height - barHeight));
                    rect.setAttribute('width', String(Math.max(2, seriesWidth - 2)));
                    rect.setAttribute('height', String(barHeight));
                    rect.setAttribute('fill', seriesColor(seriesIndex));
                    svg.appendChild(rect);
                });
            });
            renderLegend(svg, chart);
        }

        function renderLine(svg, chart) {
            var values = flattenValues(chart);
            var min = Math.min.apply(Math, values.concat([0]));
            var max = Math.max.apply(Math, values.concat([1]));
            var range = Math.max(1, max - min);
            var left = 64;
            var top = 58;
            var width = 690;
            var height = 330;
            svg.appendChild(svgLine(left, top + height, left + width, top + height,
                '#94a3b8', 1));
            svg.appendChild(svgLine(left, top, left, top + height, '#94a3b8', 1));
            var count = Math.max(1, chart.categories.length
                || chart.series[0].values.length);
            chart.series.forEach(function (series, seriesIndex) {
                var points = series.values.map(function (value, index) {
                    var x = left + (count <= 1 ? width / 2 : index * width / (count - 1));
                    var y = top + height - ((Number(value) - min) / range * height);
                    return x + ',' + y;
                });
                var line = document.createElementNS(SVG_NS, 'polyline');
                line.setAttribute('points', points.join(' '));
                line.setAttribute('fill', 'none');
                line.setAttribute('stroke', seriesColor(seriesIndex));
                line.setAttribute('stroke-width', '3');
                svg.appendChild(line);
            });
            chart.categories.forEach(function (category, index) {
                var x = left + (count <= 1 ? width / 2 : index * width / (count - 1));
                svg.appendChild(svgText(x, top + height + 22,
                    String(category).substring(0, 18), 12, 'middle', '#475569'));
            });
            renderLegend(svg, chart);
        }

        function renderPie(svg, chart) {
            var values = chart.series[0].values.map(function (value) {
                return Math.max(0, Number(value) || 0);
            });
            var total = values.reduce(function (sum, value) {
                return sum + value;
            }, 0) || 1;
            var cx = 300;
            var cy = 235;
            var radius = 145;
            var angle = -Math.PI / 2;
            values.forEach(function (value, index) {
                var next = angle + value / total * Math.PI * 2;
                if (value > 0) {
                    var x1 = cx + radius * Math.cos(angle);
                    var y1 = cy + radius * Math.sin(angle);
                    var x2 = cx + radius * Math.cos(next);
                    var y2 = cy + radius * Math.sin(next);
                    var large = next - angle > Math.PI ? 1 : 0;
                    var path = document.createElementNS(SVG_NS, 'path');
                    path.setAttribute('d', 'M ' + cx + ' ' + cy
                        + ' L ' + x1 + ' ' + y1
                        + ' A ' + radius + ' ' + radius + ' 0 ' + large
                        + ' 1 ' + x2 + ' ' + y2 + ' Z');
                    path.setAttribute('fill', seriesColor(index));
                    svg.appendChild(path);
                }
                angle = next;
            });
            values.forEach(function (value, index) {
                var y = 110 + index * 28;
                svg.appendChild(svgRect(550, y - 12, 14, 14, seriesColor(index)));
                svg.appendChild(svgText(574, y,
                    (chart.categories[index] || ('项目 ' + (index + 1)))
                    + '  ' + value, 13, 'start', '#334155'));
            });
        }

        function renderLegend(svg, chart) {
            chart.series.forEach(function (series, index) {
                var x = 64 + index * 150;
                svg.appendChild(svgRect(x, 414, 12, 12, seriesColor(index)));
                svg.appendChild(svgText(x + 18, 425,
                    series.name || ('系列 ' + (index + 1)), 12, 'start', '#334155'));
            });
        }

        function renderConnector(shape) {
            var node = element('div', 'oa-pptx-connector');
            node.style.background = shape.line ? shape.line.color : '#64748b';
            node.style.height = Math.max(1, shape.line ? shape.line.width : 1) + 'px';
            return node;
        }

        function renderPlaceholder(shape) {
            var node = element('div', 'oa-pptx-placeholder',
                shape.text || '此对象暂不支持完整显示');
            applyBoxStyle(node, shape);
            return node;
        }

        function applyBoxStyle(node, shape) {
            node.style.width = toPx(shape.width) + 'px';
            node.style.height = toPx(shape.height) + 'px';
        }

        function position(node, shape) {
            node.style.left = toPx(shape.x) + 'px';
            node.style.top = toPx(shape.y) + 'px';
            if (shape.rotation) {
                node.style.transform = 'rotate(' + shape.rotation + 'deg)';
            }
        }

        function applyScale(next, fit) {
            var available = Math.max(240, scroller.clientWidth - 48);
            var designWidth = toPx(model.width);
            var value = fit ? Math.min(1.3, available / designWidth) : next;
            scale = Math.max(0.25, Math.min(2.5, value || 1));
            slides.forEach(function (slide) {
                slide.stage.style.transform = 'scale(' + scale + ')';
                slide.shell.style.width = toPx(slide.slide.width) * scale + 'px';
                slide.shell.style.height = toPx(slide.slide.height) * scale + 'px';
            });
            scroller.scrollLeft = Math.max(0,
                (scroller.clientWidth - toPx(model.width) * scale) / 2);
            return scale;
        }

        function currentIndex() {
            var viewportTop = scroller.getBoundingClientRect().top + 20;
            var index = 0;
            slides.forEach(function (slide, slideIndex) {
                if (slide.shell.getBoundingClientRect().top <= viewportTop) {
                    index = slideIndex;
                }
            });
            return index;
        }

        scroller.addEventListener('scroll', function () {
            if (frame) {
                return;
            }
            frame = global.requestAnimationFrame(function () {
                frame = 0;
                activeIndex = currentIndex();
                if (typeof options.onSectionChange === 'function' && !destroyed()) {
                    options.onSectionChange(activeIndex);
                }
            });
        }, {passive: true});
        resizeObserver = new ResizeObserver(function () {
            applyScale(scale, false);
        });
        resizeObserver.observe(scroller);
        mountSlide(slides[0]);
        global.requestAnimationFrame(function () {
            applyScale(1, true);
        });

        return {
            sections: slides.map(function (slide) {
                return {
                    id: 'slide-' + (slide.index + 1),
                    label: slide.slide.title,
                    kind: 'slide',
                    index: slide.index
                };
            }),
            setZoom: function (value) {
                return applyScale(value, false);
            },
            fitWidth: function () {
                return applyScale(scale, true);
            },
            goTo: function (anchor) {
                var match = /^slide-(\d+)$/.exec(String(anchor || ''));
                if (!match) {
                    return false;
                }
                var target = slides[parseInt(match[1], 10) - 1];
                if (!target) {
                    return false;
                }
                mountSlide(target);
                target.shell.scrollIntoView({block: 'start', behavior: 'smooth'});
                activeIndex = target.index;
                return true;
            },
            getCurrentIndex: currentIndex,
            getCurrentAnchor: function () {
                return 'slide-' + (currentIndex() + 1);
            },
            destroy: function () {
                if (observer) {
                    observer.disconnect();
                }
                if (resizeObserver) {
                    resizeObserver.disconnect();
                }
                if (frame) {
                    global.cancelAnimationFrame(frame);
                }
                assetUrls.forEach(function (url) {
                    URL.revokeObjectURL(url);
                });
                assetUrls.clear();
                container.replaceChildren();
            }
        };
    }

    function flattenValues(chart) {
        var values = [];
        chart.series.forEach(function (series) {
            series.values.forEach(function (value) {
                if (isFinite(Number(value))) {
                    values.push(Number(value));
                }
            });
        });
        return values.length ? values : [0];
    }

    function seriesColor(index) {
        var colors = ['#0f766e', '#2563eb', '#d97706', '#dc2626',
            '#7c3aed', '#0891b2', '#65a30d', '#be185d'];
        return colors[index % colors.length];
    }

    function svgText(x, y, text, size, anchor, fill) {
        var node = document.createElementNS(SVG_NS, 'text');
        node.setAttribute('x', String(x));
        node.setAttribute('y', String(y));
        node.setAttribute('font-size', String(size));
        node.setAttribute('text-anchor', anchor);
        node.setAttribute('fill', fill);
        node.textContent = String(text);
        return node;
    }

    function svgLine(x1, y1, x2, y2, stroke, width) {
        var node = document.createElementNS(SVG_NS, 'line');
        node.setAttribute('x1', String(x1));
        node.setAttribute('y1', String(y1));
        node.setAttribute('x2', String(x2));
        node.setAttribute('y2', String(y2));
        node.setAttribute('stroke', stroke);
        node.setAttribute('stroke-width', String(width));
        return node;
    }

    function svgRect(x, y, width, height, fill) {
        var node = document.createElementNS(SVG_NS, 'rect');
        node.setAttribute('x', String(x));
        node.setAttribute('y', String(y));
        node.setAttribute('width', String(width));
        node.setAttribute('height', String(height));
        node.setAttribute('fill', fill);
        return node;
    }

    function toPx(value) {
        return Math.max(0, Number(value || 0) / EMU_PER_PIXEL);
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
    global.OfficeRenderers.pptx = {
        render: render
    };
})(self);
