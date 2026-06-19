/**
 * Copyright (C) 2012-2013 KO GmbH <copyright@kogmbh.com>
 *
 * @licstart
 * This file is part of WebODF.
 *
 * WebODF is free software: you can redistribute it and/or modify it
 * under the terms of the GNU Affero General Public License (GNU AGPL)
 * as published by the Free Software Foundation, either version 3 of
 * the License, or (at your option) any later version.
 *
 * WebODF is distributed in the hope that it will be useful, but
 * WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU Affero General Public License for more details.
 *
 * You should have received a copy of the GNU Affero General Public License
 * along with WebODF.  If not, see <http://www.gnu.org/licenses/>.
 * @licend
 *
 * @source: http://www.webodf.org/
 * @source: https://github.com/kogmbh/WebODF/
 */

/*global runtime, odf, xmldom, webodf_css, core, gui */
/*jslint sub: true*/

(function () {
    "use strict";
    /**
     * A loading queue where various tasks related to loading can be placed
     * and will be run with 10 ms between them. This gives the ui a change to
     * to update.
     * @constructor
     */
    function LoadingQueue() {
        var /**@type{!Array.<!Function>}*/
            queue = [],
            taskRunning = false;
        /**
         * @param {!Function} task
         * @return {undefined}
         */
        function run(task) {
            taskRunning = true;
            runtime.setTimeout(function () {
                try {
                    task();
                } catch (/**@type{Error}*/e) {
                    runtime.log(String(e) + "\n" + e.stack);
                }
                taskRunning = false;
                if (queue.length > 0) {
                    run(queue.pop());
                }
            }, 10);
        }
        /**
         * @return {undefined}
         */
        this.clearQueue = function () {
            queue.length = 0;
        };
        /**
         * @param {!Function} loadingTask
         * @return {undefined}
         */
        this.addToQueue = function (loadingTask) {
            if (queue.length === 0 && !taskRunning) {
                return run(loadingTask);
            }
            queue.push(loadingTask);
        };
    }
    /**
     * @constructor
     * @implements {core.Destroyable}
     * @param {!HTMLStyleElement} css
     */
    function PageSwitcher(css) {
        var sheet = /**@type{!CSSStyleSheet}*/(css.sheet),
            /**@type{number}*/
            position = 1;
        /**
         * @return {undefined}
         */
        function updateCSS() {
            while (sheet.cssRules.length > 0) {
                sheet.deleteRule(0);
            }
            // The #shadowContent contains the master pages, with each page in the slideshow
            // corresponding to a master page in #shadowContent, and in the same order.
            // So, when showing a page, also make it's master page (behind it) visible.
            sheet.insertRule('#shadowContent draw|page {display:none;}', 0);
            sheet.insertRule('office|presentation draw|page {display:none;}', 1);
            sheet.insertRule("#shadowContent draw|page:nth-of-type(" +
                position + ") {display:block;}", 2);
            sheet.insertRule("office|presentation draw|page:nth-of-type(" +
                position + ") {display:block;}", 3);
        }
        /**
         * @return {undefined}
         */
        this.showFirstPage = function () {
            position = 1;
            updateCSS();
        };
        /**
         * @return {undefined}
         */
        this.showNextPage = function () {
            position += 1;
            updateCSS();
        };
        /**
         * @return {undefined}
         */
        this.showPreviousPage = function () {
            if (position > 1) {
                position -= 1;
                updateCSS();
            }
        };

        /**
         * @param {!number} n  number of the page
         * @return {undefined}
         */
        this.showPage = function (n) {
            if (n > 0) {
                position = n;
                updateCSS();
            }
        };

        this.css = css;

        /**
         * @param {!function(!Error=)} callback, passing an error object in case of error
         * @return {undefined}
         */
        this.destroy = function (callback) {
            css.parentNode.removeChild(css);
            callback();
        };
    }
    /**
     * Register event listener on DOM element.
     * @param {!Element} eventTarget
     * @param {!string} eventType
     * @param {!Function} eventHandler
     * @return {undefined}
     */
    function listenEvent(eventTarget, eventType, eventHandler) {
        if (eventTarget.addEventListener) {
            eventTarget.addEventListener(eventType, eventHandler, false);
        } else if (eventTarget.attachEvent) {
            eventType = "on" + eventType;
            eventTarget.attachEvent(eventType, eventHandler);
        } else {
            eventTarget["on" + eventType] = eventHandler;
        }
    }

    // variables per class (so not per instance!)
    var /**@const@type {!string}*/drawns  = odf.Namespaces.drawns,
        /**@const@type {!string}*/fons    = odf.Namespaces.fons,
        /**@const@type {!string}*/officens = odf.Namespaces.officens,
        /**@const@type {!string}*/stylens = odf.Namespaces.stylens,
        /**@const@type {!string}*/svgns   = odf.Namespaces.svgns,
        /**@const@type {!string}*/tablens = odf.Namespaces.tablens,
        /**@const@type {!string}*/chartns = odf.Namespaces.chartns,
        /**@const@type {!string}*/textns  = odf.Namespaces.textns,
        /**@const@type {!string}*/xlinkns = odf.Namespaces.xlinkns,
        /**@const@type {!string}*/presentationns = odf.Namespaces.presentationns,
        /**@const@type {!string}*/webodfhelperns = "urn:webodf:names:helper",
        xpath = xmldom.XPath,
        domUtils = core.DomUtils;

    /**
     * @param {!HTMLStyleElement} style
     * @return {undefined}
     */
    function clearCSSStyleSheet(style) {
        var stylesheet = /**@type{!CSSStyleSheet}*/(style.sheet),
            cssRules = stylesheet.cssRules;

        while (cssRules.length) {
            stylesheet.deleteRule(cssRules.length - 1);
        }
    }

    /**
     * A new styles.xml has been loaded. Update the live document with it.
     * @param {!odf.OdfContainer} odfcontainer
     * @param {!odf.Formatting} formatting
     * @param {!HTMLStyleElement} stylesxmlcss
     * @return {undefined}
     **/
    function handleStyles(odfcontainer, formatting, stylesxmlcss) {
        // update the css translation of the styles
        var style2css = new odf.Style2CSS(),
            list2css = new odf.ListStyleToCss(),
            styleSheet = /**@type{!CSSStyleSheet}*/(stylesxmlcss.sheet),
            styleTree = new odf.StyleTree(
                odfcontainer.rootElement.styles,
                odfcontainer.rootElement.automaticStyles).getStyleTree();

        style2css.style2css(
            odfcontainer.getDocumentType(),
            odfcontainer.rootElement,
            styleSheet,
            formatting.getFontMap(),
            styleTree
        );

        list2css.applyListStyles(
            styleSheet,
            styleTree,
            odfcontainer.rootElement.body);

    }

    /**
     * @param {!odf.OdfContainer} odfContainer
     * @param {!HTMLStyleElement} fontcss
     * @return {undefined}
     **/
    function handleFonts(odfContainer, fontcss) {
        // update the css references to the fonts
        var fontLoader = new odf.FontLoader();
        fontLoader.loadFonts(odfContainer,
            /**@type{!CSSStyleSheet}*/(fontcss.sheet));
    }

    /**
     * @param {!Element} clonedNode <draw:page/>
     * @return {undefined}
     */
    function dropTemplateDrawFrames(clonedNode) {
        // drop all frames which are just template frames
        var i, element, presentationClass,
            clonedDrawFrameElements = domUtils.getElementsByTagNameNS(clonedNode, drawns, 'frame');
        for (i = 0; i < clonedDrawFrameElements.length; i += 1) {
            element = /**@type{!Element}*/(clonedDrawFrameElements[i]);
            presentationClass = element.getAttributeNS(presentationns, 'class');
            if (presentationClass && ! /^(date-time|footer|header|page-number)$/.test(presentationClass)) {
                element.parentNode.removeChild(element);
            }
        }
    }

    /**
     * @param {!odf.OdfContainer} odfContainer
     * @param {!Element} frame
     * @param {!string} headerFooterId
     * @return {?string}
     */
    function getHeaderFooter(odfContainer, frame, headerFooterId) {
        var headerFooter = null,
            i,
            declElements = odfContainer.rootElement.body.getElementsByTagNameNS(presentationns, headerFooterId+'-decl'),
            headerFooterName = frame.getAttributeNS(presentationns, 'use-'+headerFooterId+'-name'),
            element;

        if (headerFooterName && declElements.length > 0) {
            for (i = 0; i < declElements.length; i += 1) {
                element = /**@type{!Element}*/(declElements[i]);
                if (element.getAttributeNS(presentationns, 'name') === headerFooterName) {
                    headerFooter = element.textContent;
                    break;
                }
            }
        }
        return headerFooter;
    }

    /**
     * @param {!Element} rootElement
     * @param {string} ns
     * @param {string} localName
     * @param {?string} value
     * @return {undefined}
     */
    function setContainerValue(rootElement, ns, localName, value) {
        var i, containerList,
            document = rootElement.ownerDocument,
            e;

        containerList = domUtils.getElementsByTagNameNS(rootElement, ns, localName);
        for (i = 0; i < containerList.length; i += 1) {
            domUtils.removeAllChildNodes(containerList[i]);
            if (value) {
                e = /**@type{!Element}*/(containerList[i]);
                e.appendChild(document.createTextNode(value));
            }
        }
    }

    /**
     * @param {string} styleid
     * @param {!Element} frame
     * @param {!CSSStyleSheet} stylesheet
     * @return {undefined}
     **/
    function setDrawElementPosition(styleid, frame, stylesheet) {
        frame.setAttributeNS(webodfhelperns, 'styleid', styleid);
        var rule,
            anchor = frame.getAttributeNS(textns, 'anchor-type'),
            x = frame.getAttributeNS(svgns, 'x'),
            y = frame.getAttributeNS(svgns, 'y'),
            width = frame.getAttributeNS(svgns, 'width'),
            height = frame.getAttributeNS(svgns, 'height'),
            minheight = frame.getAttributeNS(fons, 'min-height'),
            minwidth = frame.getAttributeNS(fons, 'min-width'),
            transform = frame.getAttributeNS(drawns, 'transform'),
            tm,
            angle,
            tx,
            ty;

        // A shape positioned with draw:transform carries no svg:x/y, so it would
        // otherwise get display:block and stack in the document flow at the left.
        // LibreOffice writes "rotate (a) translate (x y)": the shape's local
        // origin (top-left, pre-rotation) lands at (x, y) and is rotated about
        // that origin. ODF rotation is counter-clockwise, so negate it for CSS.
        if (!x && !y && transform) {
            tm = /translate\s*\(\s*([0-9.eE+-]+[a-z%]*)\s+([0-9.eE+-]+[a-z%]*)\s*\)/.exec(transform);
            angle = /rotate\s*\(\s*([0-9.eE+-]+)\s*\)/.exec(transform);
            if (tm) {
                tx = tm[1];
                ty = tm[2];
                rule = 'position: absolute;'
                    + 'left: ' + tx + ';'
                    + 'top: ' + ty + ';'
                    + 'transform-origin: 0 0;';
                if (angle) {
                    rule += 'transform: rotate('
                        + (-parseFloat(angle[1]) * 180 / Math.PI).toFixed(4) + 'deg);';
                }
                if (width) { rule += 'width: ' + width + ';'; }
                if (height) { rule += 'height: ' + height + ';'; }
                rule = 'draw|' + frame.localName + '[webodfhelper|styleid="' + styleid + '"] {'
                    + rule + '}';
                stylesheet.insertRule(rule, stylesheet.cssRules.length);
                return;
            }
        }

        if (anchor === "as-char") {
            rule = 'display: inline-block;';
        } else if (anchor || x || y) {
            rule = 'position: absolute;';
        } else if (width || height || minheight || minwidth) {
            rule = 'display: block;';
        }
        if (x) {
            rule += 'left: ' + x + ';';
        }
        if (y) {
            rule += 'top: ' + y + ';';
        }
        if (width) {
            rule += 'width: ' + width + ';';
        }
        if (height) {
            rule += 'height: ' + height + ';';
        }
        if (minheight) {
            rule += 'min-height: ' + minheight + ';';
        }
        if (minwidth) {
            rule += 'min-width: ' + minwidth + ';';
        }
        if (rule) {
            rule = 'draw|' + frame.localName + '[webodfhelper|styleid="' + styleid + '"] {' +
                rule + '}';
            stylesheet.insertRule(rule, stylesheet.cssRules.length);
        }
    }
    /**
     * @param {!Element} image
     * @return {string}
     **/
    function getUrlFromBinaryDataElement(image) {
        var node = image.firstChild;
        while (node) {
            if (node.namespaceURI === officens &&
                    node.localName === "binary-data") {
                // TODO: detect mime-type, assuming png for now
                // the base64 data can be  pretty printed, hence we need remove all the line breaks and whitespaces
                return "data:image/png;base64," + node.textContent.replace(/[\r\n\s]/g, '');
            }
            node = node.nextSibling;
        }
        return "";
    }

    /**
     * @param {!Uint8Array} data
     * @param {!number} offset
     * @return {!number}
     */
    function readInt16LE(data, offset) {
        var value = data[offset] | (data[offset + 1] << 8);
        return value & 0x8000 ? value - 0x10000 : value;
    }

    /**
     * @param {!Uint8Array} data
     * @param {!number} offset
     * @return {!number}
     */
    function readUInt16LE(data, offset) {
        return data[offset] | (data[offset + 1] << 8);
    }

    /**
     * @param {!Uint8Array} data
     * @param {!number} offset
     * @return {!number}
     */
    function readUInt32LE(data, offset) {
        return (data[offset] | (data[offset + 1] << 8) | (data[offset + 2] << 16)
            | (data[offset + 3] << 24)) >>> 0;
    }

    /**
     * @param {!Uint8Array} data
     * @param {!number} offset
     * @return {!string}
     */
    function readColorRef(data, offset) {
        return "rgb(" + data[offset] + "," + data[offset + 1] + "," + data[offset + 2] + ")";
    }

    /**
     * @param {!string} text
     * @return {!string}
     */
    function escapeXml(text) {
        return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;");
    }

    /**
     * Convert the subset of placeable WMF used by some PowerPoint-exported ODP
     * clipart to SVG so Chromium/WebView can display it.
     * @param {!Uint8Array} data
     * @return {?string}
     */
    function wmfToSvgUrl(data) {
        var PLACEABLE_KEY = 0x9AC6CDD7,
            META_EOF = 0x0000,
            META_SETWINDOWORG = 0x020B,
            META_SETWINDOWEXT = 0x020C,
            META_SELECTOBJECT = 0x012D,
            META_DELETEOBJECT = 0x01F0,
            META_CREATEPENINDIRECT = 0x02FA,
            META_CREATEBRUSHINDIRECT = 0x02FC,
            META_POLYGON = 0x0324,
            META_SETPOLYFILLMODE = 0x0106,
            offset = 0,
            viewLeft,
            viewTop,
            viewRight,
            viewBottom,
            viewWidth,
            viewHeight,
            windowOrgX,
            windowOrgY,
            windowExtX,
            windowExtY,
            svgWidth,
            svgHeight,
            objects = [],
            pen = {stroke: "none", width: 0},
            brush = {fill: "none"},
            fillRule = "evenodd",
            body = "",
            recordSize,
            fn,
            params,
            i,
            count,
            points,
            style,
            color,
            width,
            objectIndex;

        function addObject(o) {
            var i;
            for (i = 0; i < objects.length; i += 1) {
                if (!objects[i]) {
                    objects[i] = o;
                    return;
                }
            }
            objects.push(o);
        }

        /**
         * @param {!number} value
         * @return {!string}
         */
        function fmt(value) {
            return String(Math.round(value * 100) / 100);
        }

        /**
         * @param {!number} x
         * @return {!number}
         */
        function mapX(x) {
            return (x - windowOrgX) * svgWidth / windowExtX;
        }

        /**
         * @param {!number} y
         * @return {!number}
         */
        function mapY(y) {
            return (y - windowOrgY) * svgHeight / windowExtY;
        }

        if (data.length < 40 || readUInt32LE(data, 0) !== PLACEABLE_KEY) {
            return null;
        }
        viewLeft = readInt16LE(data, 6);
        viewTop = readInt16LE(data, 8);
        viewRight = readInt16LE(data, 10);
        viewBottom = readInt16LE(data, 12);
        viewWidth = Math.max(1, viewRight - viewLeft);
        viewHeight = Math.max(1, viewBottom - viewTop);
        windowOrgX = viewLeft;
        windowOrgY = viewTop;
        windowExtX = viewWidth;
        windowExtY = viewHeight;
        svgWidth = viewWidth;
        svgHeight = viewHeight;
        offset = 22 + 18;
        while (offset + 6 <= data.length) {
            recordSize = readUInt32LE(data, offset);
            fn = readUInt16LE(data, offset + 4);
            params = offset + 6;
            if (recordSize < 3 || offset + recordSize * 2 > data.length) {
                return null;
            }
            if (fn === META_EOF) {
                break;
            }
            if (fn === META_CREATEBRUSHINDIRECT) {
                style = readUInt16LE(data, params);
                color = readColorRef(data, params + 2);
                addObject({type: "brush", fill: style === 0 ? color : "none"});
            } else if (fn === META_CREATEPENINDIRECT) {
                style = readUInt16LE(data, params);
                width = Math.abs(readInt16LE(data, params + 2));
                color = readColorRef(data, params + 6);
                addObject({type: "pen", stroke: style === 5 ? "none" : color, width: Math.max(1, width)});
            } else if (fn === META_SELECTOBJECT) {
                objectIndex = readUInt16LE(data, params);
                if (objects[objectIndex]) {
                    if (objects[objectIndex].type === "pen") {
                        pen = objects[objectIndex];
                    } else if (objects[objectIndex].type === "brush") {
                        brush = objects[objectIndex];
                    }
                }
            } else if (fn === META_DELETEOBJECT) {
                objects[readUInt16LE(data, params)] = null;
            } else if (fn === META_SETPOLYFILLMODE) {
                fillRule = readUInt16LE(data, params) === 2 ? "nonzero" : "evenodd";
            } else if (fn === META_SETWINDOWORG) {
                windowOrgY = readInt16LE(data, params);
                windowOrgX = readInt16LE(data, params + 2);
            } else if (fn === META_SETWINDOWEXT) {
                windowExtY = readInt16LE(data, params);
                windowExtX = readInt16LE(data, params + 2);
                if (windowExtX === 0 || windowExtY === 0) {
                    return null;
                }
                svgWidth = Math.abs(windowExtX);
                svgHeight = Math.abs(windowExtY);
            } else if (fn === META_POLYGON) {
                count = readUInt16LE(data, params);
                if (params + 2 + count * 4 > offset + recordSize * 2) {
                    return null;
                }
                points = [];
                for (i = 0; i < count; i += 1) {
                    points.push(fmt(mapX(readInt16LE(data, params + 2 + i * 4))) + ","
                        + fmt(mapY(readInt16LE(data, params + 4 + i * 4))));
                }
                body += '<polygon points="' + points.join(" ") + '" fill="' + escapeXml(brush.fill)
                    + '" stroke="' + escapeXml(pen.stroke) + '" stroke-width="' + pen.width
                    + '" fill-rule="' + fillRule + '" stroke-linejoin="round"/>';
            }
            offset += recordSize * 2;
        }
        if (!body) {
            return null;
        }
        return "data:image/svg+xml;charset=utf-8," + encodeURIComponent(
            '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ' + svgWidth
            + " " + svgHeight + '">' + body + "</svg>");
    }

    /**
     * @param {!string} url
     * @return {?string}
     */
    function wmfDataUrlToSvgUrl(url) {
        var base64Index = url.indexOf("base64,"),
            binary,
            data,
            i;
        if (base64Index < 0 || typeof atob !== "function") {
            return null;
        }
        binary = atob(url.substring(base64Index + 7));
        data = new Uint8Array(binary.length);
        for (i = 0; i < binary.length; i += 1) {
            data[i] = binary.charCodeAt(i) & 0xff;
        }
        return wmfToSvgUrl(data);
    }

    /**
     * @param {!string} url
     * @return {!string}
     */
    function cssUrl(url) {
        return 'url("' + url.replace(/\\/g, "\\\\").replace(/"/g, '\\"')
            .replace(/\r/g, "").replace(/\n/g, "") + '")';
    }

    /**
     * @param {string} id
     * @param {!odf.OdfContainer} container
     * @param {!Element} image
     * @param {!CSSStyleSheet} stylesheet
     * @return {undefined}
     **/
    function setImage(id, container, image, stylesheet) {
        image.setAttributeNS(webodfhelperns, 'styleid', id);
        var url = image.getAttributeNS(xlinkns, 'href'),
            /**@type{!odf.OdfPart}*/
            part;
        /**
         * @param {?string} url
         */
        function callback(url) {
            var rule;
            if (url) { // if part cannot be loaded, url is null
                rule = "background-image: " + cssUrl(url) + ";";
                rule = 'draw|image[webodfhelper|styleid="' + id + '"] {' + rule + '}';
                try {
                    stylesheet.insertRule(rule, stylesheet.cssRules.length);
                } catch (/**@type{*}*/e) {
                    runtime.log('image css problem: ' + String(e));
                }
            }
        }
        /**
         * @param {!odf.OdfPart} p
         */
        function onchange(p) {
            var convertedUrl = null;
            try {
                convertedUrl = /\.wmf$/i.test(p.name) && p.url ? wmfDataUrlToSvgUrl(p.url) : null;
            } catch (/**@type{*}*/e) {
                runtime.log('wmf conversion problem: ' + String(e));
            }
            callback(convertedUrl || p.url);
        }
        // look for a office:binary-data
        if (url) {
            try {
                part = container.getPart(url);
                part.onchange = onchange;
                part.load();
            } catch (/**@type{*}*/e) {
                runtime.log('slight problem: ' + String(e));
            }
        } else {
            url = getUrlFromBinaryDataElement(image);
            callback(url);
        }
    }
    /**
     * Find a <style:style> of the drawing-page family by name, searching both
     * the document's automatic styles and the common styles.
     * @param {!Element} rootElement  the ODF root element
     * @param {?string} name
     * @return {?Element}
     */
    function findDrawingPageStyle(rootElement, name) {
        var roots = [rootElement.automaticStyles, rootElement.styles],
            i,
            node;
        if (!name) {
            return null;
        }
        for (i = 0; i < roots.length; i += 1) {
            node = roots[i] && roots[i].firstElementChild;
            while (node) {
                if (node.namespaceURI === stylens && node.localName === "style"
                        && node.getAttributeNS(stylens, "family") === "drawing-page"
                        && node.getAttributeNS(stylens, "name") === name) {
                    return /**@type{!Element}*/(node);
                }
                node = node.nextElementSibling;
            }
        }
        return null;
    }
    /**
     * Resolve a <draw:fill-image> reference to the href of the image part.
     * @param {!Element} rootElement  the ODF root element
     * @param {?string} name  value of draw:fill-image-name
     * @return {?string}
     */
    function findFillImageHref(rootElement, name) {
        var images,
            i;
        if (!name) {
            return null;
        }
        images = domUtils.getElementsByTagNameNS(rootElement.styles, drawns, "fill-image");
        for (i = 0; i < images.length; i += 1) {
            if (images[i].getAttributeNS(drawns, "name") === name) {
                return images[i].getAttributeNS(xlinkns, "href");
            }
        }
        return null;
    }
    /**
     * Find a graphic-family <style:style> by name (automatic or common styles).
     * @param {!Element} rootElement
     * @param {?string} name
     * @return {?Element}
     */
    function findGraphicStyle(rootElement, name) {
        var roots = [rootElement.automaticStyles, rootElement.styles],
            i,
            node;
        if (!name) {
            return null;
        }
        for (i = 0; i < roots.length; i += 1) {
            node = roots[i] && roots[i].firstElementChild;
            while (node) {
                if (node.namespaceURI === stylens && node.localName === "style"
                        && node.getAttributeNS(stylens, "family") === "graphic"
                        && node.getAttributeNS(stylens, "name") === name) {
                    return /**@type{!Element}*/(node);
                }
                node = node.nextElementSibling;
            }
        }
        return null;
    }

    /**
     * @param {!Element} rootElement
     * @param {?string} styleName
     * @param {!string} propertyName
     * @return {?string}
     */
    function findDrawingPagePresentationProperty(rootElement, styleName, propertyName) {
        var style = findDrawingPageStyle(rootElement, styleName),
            props,
            value;
        while (style) {
            props = domUtils.getDirectChild(style, stylens, "drawing-page-properties");
            if (props) {
                value = props.getAttributeNS(presentationns, propertyName);
                if (value) {
                    return value;
                }
            }
            style = findDrawingPageStyle(rootElement,
                style.getAttributeNS(stylens, "parent-style-name"));
        }
        return null;
    }

    /**
     * @param {!Element} rootElement
     * @param {!Element} pageElement
     * @param {!Element} masterPageElement
     * @param {!Element} element
     * @return {!boolean}
     */
    function shouldCloneMasterPageElement(rootElement, pageElement, masterPageElement, element) {
        var presentationClass = element.getAttributeNS(presentationns, "class"),
            displayProperty,
            pageValue,
            masterValue,
            pageStyleName,
            masterStyleName;
        if (element.getAttributeNS(presentationns, "placeholder") === "true") {
            return false;
        }
        if (!presentationClass) {
            return true;
        }
        if (presentationClass === "page-number") {
            displayProperty = "display-page-number";
        } else if (presentationClass === "date-time") {
            displayProperty = "display-date-time";
        } else if (presentationClass === "header") {
            displayProperty = "display-header";
        } else if (presentationClass === "footer") {
            displayProperty = "display-footer";
        } else {
            return true;
        }
        pageStyleName = pageElement.getAttributeNS(drawns, "style-name");
        pageValue = findDrawingPagePresentationProperty(rootElement, pageStyleName, displayProperty);
        if (pageValue === "false") {
            return false;
        }
        if (pageValue === "true") {
            return true;
        }
        masterStyleName = masterPageElement.getAttributeNS(drawns, "style-name");
        masterValue = findDrawingPagePresentationProperty(rootElement, masterStyleName, displayProperty);
        return masterValue !== "false";
    }
    /**
     * Resolve one attribute from a graphic style, following parent-style-name.
     * @param {!Element} rootElement
     * @param {?string} styleName
     * @param {string} ns
     * @param {string} attr
     * @return {?string}
     */
    function resolveGraphicProperty(rootElement, styleName, ns, attr) {
        var name = styleName,
            depth = 0,
            style,
            gp,
            value;
        while (name && depth < 16) {
            style = findGraphicStyle(rootElement, name);
            if (!style) {
                return null;
            }
            gp = domUtils.getElementsByTagNameNS(style, stylens, "graphic-properties")[0];
            if (gp) {
                value = gp.getAttributeNS(ns, attr);
                if (value) {
                    return value;
                }
            }
            name = style.getAttributeNS(stylens, "parent-style-name");
            depth += 1;
        }
        return null;
    }
    /**
     * Find a draw:marker definition by name.
     * @param {!Element} rootElement
     * @param {?string} name
     * @return {?Element}
     */
    function findDrawMarker(rootElement, name) {
        var roots = [rootElement.automaticStyles, rootElement.styles],
            i,
            markers,
            j;
        if (!name) {
            return null;
        }
        for (i = 0; i < roots.length; i += 1) {
            markers = roots[i] ? domUtils.getElementsByTagNameNS(roots[i], drawns, "marker") : [];
            for (j = 0; j < markers.length; j += 1) {
                if (markers[j].getAttributeNS(drawns, "name") === name) {
                    return /**@type{!Element}*/(markers[j]);
                }
            }
        }
        return null;
    }
    /**
     * Resolve the effective fo:clip of a graphic style (following
     * parent-style-name). fo:clip is otherwise dropped, so cropped images
     * (a common case for "fill frame" pictures) render uncropped.
     * @param {!Element} rootElement
     * @param {?string} styleName
     * @return {?string}
     */
    function resolveGraphicClip(rootElement, styleName) {
        var name = styleName,
            depth = 0,
            style,
            gp,
            clip;
        while (name && depth < 16) {
            style = findGraphicStyle(rootElement, name);
            if (!style) {
                return null;
            }
            gp = domUtils.getElementsByTagNameNS(style, stylens, "graphic-properties")[0];
            if (gp) {
                clip = gp.getAttributeNS(fons, "clip");
                if (clip) {
                    return clip;
                }
            }
            name = style.getAttributeNS(stylens, "parent-style-name");
            depth += 1;
        }
        return null;
    }
    /**
     * Apply an image crop (fo:clip) to a draw:image. The frame's svg:width/height
     * is the visible (cropped) size; fo:clip gives the inset cut from each edge.
     * Reproduce it with background-size/position so only the kept region shows.
     * @param {!Element} image
     * @param {!string} id
     * @param {!odf.OdfContainer} container
     * @param {!CSSStyleSheet} stylesheet
     * @return {undefined}
     */
    function applyImageClip(image, id, container, stylesheet) {
        var frame = image.parentNode,
            clip,
            m,
            parts,
            top,
            right,
            bottom,
            left,
            fw,
            fh,
            unit,
            rule;
        if (!frame || frame.localName !== "frame") {
            return;
        }
        clip = resolveGraphicClip(container.rootElement,
            frame.getAttributeNS(drawns, "style-name"));
        if (!clip) {
            return;
        }
        m = /rect\(([^)]*)\)/.exec(clip);
        if (!m) {
            return;
        }
        parts = m[1].split(",");
        if (parts.length < 4) {
            return;
        }
        top = parseLength(parts[0]);
        right = parseLength(parts[1]);
        bottom = parseLength(parts[2]);
        left = parseLength(parts[3]);
        if (top.v === 0 && right.v === 0 && bottom.v === 0 && left.v === 0) {
            return; // no actual crop
        }
        fw = parseLength(frame.getAttributeNS(svgns, "width"));
        fh = parseLength(frame.getAttributeNS(svgns, "height"));
        if (!fw.v || !fh.v) {
            return;
        }
        unit = fw.u || "cm";
        image.setAttributeNS(webodfhelperns, "clipid", id);
        rule = 'draw|image[webodfhelper|clipid="' + id + '"] {'
            + "background-size: " + (fw.v + left.v + right.v) + unit + " "
                + (fh.v + top.v + bottom.v) + unit + ";"
            + "background-position: " + (-left.v) + unit + " " + (-top.v) + unit + ";"
            + "}";
        stylesheet.insertRule(rule, stylesheet.cssRules.length);
    }
    /**
     * Load the image part and apply it as the background of every draw:page that
     * uses the given drawing-page style. Mirrors setImage: the part is loaded
     * asynchronously and the CSS rule is added once its data URL is available.
     * The rule reuses the same selector that Style2CSS generates for the
     * drawing-page style, but lives in the (later) position stylesheet so it wins
     * over the "background: none" that Style2CSS emits when the page background
     * is visible.
     * @param {!odf.OdfContainer} container
     * @param {!string} styleName
     * @param {!string} href
     * @param {?string} repeat  value of style:repeat
     * @param {!CSSStyleSheet} stylesheet
     * @return {undefined}
     */
    function setPageBackgroundForSelector(container, selector, href, repeat, stylesheet) {
        var part;
        /**
         * @param {?string} url
         */
        function callback(url) {
            var rule;
            if (!url) { // if part cannot be loaded, url is null
                return;
            }
            rule = selector + " {"
                + "background-image: url(" + url + ");"
                + "background-repeat: " + (repeat === "repeat" ? "repeat" : "no-repeat") + ";";
            if (repeat === "stretch") {
                // ODF "stretch" scales the bitmap to fill the page.
                rule += "background-size: 100% 100%;";
            }
            rule += "}";
            stylesheet.insertRule(rule, stylesheet.cssRules.length);
        }
        try {
            part = container.getPart(href);
            part.onchange = function (p) {
                callback(p.url);
            };
            part.load();
        } catch (/**@type{*}*/e) {
            runtime.log('slight problem: ' + String(e));
        }
    }
    /**
     * @param {!odf.OdfContainer} container
     * @param {!string} styleName
     * @param {!string} href
     * @param {?string} repeat  value of style:repeat
     * @param {!CSSStyleSheet} stylesheet
     * @return {undefined}
     */
    function setPageBackground(container, styleName, href, repeat, stylesheet) {
        setPageBackgroundForSelector(container,
            'draw|page[draw|style-name="' + styleName + '"]', href, repeat, stylesheet);
    }
    /**
     * Render bitmap page fills (draw:fill="bitmap") as draw:page backgrounds.
     * WebODF otherwise ignores bitmap fills, but presentations are commonly
     * exported with each slide as a single full-bleed background image, which
     * would then render blank.
     * @param {!odf.OdfContainer} container
     * @param {!Element} odfbody
     * @param {!CSSStyleSheet} stylesheet
     * @return {undefined}
     */
    function loadPageBackgrounds(container, odfbody, stylesheet) {
        var rootElement = /**@type{!Element}*/(container.rootElement),
            pages = domUtils.getElementsByTagNameNS(odfbody, drawns, "page"),
            seen = {},
            i;
        for (i = 0; i < pages.length; i += 1) {
            (function (styleName) {
                var style, props = null, href, repeat;
                if (!styleName || seen.hasOwnProperty(styleName)) {
                    return;
                }
                seen[styleName] = true;
                // Walk the parent-style chain until a fill is specified.
                style = findDrawingPageStyle(rootElement, styleName);
                while (style) {
                    props = domUtils.getDirectChild(style, stylens, "drawing-page-properties");
                    if (props && props.getAttributeNS(drawns, "fill")) {
                        break;
                    }
                    props = null;
                    style = findDrawingPageStyle(rootElement,
                        style.getAttributeNS(stylens, "parent-style-name"));
                }
                if (!props || props.getAttributeNS(drawns, "fill") !== "bitmap") {
                    return;
                }
                href = findFillImageHref(rootElement,
                    props.getAttributeNS(drawns, "fill-image-name"));
                if (!href) {
                    return;
                }
                repeat = props.getAttributeNS(stylens, "repeat");
                setPageBackground(container, styleName, href, repeat, stylesheet);
            }(pages[i].getAttributeNS(drawns, "style-name")));
        }
    }
    /**
     * Apply each master page's own drawing-page fill to its rendered clone in
     * #shadowContent. WebODF's default graphic style paints every draw:page with
     * the template's default shape fill, and the master clones carry no
     * drawing-page style of their own, so without this the master background
     * shows that stray fill (e.g. a cyan deck default) instead of the master's
     * real background (e.g. white). Scoped to #shadowContent so it only affects
     * the master layer, never the slide page (which must stay transparent so the
     * master background objects remain visible).
     * @param {!odf.OdfContainer} container
     * @param {!CSSStyleSheet} stylesheet
     * @return {undefined}
     */
    function loadMasterPageBackgrounds(container, stylesheet) {
        var rootElement = /**@type{!Element}*/(container.rootElement),
            masterStyles = rootElement.masterStyles,
            master = masterStyles && masterStyles.firstElementChild,
            selectorPrefix = '#shadowContent draw|page[draw|master-page-name="';
        while (master) {
            if (master.namespaceURI === stylens && master.localName === "master-page") {
                (function (name, drawStyleName) {
                    var style = findDrawingPageStyle(rootElement, drawStyleName),
                        props = null,
                        fill,
                        color,
                        selector = selectorPrefix + name + '"]';
                    while (style) {
                        props = domUtils.getDirectChild(style, stylens, "drawing-page-properties");
                        if (props && props.getAttributeNS(drawns, "fill")) {
                            break;
                        }
                        props = null;
                        style = findDrawingPageStyle(rootElement,
                            style.getAttributeNS(stylens, "parent-style-name"));
                    }
                    if (!props) {
                        return;
                    }
                    fill = props.getAttributeNS(drawns, "fill");
                    if (fill === "solid") {
                        color = props.getAttributeNS(drawns, "fill-color");
                        if (color) {
                            stylesheet.insertRule(selector + " {background-color: " + color
                                + "; background-image: none;}", stylesheet.cssRules.length);
                        }
                    } else if (fill === "none") {
                        stylesheet.insertRule(selector + " {background: none;}", stylesheet.cssRules.length);
                    }
                    // Bitmap master fills are intentionally not painted here: the
                    // slide's own draw:page already carries its fill (see
                    // loadPageBackgrounds), and painting the master clone's bitmap
                    // as well makes the two overlap and ghost.
                }(master.getAttributeNS(stylens, "name"), master.getAttributeNS(drawns, "style-name")));
            }
            master = master.nextElementSibling;
        }
    }

    /**
     * @param {!odf.OdfContainer} container
     * @param {!CSSStyleSheet} stylesheet
     * @return {undefined}
     */
    function loadHiddenDrawLayers(container, stylesheet) {
        var rootElement = /**@type{!Element}*/(container.rootElement),
            roots = [rootElement.styles, rootElement.automaticStyles, rootElement.masterStyles],
            i,
            layers,
            j,
            layerName;
        for (i = 0; i < roots.length; i += 1) {
            if (!roots[i]) {
                continue;
            }
            layers = domUtils.getElementsByTagNameNS(roots[i], drawns, "layer");
            for (j = 0; j < layers.length; j += 1) {
                if (layers[j].getAttributeNS(drawns, "display") === "none") {
                    layerName = layers[j].getAttributeNS(drawns, "name");
                    if (layerName) {
                        stylesheet.insertRule('[draw|layer="' + layerName.replace(/\\/g, "\\\\").replace(/"/g, '\\"')
                            + '"] {display: none !important; visibility: hidden !important;}',
                            stylesheet.cssRules.length);
                    }
                }
            }
        }
    }

    /**
     * Build an evaluation context for the formulas (draw:equation) and modifiers
     * (draw:modifiers) of a draw:enhanced-geometry, exposing a resolver for the
     * predefined identifiers, the modifiers ($N) and the named equations (?fN).
     * @param {!Element} geometry  <draw:enhanced-geometry>
     * @return {!{evaluate:function(!string):!number,viewBox:!Array.<!number>}}
     */
    function createShapeContext(geometry) {
        var viewBoxAttr = (geometry.getAttributeNS(svgns, "viewBox") || "0 0 21600 21600").trim().split(/\s+/),
            vb = [parseFloat(viewBoxAttr[0]) || 0, parseFloat(viewBoxAttr[1]) || 0,
                parseFloat(viewBoxAttr[2]) || 21600, parseFloat(viewBoxAttr[3]) || 21600],
            modifiersAttr = (geometry.getAttributeNS(drawns, "modifiers") || "").trim(),
            modifiers = modifiersAttr ? modifiersAttr.split(/\s+/).map(parseFloat) : [],
            equations = {},
            memo = {},
            inProgress = {},
            identifiers = {
                width: vb[2], height: vb[3], left: vb[0], top: vb[1],
                right: vb[0] + vb[2], bottom: vb[1] + vb[3],
                logwidth: vb[2], logheight: vb[3],
                xstretch: 0, ystretch: 0, hasstroke: 1, hasfill: 1, pi: Math.PI
            },
            node = geometry.firstElementChild,
            ctx = {};

        while (node) {
            if (node.namespaceURI === drawns && node.localName === "equation") {
                equations[node.getAttributeNS(drawns, "name")] = node.getAttributeNS(drawns, "formula") || "0";
            }
            node = node.nextElementSibling;
        }

        /**
         * @param {!Array.<!string>} tokens
         * @return {!number}
         */
        function parseExpression(tokens) {
            var pos = 0;

            function peek() { return tokens[pos]; }
            function next() { pos += 1; return tokens[pos - 1]; }

            function parsePrimary() {
                var token = next(), value, args, name;
                if (token === "(") {
                    value = parseAddition();
                    next(); // ")"
                    return value;
                }
                if (token === "-") {
                    return -parsePrimary();
                }
                if (token === "+") {
                    return parsePrimary();
                }
                // function call: identifier followed by "("
                if (/^[a-z]+$/.test(token) && peek() === "(") {
                    name = token;
                    next(); // "("
                    args = [parseAddition()];
                    while (peek() === ",") {
                        next();
                        args.push(parseAddition());
                    }
                    next(); // ")"
                    switch (name) {
                    case "abs": return Math.abs(args[0]);
                    case "sqrt": return Math.sqrt(args[0]);
                    case "sin": return Math.sin(args[0]);
                    case "cos": return Math.cos(args[0]);
                    case "tan": return Math.tan(args[0]);
                    case "atan": return Math.atan(args[0]);
                    case "atan2": return Math.atan2(args[0], args[1]);
                    case "min": return Math.min(args[0], args[1]);
                    case "max": return Math.max(args[0], args[1]);
                    case "mod": return Math.sqrt(args[0] * args[0] + args[1] * args[1] + args[2] * args[2]);
                    case "if": return args[0] > 0 ? args[1] : args[2];
                    default: return 0;
                    }
                }
                return resolve(token);
            }

            function parseMultiplication() {
                var value = parsePrimary(), op;
                while (peek() === "*" || peek() === "/") {
                    op = next();
                    if (op === "*") {
                        value *= parsePrimary();
                    } else {
                        value /= parsePrimary();
                    }
                }
                return value;
            }

            function parseAdditionInner() {
                var value = parseMultiplication(), op;
                while (peek() === "+" || peek() === "-") {
                    op = next();
                    if (op === "+") {
                        value += parseMultiplication();
                    } else {
                        value -= parseMultiplication();
                    }
                }
                return value;
            }

            // hoisted reference so parsePrimary can recurse into the top rule
            function parseAddition() { return parseAdditionInner(); }

            return parseAddition();
        }

        /**
         * Resolve a single value token to a number.
         * @param {!string} token
         * @return {!number}
         */
        function resolve(token) {
            var value;
            if (token === undefined) {
                return 0;
            }
            if (token.charAt(0) === "?") {
                return evaluateEquation(token.substr(1));
            }
            if (token.charAt(0) === "$") {
                value = modifiers[parseInt(token.substr(1), 10)];
                return isNaN(value) ? 0 : value;
            }
            if (identifiers.hasOwnProperty(token)) {
                return identifiers[token];
            }
            value = parseFloat(token);
            return isNaN(value) ? 0 : value;
        }

        /**
         * Tokenize a formula/path value expression.
         * @param {!string} expression
         * @return {!Array.<!string>}
         */
        function tokenize(expression) {
            var tokens = expression.match(/\?[a-zA-Z0-9]+|\$[0-9]+|[a-z]+|[0-9]*\.?[0-9]+|[()+\-*/,]/g);
            return tokens || [];
        }

        /**
         * @param {!string} name
         * @return {!number}
         */
        function evaluateEquation(name) {
            var result;
            if (memo.hasOwnProperty(name)) {
                return memo[name];
            }
            if (inProgress[name] || !equations.hasOwnProperty(name)) {
                return 0; // cycle or missing reference
            }
            inProgress[name] = true;
            result = parseExpression(tokenize(equations[name]));
            inProgress[name] = false;
            memo[name] = result;
            return result;
        }

        ctx.viewBox = vb;
        // The stretch point keeps a shape's corners a fixed size while only the
        // straight edges between them grow when the shape's box is not square
        // (e.g. a rounded-rectangle "pill"). NaN when not specified.
        ctx.stretchX = parseFloat(geometry.getAttributeNS(drawns, "path-stretchpoint-x"));
        ctx.stretchY = parseFloat(geometry.getAttributeNS(drawns, "path-stretchpoint-y"));
        ctx.evaluate = function (token) {
            // A path value is a single token (number, ?fN, $N or identifier).
            return resolve(token);
        };
        return ctx;
    }
    /**
     * Map enhanced-path points so that, when the SVG is stretched to a non-square
     * box, the corners around the stretch point keep their (circular) size and
     * the straight edges absorb the extra length. Only handles paths built from
     * straight/quadrant commands (M/L/X/Y); returns null for anything else (arcs,
     * curves) so the caller falls back to a plain stretch.
     * @param {!Array.<!string>} tokens
     * @param {!{evaluate:function(!string):!number}} ctx
     * @param {!number} stretchPoint  coordinate of the stretch point on the axis
     * @param {!number} viewBoxSize  the axis length of the viewBox
     * @param {!number} targetSize  the widened axis length (matches the box aspect)
     * @param {!boolean} horizontal  stretch the x axis (else the y axis)
     * @return {?Array.<!{x:!number,y:!number}>}
     */
    function stretchedPoints(tokens, ctx, stretchPoint, viewBoxSize, targetSize, horizontal) {
        var commandRe = /^[MLCZNFSTUABWVXYQ]$/,
            allowed = {M: 1, L: 1, X: 1, Y: 1, Z: 1, N: 1, F: 1, S: 1},
            raw = [],
            pos = 0,
            n,
            i,
            cur = null,
            prevOnPoint = false,
            offset = targetSize - viewBoxSize,
            cmd,
            coord;
        while (pos < tokens.length) {
            cmd = tokens[pos];
            if (!commandRe.test(cmd)) { pos += 1; continue; }
            if (!allowed[cmd]) { return null; }
            pos += 1;
            while (pos < tokens.length && !commandRe.test(tokens[pos])) {
                raw.push({ x: ctx.evaluate(tokens[pos]), y: ctx.evaluate(tokens[pos + 1]) });
                pos += 2;
            }
        }
        n = raw.length;
        // Decide which side of the stretch point each coordinate belongs to. A
        // point exactly on the stretch point inherits the running side, except a
        // run of two consecutive on-point coordinates is the zero-length segment
        // that becomes the stretched straight edge, so the side flips there.
        for (i = 0; i < n; i += 1) {
            coord = horizontal ? raw[i].x : raw[i].y;
            if (coord < stretchPoint - 0.5) {
                cur = "L";
                prevOnPoint = false;
            } else if (coord > stretchPoint + 0.5) {
                cur = "R";
                prevOnPoint = false;
            } else {
                if (cur === null) {
                    // First point: take the side of the next off-point coordinate.
                    cur = "L";
                    for (var j = i + 1; j < n; j += 1) {
                        var c = horizontal ? raw[j].x : raw[j].y;
                        if (c < stretchPoint - 0.5) { cur = "L"; break; }
                        if (c > stretchPoint + 0.5) { cur = "R"; break; }
                    }
                } else if (prevOnPoint) {
                    cur = cur === "L" ? "R" : "L";
                }
                prevOnPoint = true;
            }
            if (cur === "R") {
                if (horizontal) { raw[i].x += offset; } else { raw[i].y += offset; }
            }
        }
        return raw;
    }
    /**
     * Append SVG arc command(s) tracing an elliptical segment to the path data.
     * Splits arcs of 180 degrees or more so a single SVG arc never has an
     * ambiguous large-arc/sweep combination, and handles full ellipses.
     * @param {!Array.<!string>} d  path-data fragments (mutated)
     * @param {!number} cx
     * @param {!number} cy
     * @param {!number} rx
     * @param {!number} ry
     * @param {!number} startAngle  radians
     * @param {!number} endAngle  radians
     * @param {!boolean} connectWithMove  start a new subpath instead of a line
     * @return {undefined}
     */
    function appendEllipticalArc(d, cx, cy, rx, ry, startAngle, endAngle, connectWithMove) {
        var TWO_PI = Math.PI * 2,
            sweep = endAngle - startAngle,
            steps,
            i,
            a0,
            a1,
            x0 = cx + rx * Math.cos(startAngle),
            y0 = cy + ry * Math.sin(startAngle),
            x1,
            y1,
            sweepFlag = sweep >= 0 ? 1 : 0,
            absSweep = Math.abs(sweep);
        d.push((connectWithMove ? "M" : "L") + x0.toFixed(2) + " " + y0.toFixed(2));
        // Break the arc into pieces no larger than (just under) 180 degrees.
        steps = Math.max(1, Math.ceil(absSweep / (Math.PI - 0.001)));
        for (i = 1; i <= steps; i += 1) {
            a0 = startAngle + sweep * (i - 1) / steps;
            a1 = startAngle + sweep * i / steps;
            x1 = cx + rx * Math.cos(a1);
            y1 = cy + ry * Math.sin(a1);
            d.push("A" + rx.toFixed(2) + " " + ry.toFixed(2) + " 0 0 " + sweepFlag + " "
                + x1.toFixed(2) + " " + y1.toFixed(2));
        }
    }
    /**
     * Translate a draw:enhanced-path into one or more SVG sub-paths. Each entry
     * has its own fill/stroke flags (the path "F"/"S" commands disable fill and
     * stroke for the sub-paths that follow within the same geometry).
     * @param {!string} path
     * @param {!{evaluate:function(!string):!number}} ctx
     * @param {?Array.<!{x:!number,y:!number}>=} mappedPts  pre-computed,
     *     stretch-corrected points consumed in order for M/L/X/Y commands
     * @return {!Array.<!{d:!string,fill:!boolean,stroke:!boolean}>}
     */
    function parseEnhancedPath(path, ctx, mappedPts) {
        var tokens = path.match(/[A-Za-z]|\?[a-zA-Z0-9]+|\$[0-9]+|[a-z]+|-?[0-9]*\.?[0-9]+/g) || [],
            commandRe = /^[MLCZNFSTUABWVXYQ]$/,
            subPaths = [],
            d = [],
            fill = true,
            stroke = true,
            curX = 0,
            curY = 0,
            pos = 0,
            ptIdx = 0,
            command,
            count = 0;

        function flush() {
            if (d.length) {
                subPaths.push({ d: d.join(" "), fill: fill, stroke: stroke });
            }
            d = [];
        }
        function val() {
            var t = tokens[pos];
            pos += 1;
            return ctx.evaluate(t);
        }
        function hasValue() {
            return pos < tokens.length && !commandRe.test(tokens[pos]);
        }
        // Read one (x,y) point: from the stretch-corrected list when present,
        // otherwise straight from the token stream.
        function nextPoint() {
            var p;
            if (mappedPts) {
                p = mappedPts[ptIdx];
                ptIdx += 1;
                pos += 2;
                return [p.x, p.y];
            }
            return [val(), val()];
        }
        // Elliptical-quadrant helper (X/Y). Draws a 90-degree arc from the
        // current point to (tx,ty); axis is whichever of the two the tangent
        // starts along.
        function quadrant(tx, ty, startAlongX) {
            // The radii span the rectangle between the current point and the
            // target; only the centre differs between X (tangent starts
            // horizontal) and Y (tangent starts vertical).
            var cx = startAlongX ? curX : tx,
                cy = startAlongX ? ty : curY,
                rx = Math.abs(tx - curX),
                ry = Math.abs(ty - curY),
                a0 = Math.atan2((curY - cy) / (ry || 1), (curX - cx) / (rx || 1)),
                a1 = Math.atan2((ty - cy) / (ry || 1), (tx - cx) / (rx || 1)),
                diff = a1 - a0;
            // Choose the short way round (a quadrant is always 90 degrees).
            if (diff > Math.PI) { a1 -= 2 * Math.PI; }
            if (diff < -Math.PI) { a1 += 2 * Math.PI; }
            appendEllipticalArc(d, cx, cy, rx, ry, a0, a1, false);
            curX = tx;
            curY = ty;
        }
        // Arc-to helper (A/W). Ellipse bounding box (x1,y1)-(x2,y2); the segment
        // runs from the direction of (x3,y3) to that of (x4,y4).
        function arcTo(x1, y1, x2, y2, x3, y3, x4, y4, clockwise, connectWithMove) {
            var cx = (x1 + x2) / 2,
                cy = (y1 + y2) / 2,
                rx = Math.abs(x2 - x1) / 2,
                ry = Math.abs(y2 - y1) / 2,
                a0 = Math.atan2((y3 - cy) / (ry || 1), (x3 - cx) / (rx || 1)),
                a1 = Math.atan2((y4 - cy) / (ry || 1), (x4 - cx) / (rx || 1));
            if (clockwise) {
                while (a1 <= a0) { a1 += 2 * Math.PI; }
            } else {
                while (a1 >= a0) { a1 -= 2 * Math.PI; }
            }
            appendEllipticalArc(d, cx, cy, rx, ry, a0, a1, connectWithMove);
            curX = cx + rx * Math.cos(a1);
            curY = cy + ry * Math.sin(a1);
        }

        while (pos < tokens.length) {
            command = tokens[pos];
            if (!commandRe.test(command)) {
                break; // malformed
            }
            pos += 1;
            switch (command) {
            case "M":
                (function () { var p = nextPoint(); curX = p[0]; curY = p[1]; }());
                d.push("M" + curX + " " + curY);
                while (hasValue()) {
                    (function () { var p = nextPoint(); curX = p[0]; curY = p[1]; }());
                    d.push("L" + curX + " " + curY);
                }
                break;
            case "L":
                while (hasValue()) {
                    (function () { var p = nextPoint(); curX = p[0]; curY = p[1]; }());
                    d.push("L" + curX + " " + curY);
                }
                break;
            case "C":
                while (hasValue()) {
                    d.push("C" + val() + " " + val() + " " + val() + " " + val() + " ");
                    curX = val(); curY = val();
                    d[d.length - 1] += curX + " " + curY;
                }
                break;
            case "Q":
                while (hasValue()) {
                    d.push("Q" + val() + " " + val() + " ");
                    curX = val(); curY = val();
                    d[d.length - 1] += curX + " " + curY;
                }
                break;
            case "X":
                while (hasValue()) {
                    (function () { var p = nextPoint(); quadrant(p[0], p[1], true); }());
                }
                break;
            case "Y":
                while (hasValue()) {
                    (function () { var p = nextPoint(); quadrant(p[0], p[1], false); }());
                }
                break;
            case "A": // arcto, counter-clockwise, connect with line
            case "B": // arc, counter-clockwise, connect with move
                while (hasValue()) {
                    arcTo(val(), val(), val(), val(), val(), val(), val(), val(), false, command === "B");
                }
                break;
            case "W": // clockwise arcto, connect with line
            case "V": // clockwise arc, connect with move
                while (hasValue()) {
                    arcTo(val(), val(), val(), val(), val(), val(), val(), val(), true, command === "V");
                }
                break;
            case "T": // angle-ellipse-to, connect with line
            case "U": // angle-ellipse, connect with move
                while (hasValue()) {
                    (function () {
                        var ecx = val(), ecy = val(), erx = val(), ery = val(),
                            a0 = val() / 65536 * Math.PI / 180,
                            a1 = val() / 65536 * Math.PI / 180;
                        appendEllipticalArc(d, ecx, ecy, erx, ery, a0, a1, command === "U");
                        curX = ecx + erx * Math.cos(a1);
                        curY = ecy + ery * Math.sin(a1);
                    }());
                }
                break;
            case "Z":
                d.push("Z");
                break;
            case "N":
                flush();
                break;
            case "F":
                fill = false;
                break;
            case "S":
                stroke = false;
                break;
            default:
                break;
            }
            count += 1;
            if (count > 100000) {
                break; // safety
            }
        }
        flush();
        return subPaths;
    }
    /**
     * Render a draw:custom-shape's enhanced geometry as an inline SVG image set
     * as the element background, so the actual shape outline (banner, callout,
     * rounded rectangle, ...) is drawn instead of WebODF's default filled
     * bounding box. The bounding-box fill/border that Style2CSS applies is
     * suppressed at the same time.
     * @param {!Element} shape  <draw:custom-shape>
     * @param {!Element} geometry  <draw:enhanced-geometry>
     * @param {!string} shapeId
     * @param {!CSSStyleSheet} stylesheet
     * @return {undefined}
     */
    /**
     * Translate a computed CSS gradient (as emitted by Style2CSS for a
     * draw:fill="gradient" shape) into an SVG <defs> gradient plus a fill
     * reference, so a custom shape can be filled with it. Returns null when the
     * value is not a gradient we can parse. Keeping Style2CSS as the single
     * source of gradient semantics, this only re-expresses it for SVG.
     * @param {string} css  e.g. "linear-gradient(63deg, rgb(..) 0%, rgb(..) 100%)"
     * @param {string} id
     * @return {?{def: !string, ref: !string}}
     */
    function cssGradientToSvg(css, id) {
        var isRadial = css.indexOf('radial-gradient') !== -1,
            gid = id + "_grad",
            stopRe = /(rgba?\([^)]*\)|#[0-9a-fA-F]{3,8})\s*(-?[0-9.]+)%/g,
            stops = [],
            angleMatch = css.match(/(-?[0-9.]+)deg/),
            angle = angleMatch ? parseFloat(angleMatch[1]) : 180,
            rad,
            dx,
            dy,
            def,
            m,
            i;
        m = stopRe.exec(css);
        while (m) {
            stops.push('<stop offset="' + m[2] + '%" stop-color="' + m[1] + '"/>');
            m = stopRe.exec(css);
        }
        if (stops.length < 2) {
            return null;
        }
        if (isRadial) {
            def = '<radialGradient id="' + gid + '" cx="50%" cy="50%" r="50%">';
        } else {
            // CSS gradient angle: 0deg = "to top", growing clockwise. Convert to
            // an axis vector in objectBoundingBox space (y grows downward).
            rad = angle * Math.PI / 180;
            dx = Math.sin(rad) / 2;
            dy = -Math.cos(rad) / 2;
            def = '<linearGradient id="' + gid + '"'
                + ' x1="' + (0.5 - dx).toFixed(4) + '" y1="' + (0.5 - dy).toFixed(4) + '"'
                + ' x2="' + (0.5 + dx).toFixed(4) + '" y2="' + (0.5 + dy).toFixed(4) + '">';
        }
        for (i = 0; i < stops.length; i += 1) {
            def += stops[i];
        }
        def += isRadial ? '</radialGradient>' : '</linearGradient>';
        return { def: '<defs>' + def + '</defs>', ref: 'url(#' + gid + ')' };
    }
    /**
     * Built-in geometry for the common OOXML preset shapes whose ODF
     * enhanced-path LibreOffice writes degenerately (a single point for an
     * ellipse, bare movetos for a cloud, etc.) or whose formula evaluation is
     * unreliable. Returns SVG sub-paths in a normalised 0..100 viewBox (stretched
     * to the shape box), or null for types handled fine by the enhanced-path.
     * Adjustment handles are ignored (OOXML defaults are used) — enough for a
     * faithful preview. Polygons (triangle/diamond/pentagon/...) are intentionally
     * absent: their enhanced-path renders correctly already.
     * @param {?string} type  value of draw:type
     * @return {?{vb: !Array.<!number>, subPaths: !Array}}
     */
    function buildPresetShape(type, rect) {
        var sub, aspect, headX, bodyTop, bodyBottom;
        switch (type) {
        case "ooxml-ellipse":
            sub = [{ d: "M 0 50 A 50 50 0 1 1 100 50 A 50 50 0 1 1 0 50 Z",
                fill: true, stroke: true }];
            break;
        case "ooxml-roundRect":
            sub = [{ d: "M 16.7 0 L 83.3 0 A 16.7 16.7 0 0 1 100 16.7 L 100 83.3"
                + " A 16.7 16.7 0 0 1 83.3 100 L 16.7 100 A 16.7 16.7 0 0 1 0 83.3"
                + " L 0 16.7 A 16.7 16.7 0 0 1 16.7 0 Z", fill: true, stroke: true }];
            break;
        case "ooxml-donut":
            // outer ring clockwise, inner circle counter-clockwise; even-odd
            // fill leaves the centre hole transparent.
            sub = [{ d: "M 0 50 A 50 50 0 1 1 100 50 A 50 50 0 1 1 0 50 Z"
                + " M 25 50 A 25 25 0 1 0 75 50 A 25 25 0 1 0 25 50 Z",
                fill: true, stroke: true, fillRule: "evenodd" }];
            break;
        case "ooxml-arc":
            // OOXML default arc: a quarter ellipse (top to right), open (no fill).
            sub = [{ d: "M 50 0 A 50 50 0 0 1 100 50", fill: false, stroke: true }];
            break;
        case "ooxml-leftRightArrow":
            aspect = rect && rect.height > 0 ? rect.width / rect.height : 1;
            headX = aspect > 0 ? Math.max(12, Math.min(35, 50 / aspect)) : 25;
            bodyTop = 25;
            bodyBottom = 75;
            sub = [{ d: "M 0 50 L " + headX + " 0 L " + headX + " " + bodyTop
                + " L " + (100 - headX) + " " + bodyTop + " L " + (100 - headX)
                + " 0 L 100 50 L " + (100 - headX) + " 100 L " + (100 - headX)
                + " " + bodyBottom + " L " + headX + " " + bodyBottom
                + " L " + headX + " 100 Z", fill: true, stroke: true }];
            break;
        case "ooxml-cloud":
            // Approximate cloud: a ring of circular bumps. Adjustment-free.
            sub = [{ d: "M 22 66 A 16 16 0 0 1 10 40 A 15 15 0 0 1 26 20"
                + " A 18 18 0 0 1 56 12 A 16 16 0 0 1 82 22 A 15 15 0 0 1 92 46"
                + " A 15 15 0 0 1 78 68 A 18 18 0 0 1 48 74 A 16 16 0 0 1 22 66 Z",
                fill: true, stroke: true }];
            break;
        default:
            return null;
        }
        return { vb: [0, 0, 100, 100], subPaths: sub };
    }
    function renderCustomShape(shape, geometry, shapeId, stylesheet) {
        var window = runtime.getWindow(),
            computed = window && window.getComputedStyle(shape, null),
            rect = shape.getBoundingClientRect && shape.getBoundingClientRect(),
            preset,
            ctx,
            path,
            tokens,
            mappedPts = null,
            subPaths,
            fillColor,
            strokeColor,
            strokeWidth,
            paths = "",
            svg,
            i,
            sp,
            vb,
            vbW,
            vbH,
            realAspect,
            viewBoxAspect,
            target,
            mirrorH,
            mirrorV,
            mirrorTransform,
            fillImage,
            gradient,
            gradientDefs = "",
            fillRef,
            dropShadow,
            shadowColorMatch,
            shadowOffsets;
        if (!computed) {
            return;
        }
        fillColor = computed.backgroundColor;
        if (!fillColor || fillColor === "rgba(0, 0, 0, 0)" || fillColor === "transparent") {
            fillColor = "none";
        }
        // A draw:fill="gradient" shape carries its gradient as a CSS
        // background-image (from Style2CSS). Bake it into the SVG so the shape
        // is actually filled with the gradient rather than left unfilled.
        fillImage = computed.backgroundImage;
        fillRef = fillColor;
        if (fillImage && fillImage.indexOf("repeating-linear-gradient") !== -1) {
            // Hatch fill. A faithful per-shape SVG hatch pattern would need
            // viewBox-unit conversion under non-uniform stretch and hatch fills
            // are rare, so approximate the custom shape with the hatch line
            // colour. Plain (non-custom) shapes keep the real CSS hatch pattern.
            gradient = fillImage.match(/rgba?\([^)]*\)|#[0-9a-fA-F]{3,8}/);
            if (gradient) {
                fillRef = gradient[0];
            }
        } else if (fillImage && fillImage.indexOf("gradient") !== -1) {
            gradient = cssGradientToSvg(fillImage, shapeId);
            if (gradient) {
                gradientDefs = gradient.def;
                fillRef = gradient.ref;
            }
        }
        strokeColor = computed.borderTopStyle === "none" ? "none" : computed.borderTopColor;
        // Approximate the stroke width in viewBox units (the SVG is stretched to
        // the shape box via preserveAspectRatio="none").
        strokeWidth = 0;
        if (strokeColor !== "none" && rect && rect.width > 0) {
            strokeWidth = (parseFloat(computed.borderTopWidth) || 1) / rect.width;
        }

        // Prefer built-in geometry for the OOXML presets LibreOffice exports with
        // a degenerate enhanced-path (ellipse, roundRect, donut, cloud, arc,
        // leftRightArrow). Otherwise drive the shape from its enhanced-path.
        preset = buildPresetShape(geometry.getAttributeNS(drawns, "type"), rect);
        if (preset) {
            vb = preset.vb;
            vbW = vb[2];
            vbH = vb[3];
            subPaths = preset.subPaths;
        } else {
            ctx = createShapeContext(geometry);
            vb = ctx.viewBox;
            vbW = vb[2];
            vbH = vb[3];
            path = geometry.getAttributeNS(drawns, "enhanced-path") || "";
            // When the shape's box is not as square as its viewBox, keep the
            // corners around the stretch point circular by widening (or
            // heightening) the viewBox and remapping the path points instead of
            // scaling uniformly.
            if (rect && rect.width > 0 && rect.height > 0) {
                realAspect = rect.width / rect.height;
                viewBoxAspect = vbW / vbH;
                tokens = path.match(/[A-Za-z]|\?[a-zA-Z0-9]+|\$[0-9]+|[a-z]+|-?[0-9]*\.?[0-9]+/g) || [];
                if (!isNaN(ctx.stretchX) && realAspect > viewBoxAspect * 1.02) {
                    target = vbH * realAspect;
                    mappedPts = stretchedPoints(tokens, ctx, ctx.stretchX, vbW, target, true);
                    if (mappedPts) { vbW = target; }
                } else if (!isNaN(ctx.stretchY) && realAspect < viewBoxAspect / 1.02) {
                    target = vbW / realAspect;
                    mappedPts = stretchedPoints(tokens, ctx, ctx.stretchY, vbH, target, false);
                    if (mappedPts) { vbH = target; }
                }
            }
            subPaths = parseEnhancedPath(path, ctx, mappedPts);
            if (!subPaths.length) {
                return;
            }
            // Some ooxml presets (e.g. flowChartDecision) declare
            // svg:viewBox="0 0 0 0" and use small literal path coordinates (0..2)
            // instead of viewBox-scale formulas. createShapeContext then falls
            // back to the 21600 default and the path would be drawn
            // microscopically. When the path clearly does not span that box,
            // derive the viewBox from the path's own bounds.
            (function () {
                var j, k, coords, x, y,
                    minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
                for (j = 0; j < subPaths.length; j += 1) {
                    coords = subPaths[j].d.match(/-?[0-9]*\.?[0-9]+/g) || [];
                    for (k = 0; k + 1 < coords.length; k += 2) {
                        x = parseFloat(coords[k]);
                        y = parseFloat(coords[k + 1]);
                        if (x < minX) { minX = x; }
                        if (x > maxX) { maxX = x; }
                        if (y < minY) { minY = y; }
                        if (y > maxY) { maxY = y; }
                    }
                }
                if (maxX > minX && maxY > minY
                        && (maxX - minX) < vbW / 2 && (maxY - minY) < vbH / 2) {
                    vb[0] = minX;
                    vb[1] = minY;
                    vbW = maxX - minX;
                    vbH = maxY - minY;
                }
            }());
            vb = [vb[0], vb[1], vbW, vbH];
        }
        for (i = 0; i < subPaths.length; i += 1) {
            sp = subPaths[i];
            paths += '<path d="' + sp.d + '" fill="' + (sp.fill ? fillRef : "none") + '"';
            if (sp.fillRule) {
                paths += ' fill-rule="' + sp.fillRule + '"';
            }
            if (sp.stroke && strokeColor !== "none") {
                paths += ' stroke="' + strokeColor + '" stroke-width="'
                    + (strokeWidth * vb[2]).toFixed(2) + '"';
            }
            paths += "/>";
        }

        // draw:mirror-horizontal / -vertical flip the geometry within its box.
        // Reflect the path coordinates around the viewBox centre; with
        // preserveAspectRatio="none" this matches a flip of the shape box.
        mirrorH = geometry.getAttributeNS(drawns, "mirror-horizontal") === "true";
        mirrorV = geometry.getAttributeNS(drawns, "mirror-vertical") === "true";
        if (mirrorH || mirrorV) {
            mirrorTransform = "translate("
                + (mirrorH ? (2 * vb[0] + vbW) : 0) + " "
                + (mirrorV ? (2 * vb[1] + vbH) : 0) + ") scale("
                + (mirrorH ? -1 : 1) + " " + (mirrorV ? -1 : 1) + ")";
            paths = '<g transform="' + mirrorTransform + '">' + paths + "</g>";
        }

        svg = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="'
            + vb.join(" ") + '" preserveAspectRatio="none">' + gradientDefs + paths + "</svg>";

        shape.setAttributeNS(webodfhelperns, "shapeid", shapeId);
        // A draw:shadow maps to box-shadow in Style2CSS, but box-shadow follows
        // the element's rectangular border-box, so for a non-rectangular shape
        // its straight edges poke out from behind the outline. Re-express it as a
        // filter: drop-shadow, which casts from the SVG background's alpha (the
        // real shape outline).
        dropShadow = "";
        if (computed.boxShadow && computed.boxShadow !== "none") {
            shadowColorMatch = computed.boxShadow.match(/rgba?\([^)]*\)|#[0-9a-fA-F]{3,8}/);
            shadowOffsets = computed.boxShadow.replace(/rgba?\([^)]*\)/, "").match(/-?[0-9.]+px/g) || [];
            if (shadowColorMatch && shadowOffsets.length >= 3) {
                dropShadow = "filter: drop-shadow(" + shadowOffsets[0] + " "
                    + shadowOffsets[1] + " " + shadowOffsets[2] + " "
                    + shadowColorMatch[0] + ");";
            }
        }
        stylesheet.insertRule('draw|custom-shape[webodfhelper|shapeid="' + shapeId + '"] {'
            + "background-image: url(\"data:image/svg+xml;charset=utf-8," + encodeURIComponent(svg) + "\");"
            + "background-repeat: no-repeat;"
            + "background-size: 100% 100%;"
            + "background-color: transparent;"
            + "border: none;"
            + "box-shadow: none;"
            + dropShadow
            + "}", stylesheet.cssRules.length);
    }
    /**
     * Render all draw:custom-shape elements that carry an enhanced geometry.
     * @param {!Element} odfbody
     * @param {!CSSStyleSheet} stylesheet
     * @return {undefined}
     */
    function loadCustomShapes(odfbody, stylesheet) {
        var shapes = domUtils.getElementsByTagNameNS(odfbody, drawns, "custom-shape"),
            i,
            geometry;
        for (i = 0; i < shapes.length; i += 1) {
            geometry = domUtils.getDirectChild(shapes[i], drawns, "enhanced-geometry");
            if (geometry && geometry.getAttributeNS(drawns, "enhanced-path")) {
                try {
                    renderCustomShape(shapes[i], geometry, "shape" + i, stylesheet);
                } catch (/**@type{*}*/e) {
                    runtime.log("could not render custom shape: " + String(e));
                }
            }
        }
    }

    // ---- Embedded chart rendering ------------------------------------------
    // WebODF does not render charts. ODF embeds each chart as a sub-document
    // (draw:object -> "<dir>/content.xml" with a <chart:chart>) whose data lives
    // in a local table. We parse that and draw a static SVG (bar/line/pie/ring),
    // which is enough for a preview.

    /**@const@type{!Array.<!string>}*/
    var chartPalette = ["#4f81bd", "#c0504d", "#9bbb59", "#8064a2", "#4bacc6",
        "#f79646", "#5da5da", "#faa43a", "#60bd68", "#f17cb0"];

    /**
     * @param {!string} s
     * @return {!string}
     */
    function svgEsc(s) {
        return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;")
            .replace(/>/g, "&gt;").replace(/"/g, "&quot;");
    }

    /**
     * Map every chart-family style name to its fill/stroke colour.
     * @param {!Element} chartRoot
     * @return {!Object.<string,!{fill:?string,stroke:?string}>}
     */
    function collectChartStyleColors(chartRoot) {
        var styles = domUtils.getElementsByTagNameNS(chartRoot, stylens, "style"),
            map = {},
            i,
            s,
            name,
            gp;
        for (i = 0; i < styles.length; i += 1) {
            s = styles[i];
            if (s.getAttributeNS(stylens, "family") !== "chart") {
                continue;
            }
            name = s.getAttributeNS(stylens, "name");
            gp = domUtils.getElementsByTagNameNS(s, stylens, "graphic-properties")[0];
            if (name && gp) {
                map[name] = {
                    fill: gp.getAttributeNS(drawns, "fill-color") || null,
                    stroke: gp.getAttributeNS(svgns, "stroke-color") || null
                };
            }
        }
        return map;
    }

    /**
     * Parse a chart sub-document into a simple model.
     * @param {!Document} doc
     * @return {?{type:!string,title:!string,hasLegend:!boolean,categories:!Array.<!string>,
     *            series:!Array.<!{label:!string,color:!string,
     *            values:!Array.<!number>,pointColors:!Array.<!string>}>}}
     */
    function parseChart(doc) {
        var root = doc && doc.documentElement,
            chart = root && domUtils.getElementsByTagNameNS(root, chartns, "chart")[0],
            plotArea,
            colors,
            titleEl,
            cls,
            table,
            headerRows,
            bodyRows,
            headerCells,
            categories = [],
            series = [],
            seriesEls,
            i,
            j,
            row,
            cells,
            sEl,
            pts,
            pointColors,
            styleName;
        if (!chart) {
            return null;
        }
        colors = collectChartStyleColors(root);
        plotArea = domUtils.getElementsByTagNameNS(chart, chartns, "plot-area")[0];
        titleEl = domUtils.getElementsByTagNameNS(chart, chartns, "title")[0];
        cls = (chart.getAttributeNS(chartns, "class") || "").replace("chart:", "");
        table = domUtils.getElementsByTagNameNS(chart, tablens, "table")[0];
        if (!plotArea || !table) {
            return null;
        }
        /**
         * @param {!Element} rowEl
         * @return {!Array.<!Element>}
         */
        function cellsOf(rowEl) {
            return domUtils.getElementsByTagNameNS(rowEl, tablens, "table-cell");
        }
        headerRows = domUtils.getElementsByTagNameNS(table, tablens, "table-header-rows")[0];
        headerCells = headerRows
            ? cellsOf(domUtils.getElementsByTagNameNS(headerRows, tablens, "table-row")[0])
            : [];
        bodyRows = domUtils.getElementsByTagNameNS(
            domUtils.getElementsByTagNameNS(table, tablens, "table-rows")[0] || table,
            tablens, "table-row");
        // categories = first column of the body rows
        for (i = 0; i < bodyRows.length; i += 1) {
            cells = cellsOf(bodyRows[i]);
            categories.push(cells[0] ? (cells[0].textContent || "").trim() : "");
        }
        seriesEls = domUtils.getElementsByTagNameNS(plotArea, chartns, "series");
        for (j = 0; j < seriesEls.length; j += 1) {
            sEl = seriesEls[j];
            styleName = sEl.getAttributeNS(chartns, "style-name");
            // per-point colours (pie/ring give each slice its own style)
            pointColors = [];
            pts = domUtils.getElementsByTagNameNS(sEl, chartns, "data-point");
            for (i = 0; i < pts.length; i += 1) {
                if (pts[i].getAttributeNS(chartns, "style-name")
                        && colors[pts[i].getAttributeNS(chartns, "style-name")]) {
                    pointColors.push(colors[pts[i].getAttributeNS(chartns, "style-name")].fill
                        || chartPalette[pointColors.length % chartPalette.length]);
                }
            }
            series.push({
                label: headerCells[j + 1]
                    ? (headerCells[j + 1].textContent || "").trim()
                    : ("Series " + (j + 1)),
                color: (colors[styleName] && (colors[styleName].fill || colors[styleName].stroke))
                    || chartPalette[j % chartPalette.length],
                values: (function () {
                    var vals = [], k, c;
                    for (k = 0; k < bodyRows.length; k += 1) {
                        c = cellsOf(bodyRows[k])[j + 1];
                        vals.push(c ? parseFloat(c.getAttributeNS(officens, "value")) || 0 : 0);
                    }
                    return vals;
                }()),
                pointColors: pointColors
            });
        }
        return { type: cls, title: titleEl ? (titleEl.textContent || "").trim() : "",
            hasLegend: domUtils.getElementsByTagNameNS(chart, chartns, "legend").length > 0,
            categories: categories, series: series };
    }

    /**
     * A "nice" axis maximum and step for a value range 0..maxVal.
     * @param {!number} maxVal
     * @return {!{max:!number,step:!number}}
     */
    function niceAxis(maxVal, preferExactMax) {
        var exact,
            target = (maxVal <= 0 ? 1 : maxVal) * (preferExactMax ? 1 : 1.05),
            rough = target / 5,
            pow = Math.pow(10, Math.floor(Math.log(rough) / Math.LN10)),
            n = rough / pow,
            step = (n <= 1 ? 1 : n <= 2 ? 2 : n <= 5 ? 5 : 10) * pow;
        if (preferExactMax) {
            exact = exactAxis(maxVal);
            if (exact) {
                return exact;
            }
        }
        return { max: Math.ceil(target / step) * step, step: step };
    }
    /**
     * Prefer the data maximum when it already falls on a readable axis step.
     * @param {!number} maxVal
     * @return {?{max:!number,step:!number}}
     */
    function exactAxis(maxVal) {
        var steps = [1, 2, 2.5, 4, 5, 10],
            pow = Math.pow(10, Math.floor(Math.log(maxVal) / Math.LN10)),
            scales = [pow / 10, pow, pow * 10],
            i,
            j,
            step,
            intervals;
        if (maxVal <= 0) {
            return null;
        }
        for (i = 0; i < scales.length; i += 1) {
            for (j = 0; j < steps.length; j += 1) {
                step = steps[j] * scales[i];
                intervals = maxVal / step;
                if (intervals >= 3 && intervals <= 6
                        && Math.abs(intervals - Math.round(intervals)) < 1e-9) {
                    return { max: maxVal, step: step };
                }
            }
        }
        return null;
    }

    /**
     * Render the parsed chart model to an SVG string at the given cm size.
     * @param {!Object} c  parsed chart model
     * @param {!number} wcm
     * @param {!number} hcm
     * @return {!string}
     */
    function buildChartSvg(c, wcm, hcm) {
        var W = Math.max(120, Math.round(wcm * 40)),
            H = Math.max(90, Math.round(hcm * 40)),
            body = "",
            legendW = 0;
        /**
         * @param {!number} x @param {!number} y @param {!string} t
         * @param {!string} anchor @param {!number} size @param {string=} weight
         * @return {!string}
         */
        function text(x, y, t, anchor, size, weight) {
            return '<text x="' + x + '" y="' + y + '" text-anchor="' + anchor
                + '" font-family="Arial,sans-serif" font-size="' + size + '"'
                + (weight ? ' font-weight="' + weight + '"' : '')
                + ' fill="#404040">' + svgEsc(t) + '</text>';
        }
        /**
         * Legend at the right; returns its width and appends to body.
         * @param {!Array.<!{label:!string,color:!string}>} items
         * @return {undefined}
         */
        function drawLegend(items, top) {
            var maxLen = 0, k, ly;
            for (k = 0; k < items.length; k += 1) {
                maxLen = Math.max(maxLen, items[k].label.length);
            }
            legendW = 22 + maxLen * 6.2;
            for (k = 0; k < items.length; k += 1) {
                ly = top + k * 18;
                body += '<rect x="' + (W - legendW) + '" y="' + ly
                    + '" width="11" height="11" fill="' + items[k].color + '"/>';
                body += text(W - legendW + 16, ly + 10, items[k].label, "start", 11);
            }
        }

        if (c.title) {
            body += text(W / 2, 18, c.title, "middle", 14, "bold");
        }

        if (c.type === "circle" || c.type === "ring") {
            // pie / doughnut: one series, slice per category
            (function () {
                var vals = c.series[0] ? c.series[0].values : [],
                    pcols = c.series[0] ? c.series[0].pointColors : [],
                    total = 0, k, ang0 = -Math.PI / 2, ang1, frac,
                    cx, cy, r, rIn, mid, lx, ly, x0, y0, x1, y1, large, col;
                for (k = 0; k < vals.length; k += 1) { total += vals[k]; }
                if (c.hasLegend) {
                    drawLegend((function () {
                        var it = [], m;
                        for (m = 0; m < c.categories.length; m += 1) {
                            it.push({ label: c.categories[m],
                                color: pcols[m] || chartPalette[m % chartPalette.length] });
                        }
                        return it;
                    }()), 30);
                }
                r = Math.min((W - legendW - 20) / 2, (H - 36) / 2) - 4;
                cx = (W - legendW) / 2;
                cy = 28 + (H - 28) / 2;
                rIn = c.type === "ring" ? r * 0.55 : 0;
                if (total <= 0) { return; }
                for (k = 0; k < vals.length; k += 1) {
                    frac = vals[k] / total;
                    ang1 = ang0 + frac * 2 * Math.PI;
                    x0 = cx + r * Math.cos(ang0); y0 = cy + r * Math.sin(ang0);
                    x1 = cx + r * Math.cos(ang1); y1 = cy + r * Math.sin(ang1);
                    large = (ang1 - ang0) > Math.PI ? 1 : 0;
                    col = pcols[k] || chartPalette[k % chartPalette.length];
                    if (rIn > 0) {
                        body += '<path d="M' + (cx + rIn * Math.cos(ang0)) + ' '
                            + (cy + rIn * Math.sin(ang0)) + ' L' + x0 + ' ' + y0
                            + ' A' + r + ' ' + r + ' 0 ' + large + ' 1 ' + x1 + ' ' + y1
                            + ' L' + (cx + rIn * Math.cos(ang1)) + ' ' + (cy + rIn * Math.sin(ang1))
                            + ' A' + rIn + ' ' + rIn + ' 0 ' + large + ' 0 '
                            + (cx + rIn * Math.cos(ang0)) + ' ' + (cy + rIn * Math.sin(ang0))
                            + ' Z" fill="' + col + '" stroke="#ffffff" stroke-width="1.5"/>';
                    } else {
                        body += '<path d="M' + cx + ' ' + cy + ' L' + x0 + ' ' + y0
                            + ' A' + r + ' ' + r + ' 0 ' + large + ' 1 ' + x1 + ' ' + y1
                            + ' Z" fill="' + col + '" stroke="#ffffff" stroke-width="1.5"/>';
                    }
                    mid = (ang0 + ang1) / 2;
                    lx = cx + (rIn > 0 ? (r + rIn) / 2 : r * 0.6) * Math.cos(mid);
                    ly = cy + (rIn > 0 ? (r + rIn) / 2 : r * 0.6) * Math.sin(mid);
                    body += text(lx, ly + 4, String(vals[k]), "middle", 11);
                    ang0 = ang1;
                }
            }());
        } else {
            // bar / line: shared cartesian axes
            (function () {
                var ml = 30, mr, mt = c.title ? 26 : 10, mb = 20,
                    maxVal = 0, k, s, axis, plotW, plotH, x0, y0,
                    nCat = c.categories.length, gi, catW, t, gx, gy,
                    nSer = c.series.length, bw, bx, vh, px, py, pts, legendItems = [];
                for (s = 0; s < c.series.length; s += 1) {
                    for (k = 0; k < c.series[s].values.length; k += 1) {
                        maxVal = Math.max(maxVal, c.series[s].values[k]);
                    }
                    legendItems.push({ label: c.series[s].label, color: c.series[s].color });
                }
                if (c.hasLegend) {
                    drawLegend(legendItems, mt + 4);
                }
                mr = legendW + 6;
                axis = niceAxis(maxVal, c.type === "line");
                plotW = W - ml - mr;
                plotH = H - mt - mb;
                x0 = ml;
                y0 = mt + plotH;
                // y gridlines + labels
                for (t = 0; t <= axis.max + 1e-9; t += axis.step) {
                    gy = y0 - (t / axis.max) * plotH;
                    body += '<line x1="' + x0 + '" y1="' + gy + '" x2="' + (x0 + plotW)
                        + '" y2="' + gy + '" stroke="#cccccc" stroke-width="1"/>';
                    body += text(x0 - 4, gy + 4,
                        String(Math.round(t * 100) / 100), "end", 10);
                }
                // axis lines
                body += '<line x1="' + x0 + '" y1="' + mt + '" x2="' + x0 + '" y2="'
                    + y0 + '" stroke="#888888" stroke-width="1"/>';
                catW = plotW / (nCat || 1);
                // category labels
                for (gi = 0; gi < nCat; gi += 1) {
                    gx = x0 + catW * (gi + 0.5);
                    body += text(gx, y0 + 14, c.categories[gi], "middle", 10);
                }
                if (c.type === "line") {
                    for (s = 0; s < nSer; s += 1) {
                        pts = "";
                        for (k = 0; k < c.series[s].values.length; k += 1) {
                            px = x0 + catW * (k + 0.5);
                            py = y0 - (c.series[s].values[k] / axis.max) * plotH;
                            pts += (k ? " " : "") + px + "," + py;
                        }
                        body += '<polyline points="' + pts + '" fill="none" stroke="'
                            + c.series[s].color + '" stroke-width="2"/>';
                        for (k = 0; k < c.series[s].values.length; k += 1) {
                            px = x0 + catW * (k + 0.5);
                            py = y0 - (c.series[s].values[k] / axis.max) * plotH;
                            body += '<circle cx="' + px + '" cy="' + py + '" r="3" fill="'
                                + c.series[s].color + '"/>';
                        }
                    }
                } else {
                    // grouped bars
                    bw = (catW * 0.7) / (nSer || 1);
                    for (gi = 0; gi < nCat; gi += 1) {
                        for (s = 0; s < nSer; s += 1) {
                            vh = (c.series[s].values[gi] / axis.max) * plotH;
                            bx = x0 + catW * gi + catW * 0.15 + s * bw;
                            body += '<rect x="' + bx + '" y="' + (y0 - vh) + '" width="'
                                + (bw * 0.92) + '" height="' + vh + '" fill="'
                                + c.series[s].color + '"/>';
                            body += text(bx + bw * 0.46, y0 - vh - 3,
                                String(Math.round(c.series[s].values[gi])), "middle", 10);
                        }
                    }
                }
            }());
        }
        return '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ' + W + ' ' + H
            + '" preserveAspectRatio="xMidYMid meet">'
            + '<rect width="' + W + '" height="' + H + '" fill="#ffffff"/>'
            + body + '</svg>';
    }
    /**
     * Parse an ODF length ("7.112cm", "-3pt") into value + unit.
     * @param {?string} value
     * @return {!{v: !number, u: !string}}
     */
    function parseLength(value) {
        var m = /^(-?[0-9.]+)([a-z%]*)$/.exec((value || "").trim());
        return m ? { v: parseFloat(m[1]), u: m[2] || "" } : { v: 0, u: "" };
    }
    /**
     * @param {!Element} rootElement
     * @param {?string} name
     * @param {!string} id
     * @param {!string} orient
     * @param {!number} width
     * @param {!string} color
     * @param {!boolean} centered
     * @return {!string}
     */
    function buildLineMarker(rootElement, name, id, orient, width, color, centered) {
        var marker = findDrawMarker(rootElement, name),
            vbAttr,
            vb,
            d,
            vbW,
            vbH,
            height,
            refX,
            refY;
        if (!marker || !width) {
            return "";
        }
        vbAttr = (marker.getAttributeNS(svgns, "viewBox") || "0 0 20 20").trim().split(/\s+/);
        vb = [parseFloat(vbAttr[0]) || 0, parseFloat(vbAttr[1]) || 0,
            parseFloat(vbAttr[2]) || 20, parseFloat(vbAttr[3]) || 20];
        d = marker.getAttributeNS(svgns, "d");
        if (!d) {
            return "";
        }
        vbW = vb[2];
        vbH = vb[3];
        height = width * (vbH / vbW);
        refX = centered ? vbW / 2 : vbW;
        refY = vbH / 2;
        return '<marker id="' + id + '" viewBox="' + vb.join(" ")
            + '" markerUnits="userSpaceOnUse" markerWidth="' + width
            + '" markerHeight="' + height + '" refX="' + refX + '" refY="' + refY
            + '" orient="' + orient + '"><path d="' + svgEsc(d)
            + '" transform="rotate(90 ' + (vb[0] + vbW / 2) + " " + (vb[1] + vbH / 2)
            + ')" fill="' + color + '"/></marker>';
    }
    /**
     * Render a draw:line connector. WebODF positions frames/shapes from
     * svg:x/y/width/height, but a line carries svg:x1/y1/x2/y2 and no geometry,
     * so it is otherwise invisible. Draw it as an absolutely positioned element
     * spanning the segment's bounding box with an inline SVG line; the stroke
     * uses non-scaling-stroke so it keeps its width under the box stretch.
     * @param {!Element} line
     * @param {!string} lineId
     * @param {!Element} rootElement
     * @param {!CSSStyleSheet} stylesheet
     * @return {undefined}
     */
    function renderLine(line, lineId, rootElement, stylesheet) {
        var window = runtime.getWindow(),
            computed = window && window.getComputedStyle(line, null),
            x1 = parseLength(line.getAttributeNS(svgns, "x1")),
            y1 = parseLength(line.getAttributeNS(svgns, "y1")),
            x2 = parseLength(line.getAttributeNS(svgns, "x2")),
            y2 = parseLength(line.getAttributeNS(svgns, "y2")),
            styleName = line.getAttributeNS(drawns, "style-name"),
            unit = x1.u || "cm",
            strokeColor,
            strokeWidth,
            minThick,
            markerStartName,
            markerEndName,
            markerStartWidth,
            markerEndWidth,
            markerStart,
            markerEnd,
            markerDefs = "",
            markerAttrs = "",
            left = Math.min(x1.v, x2.v),
            top = Math.min(y1.v, y2.v),
            w = Math.abs(x2.v - x1.v),
            h = Math.abs(y2.v - y1.v),
            svg,
            rule;
        if (!computed) {
            return;
        }
        strokeColor = computed.borderTopStyle === "none" ? "none" : computed.borderTopColor;
        if (!strokeColor || strokeColor === "none") {
            // a line with no explicit stroke still needs to be visible
            strokeColor = computed.color || "#000000";
        }
        strokeWidth = parseFloat(computed.borderTopWidth) || 1;
        markerStartName = resolveGraphicProperty(rootElement, styleName, drawns, "marker-start");
        markerEndName = resolveGraphicProperty(rootElement, styleName, drawns, "marker-end");
        markerStartWidth = parseLength(resolveGraphicProperty(rootElement, styleName,
            drawns, "marker-start-width"));
        markerEndWidth = parseLength(resolveGraphicProperty(rootElement, styleName,
            drawns, "marker-end-width"));
        markerStart = buildLineMarker(rootElement, markerStartName, lineId + "start",
            "auto-start-reverse", markerStartWidth.v, strokeColor,
            resolveGraphicProperty(rootElement, styleName, drawns, "marker-start-center") === "true");
        markerEnd = buildLineMarker(rootElement, markerEndName, lineId + "end",
            "auto", markerEndWidth.v, strokeColor,
            resolveGraphicProperty(rootElement, styleName, drawns, "marker-end-center") === "true");
        if (markerStart) {
            markerDefs += markerStart;
            markerAttrs += ' marker-start="url(#' + lineId + 'start)"';
        }
        if (markerEnd) {
            markerDefs += markerEnd;
            markerAttrs += ' marker-end="url(#' + lineId + 'end)"';
        }
        // Give a horizontal/vertical line (zero extent on one axis) enough box to
        // paint the stroke, and recentre it on the true segment.
        minThick = (strokeWidth / 37.8) * 4; // px -> cm, a few stroke widths
        if (h < minThick) {
            top -= (minThick - h) / 2;
            h = minThick;
        }
        if (w < minThick) {
            left -= (minThick - w) / 2;
            w = minThick;
        }
        svg = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ' + w + ' ' + h
            + '" preserveAspectRatio="none">' + (markerDefs ? "<defs>" + markerDefs + "</defs>" : "")
            + '<line'
            + ' x1="' + (x1.v - left) + '" y1="' + (y1.v - top) + '"'
            + ' x2="' + (x2.v - left) + '" y2="' + (y2.v - top) + '"'
            + ' stroke="' + strokeColor + '" stroke-width="' + strokeWidth
            + '" vector-effect="non-scaling-stroke"' + markerAttrs + "/></svg>";
        line.setAttributeNS(webodfhelperns, "lineid", lineId);
        rule = 'draw|line[webodfhelper|lineid="' + lineId + '"] {'
            + 'position: absolute;'
            + 'left: ' + left + unit + ';'
            + 'top: ' + top + unit + ';'
            + 'width: ' + w + unit + ';'
            + 'height: ' + h + unit + ';'
            + 'background-image: url("data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svg) + '");'
            + 'background-repeat: no-repeat;'
            + 'background-size: 100% 100%;'
            // the stroke is drawn by the SVG; suppress the CSS border that
            // draw:stroke="solid" maps onto the element (it would show as a
            // doubled line along the thin box edges).
            + 'border: none;'
            + '}';
        stylesheet.insertRule(rule, stylesheet.cssRules.length);
    }
    /**
     * Render all draw:line connectors.
     * @param {!Element} odfbody
     * @param {!Element} rootElement
     * @param {!CSSStyleSheet} stylesheet
     * @return {undefined}
     */
    function loadLines(odfbody, rootElement, stylesheet) {
        var lines = domUtils.getElementsByTagNameNS(odfbody, drawns, "line"),
            i;
        for (i = 0; i < lines.length; i += 1) {
            try {
                renderLine(lines[i], "line" + i, rootElement, stylesheet);
            } catch (/**@type{*}*/e) {
                runtime.log("could not render line: " + String(e));
            }
        }
    }
    /**
     * @param {!Element} odfbody
     * @return {undefined}
     */
    function formatParagraphAnchors(odfbody) {
        var n,
            i,
            nodes = xpath.getODFElementsWithXPath(odfbody,
                ".//*[*[@text:anchor-type='paragraph']]",
                odf.Namespaces.lookupNamespaceURI);
        for (i = 0; i < nodes.length; i += 1) {
            n = nodes[i];
            if (n.setAttributeNS) {
                n.setAttributeNS(webodfhelperns, "containsparagraphanchor", true);
            }
        }
    }
    /**
     * Modify tables to support merged cells (col/row span)
     * @param {!Element} odffragment
     * @param {!string} documentns
     * @return {undefined}
     */
    function modifyTables(odffragment, documentns) {
        var i,
            tableCells,
            node;

        /**
         * @param {!Element} node
         * @return {undefined}
         */
        function modifyTableCell(node) {
            // If we have a cell which spans columns or rows,
            // then add col-span or row-span attributes.
            if (node.hasAttributeNS(tablens, "number-columns-spanned")) {
                node.setAttributeNS(documentns, "colspan",
                    node.getAttributeNS(tablens, "number-columns-spanned"));
            }
            if (node.hasAttributeNS(tablens, "number-rows-spanned")) {
                node.setAttributeNS(documentns, "rowspan",
                    node.getAttributeNS(tablens, "number-rows-spanned"));
            }
        }
        /**
         * A bare <table:table-cell> (no table:style-name) inherits its style
         * from the row's table:default-cell-style-name, else the column's.
         * WebODF only styles cells via their own table:style-name, so without
         * this the cell borders/fills (which presentation tables put on the
         * default cell style) are lost. Copy the resolved default onto the cell
         * so the existing style rules apply.
         * @param {!Element} table
         * @return {undefined}
         */
        function applyDefaultCellStyles(table) {
            var columnDefaults = [],
                rowDefault,
                styleName;
            /**
             * @param {!Element} parent  scans table-column descendants in order
             * @return {undefined}
             */
            function collectColumns(parent) {
                var c = parent.firstElementChild,
                    rep,
                    def,
                    j;
                while (c) {
                    if (c.namespaceURI === tablens && c.localName === "table-column") {
                        def = c.getAttributeNS(tablens, "default-cell-style-name");
                        rep = parseInt(c.getAttributeNS(tablens, "number-columns-repeated"), 10) || 1;
                        for (j = 0; j < rep; j += 1) {
                            columnDefaults.push(def || "");
                        }
                    } else if (c.namespaceURI === tablens
                            && (c.localName === "table-columns"
                                || c.localName === "table-header-columns")) {
                        collectColumns(c);
                    }
                    c = c.nextElementSibling;
                }
            }
            /**
             * @param {!Element} parent  styles cells across the row groups
             * @return {undefined}
             */
            function walkRows(parent) {
                var r = parent.firstElementChild,
                    c,
                    rep,
                    idx;
                while (r) {
                    if (r.namespaceURI === tablens && r.localName === "table-row") {
                        rowDefault = r.getAttributeNS(tablens, "default-cell-style-name");
                        idx = 0;
                        c = r.firstElementChild;
                        while (c) {
                            if (c.namespaceURI === tablens
                                    && (c.localName === "table-cell"
                                        || c.localName === "covered-table-cell")) {
                                if (c.localName === "table-cell"
                                        && !c.getAttributeNS(tablens, "style-name")) {
                                    styleName = rowDefault || columnDefaults[idx];
                                    if (styleName) {
                                        c.setAttributeNS(tablens, "table:style-name", styleName);
                                    }
                                }
                                rep = parseInt(c.getAttributeNS(tablens, "number-columns-repeated"), 10) || 1;
                                idx += rep;
                            }
                            c = c.nextElementSibling;
                        }
                    } else if (r.namespaceURI === tablens
                            && (r.localName === "table-rows"
                                || r.localName === "table-header-rows")) {
                        walkRows(r);
                    }
                    r = r.nextElementSibling;
                }
            }
            collectColumns(table);
            walkRows(table);
        }

        domUtils.getElementsByTagNameNS(odffragment, tablens, 'table').forEach(function (table) {
            applyDefaultCellStyles(/**@type{!Element}*/(table));
        });
        tableCells = domUtils.getElementsByTagNameNS(odffragment, tablens, 'table-cell');
        for (i = 0; i < tableCells.length; i += 1) {
            node = /**@type{!Element}*/(tableCells[i]);
            modifyTableCell(node);
        }
    }

    /**
     * Make the text:line-break elements behave like html br element.
     * @param {!Element} odffragment
     * @return {undefined}
     */
    function modifyLineBreakElements(odffragment) {
        var document = odffragment.ownerDocument,
            lineBreakElements = domUtils.getElementsByTagNameNS(odffragment, textns, "line-break");
        lineBreakElements.forEach(function (lineBreak) {
            // Make sure we don't add br more than once as this method is executed whenever user undo an operation.
            if (!lineBreak.hasChildNodes()) {
                lineBreak.appendChild(document.createElement("br"));
            }
        });
    }

    /**
     * PowerPoint-exported ODP can contain empty list items that are not shown by
     * PowerPoint, but WebODF's list CSS still paints their label.
     * @param {!Element} odffragment
     * @param {!CSSStyleSheet} stylesheet
     * @return {undefined}
     */
    function hideEmptyListItems(odffragment, stylesheet) {
        var listItems = domUtils.getElementsByTagNameNS(odffragment, textns, "list-item"),
            i,
            item,
            child,
            hasNestedList,
            hasText,
            firstNonListChild;

        for (i = 0; i < listItems.length; i += 1) {
            item = /**@type{!Element}*/(listItems[i]);
            hasNestedList = false;
            hasText = false;
            firstNonListChild = null;
            child = item.firstElementChild;
            while (child) {
                if (child.namespaceURI === textns && child.localName === "list") {
                    hasNestedList = true;
                } else {
                    if (!firstNonListChild) {
                        firstNonListChild = child;
                    }
                    if (child.textContent.replace(/\s/g, "")) {
                        hasText = true;
                    }
                }
                child = child.nextElementSibling;
            }
            if (!hasText && !hasNestedList) {
                item.setAttributeNS(webodfhelperns, "emptylistitem", "true");
            } else if (!hasText && firstNonListChild) {
                firstNonListChild.setAttributeNS(webodfhelperns, "emptylistlabel", "true");
            }
        }
        stylesheet.insertRule('text|list-item[webodfhelper|emptylistitem="true"] {display: none !important;}',
            stylesheet.cssRules.length);
        stylesheet.insertRule('text|list-item > [webodfhelper|emptylistlabel="true"]:first-child:before '
            + '{content: none !important;}', stylesheet.cssRules.length);
    }

    /**
     * Expand ODF spaces of the form <text:s text:c=N/> to N consecutive
     * <text:s/> elements. This makes things simpler for WebODF during
     * handling of spaces, in particular during editing.
     * @param {!Element} odffragment
     * @return {undefined}
     */
    function expandSpaceElements(odffragment) {
        var spaces,
            doc = odffragment.ownerDocument;

        /**
         * @param {!Element} space
         * @return {undefined}
         */
        function expandSpaceElement(space) {
            var j, count;
            // If the space has any children, remove them and put a " " text
            // node in place.
            domUtils.removeAllChildNodes(space);
            space.appendChild(doc.createTextNode(" "));

            count = parseInt(space.getAttributeNS(textns, "c"), 10);
            if (count > 1) {
                // Make it a 'simple' space node
                space.removeAttributeNS(textns, "c");
                // Prepend count-1 clones of this space node to itself
                for (j = 1; j < count; j += 1) {
                    space.parentNode.insertBefore(space.cloneNode(true), space);
                }
            }
        }

        spaces = domUtils.getElementsByTagNameNS(odffragment, textns, "s");
        spaces.forEach(expandSpaceElement);
    }

    /**
     * Expand tabs to contain tab characters. This eases cursor behaviour
     * during editing
     * @param {!Element} odffragment
     */
    function expandTabElements(odffragment) {
        var tabs;

        tabs = domUtils.getElementsByTagNameNS(odffragment, textns, "tab");
        tabs.forEach(function(tab) {
            tab.textContent = "\t";
        });
    }
    /**
     * @param {!Element} odfbody
     * @param {!CSSStyleSheet} stylesheet
     * @return {undefined}
     **/
    function modifyDrawElements(odfbody, stylesheet) {
        var node,
            /**@type{!Array.<!Element>}*/
            drawElements = [],
            i;
        // find all the draw:* elements
        node = odfbody.firstElementChild;
        while (node && node !== odfbody) {
            if (node.namespaceURI === drawns) {
                drawElements[drawElements.length] = node;
            }
            if (node.firstElementChild) {
                node = node.firstElementChild;
            } else {
                while (node && node !== odfbody && !node.nextElementSibling) {
                    node = /**@type{!Element}*/(node.parentNode);
                }
                if (node && node.nextElementSibling) {
                    node = node.nextElementSibling;
                }
            }
        }
        // adjust all the frame positions
        for (i = 0; i < drawElements.length; i += 1) {
            node = drawElements[i];
            setDrawElementPosition('frame' + String(i), node, stylesheet);
        }
        formatParagraphAnchors(odfbody);
    }

    /**
     * @param {!odf.Formatting} formatting
     * @param {!odf.OdfContainer} odfContainer
     * @param {!Element} shadowContent
     * @param {!Element} odfbody
     * @param {!CSSStyleSheet} stylesheet
     * @return {undefined}
     **/
    function cloneMasterPages(formatting, odfContainer, shadowContent, odfbody, stylesheet) {
        var masterPageName,
            masterPageElement,
            styleId,
            clonedPageElement,
            clonedElement,
            clonedDrawElements,
            pageNumber = 0,
            i,
            element,
            elementToClone,
            document = odfContainer.rootElement.ownerDocument;

        element = odfbody.firstElementChild;
        // no master pages to expect?
        if (!(element && element.namespaceURI === officens &&
              (element.localName === "presentation" || element.localName === "drawing"))) {
            return;
        }

        element = element.firstElementChild;
        while (element) {
            // If there was a master-page-name attribute, then we are dealing with a draw:page.
            // Get the referenced master page element from the master styles
            masterPageName = element.getAttributeNS(drawns, 'master-page-name');
            masterPageElement = masterPageName ? formatting.getMasterPageElement(masterPageName) : null;

            // If the referenced master page exists, create a new page and copy over it's contents into the new page,
            // except for the ones that are placeholders. Also, call setDrawElementPosition on each of those child frames.
            if (masterPageElement) {
                styleId = element.getAttributeNS(webodfhelperns, 'styleid');
                clonedPageElement = document.createElementNS(drawns, 'draw:page');

                elementToClone = masterPageElement.firstElementChild;
                i = 0;
                while (elementToClone) {
                    if (shouldCloneMasterPageElement(odfContainer.rootElement, /**@type{!Element}*/(element), masterPageElement,
                            /**@type{!Element}*/(elementToClone))) {
                        clonedElement = /**@type{!Element}*/(elementToClone.cloneNode(true));
                        clonedPageElement.appendChild(clonedElement);
                    }
                    elementToClone = elementToClone.nextElementSibling;
                    i += 1;
                }
                // TODO: above already do not clone nodes which match the rule for being dropped
                dropTemplateDrawFrames(clonedPageElement);

                // Position all elements
                clonedDrawElements = domUtils.getElementsByTagNameNS(clonedPageElement, drawns, '*');
                for (i = 0; i < clonedDrawElements.length; i += 1) {
                    setDrawElementPosition(styleId + '_' + i, clonedDrawElements[i], stylesheet);
                }

                // Append the cloned master page to the "Shadow Content" element outside the main ODF dom
                shadowContent.appendChild(clonedPageElement);

                // Get the page number by counting the number of previous master pages in this shadowContent
                pageNumber = String(shadowContent.getElementsByTagNameNS(drawns, 'page').length);
                // Get the page-number tag in the cloned master page and set the text content to the calculated number
                setContainerValue(clonedPageElement, textns, 'page-number', pageNumber);

                // Care for header
                setContainerValue(clonedPageElement, presentationns, 'header', getHeaderFooter(odfContainer, /**@type{!Element}*/(element), 'header'));
                // Care for footer
                setContainerValue(clonedPageElement, presentationns, 'footer', getHeaderFooter(odfContainer, /**@type{!Element}*/(element), 'footer'));

                // Now call setDrawElementPosition on this new page to set the proper dimensions
                setDrawElementPosition(styleId, clonedPageElement, stylesheet);
                // Add a custom attribute with the style name of the normal page, so the CSS rules created for the styles of the normal page
                // to display/hide frames of certain classes from the master page can address the cloned master page belonging to that normal page
                // Cmp. addDrawPageFrameDisplayRules in Style2CSS
                clonedPageElement.setAttributeNS(webodfhelperns, 'page-style-name', element.getAttributeNS(drawns, 'style-name'));
                // TODO: investigate if the attributes draw:style-name and style:page-layoutname should be copied over
                // to the cloned page from the master page as well, or if this one below is enough already
                // And finally, add an attribute referring to the master page, so the CSS targeted for that master page will style this
                clonedPageElement.setAttributeNS(drawns, 'draw:master-page-name', masterPageElement.getAttributeNS(stylens, 'name'));
            }

            element = element.nextElementSibling;
        }
    }

    /**
     * @param {!odf.OdfContainer} container
     * @param {!Element} plugin
     * @return {undefined}
     **/
    function setVideo(container, plugin) {
        var video, source, url, doc = plugin.ownerDocument,
            /**@type{!odf.OdfPart}*/
            part;

        url = plugin.getAttributeNS(xlinkns, 'href');

        /**
         * @param {?string} url
         * @param {string} mimetype
         * @return {undefined}
         */
        function callback(url, mimetype) {
            var ns = doc.documentElement.namespaceURI;
            // test for video mimetypes
            if (mimetype.substr(0, 6) === 'video/') {
                video = doc.createElementNS(ns, "video");
                video.setAttribute('controls', 'controls');

                source = doc.createElementNS(ns, 'source');
                if (url) {
                    source.setAttribute('src', url);
                }
                source.setAttribute('type', mimetype);

                video.appendChild(source);
                plugin.parentNode.appendChild(video);
            } else {
                plugin.innerHtml = 'Unrecognised Plugin';
            }
        }
        /**
         * @param {!odf.OdfPart} p
         */
        function onchange(p) {
            callback(p.url, p.mimetype);
        }
        // look for a office:binary-data
        if (url) {
            try {
                part = container.getPart(url);
                part.onchange = onchange;
                part.load();
            } catch (/**@type{*}*/e) {
                runtime.log('slight problem: ' + String(e));
            }
        } else {
        // this will fail  atm - following function assumes PNG data]
            runtime.log('using MP4 data fallback');
            url = getUrlFromBinaryDataElement(plugin);
            callback(url, 'video/mp4');
        }
    }

    /**
     * @param {!HTMLHeadElement} head
     * @return {?HTMLStyleElement}
     */
    function findWebODFStyleSheet(head) {
        var style = head.firstElementChild;
        while (style && !(style.localName === "style"
                && style.hasAttribute("webodfcss"))) {
            style = style.nextElementSibling;
        }
        return /**@type{?HTMLStyleElement}*/(style);
    }

    /**
     * @param {!Document} document
     * @return {!HTMLStyleElement}
     */
    function addWebODFStyleSheet(document) {
        var head = /**@type{!HTMLHeadElement}*/(document.getElementsByTagName('head')[0]),
            css,
            /**@type{?HTMLStyleElement}*/
            style,
            href,
            count = document.styleSheets.length;
        // make sure this is only added once per HTML document, e.g. in case of
        // multiple odfCanvases
        style = findWebODFStyleSheet(head);
        if (style) {
            count = parseInt(style.getAttribute("webodfcss"), 10);
            style.setAttribute("webodfcss", count + 1);
            return style;
        }
        if (String(typeof webodf_css) === "string") {
            css = /**@type{!string}*/(webodf_css);
        } else {
            href = "webodf.css";
            if (runtime.currentDirectory) {
                href = runtime.currentDirectory();
                if (href.length > 0 && href.substr(-1) !== "/") {
                    href += "/";
                }
                href += "../webodf.css";
            }
            css = /**@type{!string}*/(runtime.readFileSync(href, "utf-8"));
        }
        style = /**@type{!HTMLStyleElement}*/(document.createElementNS(head.namespaceURI, 'style'));
        style.setAttribute('media', 'screen, print, handheld, projection');
        style.setAttribute('type', 'text/css');
        style.setAttribute('webodfcss', '1');
        style.appendChild(document.createTextNode(css));
        head.appendChild(style);
        return style;
    }

    /**
     * @param {!HTMLStyleElement} webodfcss
     * @return {undefined}
     */
    function removeWebODFStyleSheet(webodfcss) {
        var count = parseInt(webodfcss.getAttribute("webodfcss"), 10);
        if (count === 1) {
             webodfcss.parentNode.removeChild(webodfcss);
        } else {
             webodfcss.setAttribute("count", count - 1);
        }
    }

    /**
     * @param {!Document} document Put and ODF Canvas inside this element.
     * @return {!HTMLStyleElement}
     */
    function addStyleSheet(document) {
        var head = /**@type{!HTMLHeadElement}*/(document.getElementsByTagName('head')[0]),
            style = document.createElementNS(head.namespaceURI, 'style'),
            /**@type{string}*/
            text = '';
        style.setAttribute('type', 'text/css');
        style.setAttribute('media', 'screen, print, handheld, projection');
        odf.Namespaces.forEachPrefix(function(prefix, ns) {
            text += "@namespace " + prefix + " url(" + ns + ");\n";
        });
        text += "@namespace webodfhelper url(" + webodfhelperns + ");\n";
        style.appendChild(document.createTextNode(text));
        head.appendChild(style);
        return /**@type {!HTMLStyleElement}*/(style);
    }
    /**
     * This class manages a loaded ODF document that is shown in an element.
     * It takes care of giving visual feedback on loading, ensures that the
     * stylesheets are loaded.
     * @constructor
     * @implements {gui.AnnotatableCanvas}
     * @implements {ops.Canvas}
     * @implements {core.Destroyable}
     * @param {!HTMLElement} element Put and ODF Canvas inside this element.
     * @param {!gui.Viewport=} viewport Viewport used for scrolling elements and ranges into view
     */
    odf.OdfCanvas = function OdfCanvas(element, viewport) {
        runtime.assert((element !== null) && (element !== undefined),
            "odf.OdfCanvas constructor needs DOM element");
        runtime.assert((element.ownerDocument !== null) && (element.ownerDocument !== undefined),
            "odf.OdfCanvas constructor needs DOM");
        var self = this,
            doc = /**@type{!Document}*/(element.ownerDocument),
            /**@type{!odf.OdfContainer}*/
            odfcontainer,
            /**@type{!odf.Formatting}*/
            formatting = new odf.Formatting(),
            /**@type{!PageSwitcher}*/
            pageSwitcher,
            /**@type{HTMLDivElement}*/
            sizer = null,
            /**@type{HTMLDivElement}*/
            annotationsPane = null,
            allowAnnotations = false,
            showAnnotationRemoveButton = false,
            /**@type{gui.AnnotationViewManager}*/
            annotationViewManager = null,
            /**@type{!HTMLStyleElement}*/
            webodfcss,
            /**@type{!HTMLStyleElement}*/
            fontcss,
            /**@type{!HTMLStyleElement}*/
            stylesxmlcss,
            /**@type{!HTMLStyleElement}*/
            positioncss,
            shadowContent,
            /**@type{!number}*/
            autofitCounter = 0,
            /**@type{!Object.<string,!Array.<!Function>>}*/
            eventHandlers = {},
            waitingForDoneTimeoutId,
            /**@type{!core.ScheduledTask}*/redrawContainerTask,
            shouldRefreshCss = false,
            shouldRerenderAnnotations = false,
            loadingQueue = new LoadingQueue(),
            /**@type{!gui.ZoomHelper}*/
            zoomHelper = new gui.ZoomHelper(),
            /**@type{!gui.Viewport}*/
            canvasViewport = viewport || new gui.SingleScrollViewport(/**@type{!HTMLElement}*/(element.parentNode));

        /**
         * Load all the images that are inside an odf element.
         * @param {!odf.OdfContainer} container
         * @param {!Element} odffragment
         * @param {!CSSStyleSheet} stylesheet
         * @return {undefined}
         */
        function loadImages(container, odffragment, stylesheet) {
            var i,
                images,
                node;
            /**
             * Do delayed loading for all the images
             * @param {string} name
             * @param {!odf.OdfContainer} container
             * @param {!Element} node
             * @param {!CSSStyleSheet} stylesheet
             * @return {undefined}
             */
            function loadImage(name, container, node, stylesheet) {
                // load image with a small delay to give the html ui a chance to
                // update
                loadingQueue.addToQueue(function () {
                    setImage(name, container, node, stylesheet);
                });
            }
            images = odffragment.getElementsByTagNameNS(drawns, 'image');
            for (i = 0; i < images.length; i += 1) {
                node = /**@type{!Element}*/(images.item(i));
                loadImage('image' + String(i), container, node, stylesheet);
                applyImageClip(node, 'imageclip' + String(i), container, stylesheet);
            }
        }
        /**
         * Render embedded chart objects (draw:object -> chart sub-document) as
         * static SVG backgrounds on the object element.
         * @param {!odf.OdfContainer} container
         * @param {!Element} odffragment
         * @param {!CSSStyleSheet} stylesheet
         * @return {undefined}
         */
        function loadCharts(container, odffragment, stylesheet) {
            var objects = odffragment.getElementsByTagNameNS(drawns, 'object'),
                i;
            /**
             * @param {!Element} object
             * @param {!string} chartId
             * @return {undefined}
             */
            function loadChart(object, chartId) {
                var href = object.getAttributeNS(xlinkns, 'href'),
                    path,
                    frame = object.parentNode;
                if (!href) {
                    return;
                }
                path = href.replace(/^\.?\//, "").replace(/\/$/, "") + "/content.xml";
                container.getPartData(path, function (err, data) {
                    var doc, model, svg, w, h, rule;
                    if (err || !data) {
                        return;
                    }
                    try {
                        doc = runtime.parseXML(runtime.byteArrayToString(data, "utf8"));
                        model = parseChart(doc);
                        if (!model || !model.series.length) {
                            return;
                        }
                        w = parseLength(frame.getAttributeNS(svgns, "width"));
                        h = parseLength(frame.getAttributeNS(svgns, "height"));
                        svg = buildChartSvg(model, w.v || 12, h.v || 7);
                        object.setAttributeNS(webodfhelperns, "chartid", chartId);
                        rule = 'draw|object[webodfhelper|chartid="' + chartId + '"] {'
                            + 'display: block; width: 100%; height: 100%;'
                            + 'background-image: url("data:image/svg+xml;charset=utf-8,'
                            + encodeURIComponent(svg) + '");'
                            + 'background-repeat: no-repeat;'
                            + 'background-position: center;'
                            + 'background-size: 100% 100%;'
                            + '}';
                        stylesheet.insertRule(rule, stylesheet.cssRules.length);
                    } catch (/**@type{*}*/e) {
                        runtime.log("could not render chart: " + String(e));
                    }
                });
            }
            for (i = 0; i < objects.length; i += 1) {
                loadChart(/**@type{!Element}*/(objects.item(i)), 'chart' + String(i));
            }
        }
        /**
         * Load all the video that are inside an odf element.
         * @param {!odf.OdfContainer} container
         * @param {!Element} odffragment
         * @return {undefined}
         */
        function loadVideos(container, odffragment) {
            var i,
                plugins,
                node;
            /**
             * Do delayed loading for all the videos
             * @param {!odf.OdfContainer} container
             * @param {!Element} node
             * @return {undefined}
             */
            function loadVideo(container, node) {
                // load video with a small delay to give the html ui a chance to
                // update
                loadingQueue.addToQueue(function () {
                    setVideo(container, node);
                });
            }
            // embedded video is stored in a draw:plugin element
            plugins = odffragment.getElementsByTagNameNS(drawns, 'plugin');
            for (i = 0; i < plugins.length; i += 1) {
                node = /**@type{!Element}*/(plugins.item(i));
                loadVideo(container, node);
            }
        }

        /**
         * Register an event handler
         * @param {!string} eventType
         * @param {!Function} eventHandler
         * @return {undefined}
         */
        function addEventListener(eventType, eventHandler) {
            var handlers;
            if (eventHandlers.hasOwnProperty(eventType)) {
                handlers = eventHandlers[eventType];
            } else {
                handlers = eventHandlers[eventType] = [];
            }
            if (eventHandler && handlers.indexOf(eventHandler) === -1) {
                handlers.push(eventHandler);
            }
        }
        /**
         * Fire an event
         * @param {!string} eventType
         * @param {Array.<Object>=} args
         * @return {undefined}
         */
        function fireEvent(eventType, args) {
            if (!eventHandlers.hasOwnProperty(eventType)) {
                return;
            }
            var handlers = eventHandlers[eventType], i;
            for (i = 0; i < handlers.length; i += 1) {
                handlers[i].apply(null, args);
            }
        }

        /**
         * @return {undefined}
         */
        function fixContainerSize() {
            var minHeight,
                odfdoc = sizer.firstChild;

            if (!odfdoc) {
                return;
            }

            // All zooming of the sizer within the canvas
            // is done relative to the top-left corner.
            sizer.style.WebkitTransformOrigin = "0% 0%";
            sizer.style.MozTransformOrigin = "0% 0%";
            sizer.style.msTransformOrigin = "0% 0%";
            sizer.style.OTransformOrigin = "0% 0%";
            sizer.style.transformOrigin = "0% 0%";

            if (annotationViewManager) {
                minHeight = annotationViewManager.getMinimumHeightForAnnotationPane();
                if (minHeight) {
                    sizer.style.minHeight = minHeight;
                } else {
                    sizer.style.removeProperty('min-height');
                }
            }

            // The sizer is now scaled with the CSS `zoom` property (see
            // gui.ZoomHelper), which rescales its layout box, so its offset
            // dimensions already reflect the zoom. Let the inline-block canvas
            // element shrink-wrap the floated sizer instead of pinning it to a
            // transform-scaled size, so the white page box matches exactly.
            element.style.width = "";
            element.style.height = "";
            // Re-apply inline-block to canvas element on resizing.
            // Chrome tends to forget this property after a relayout
            element.style.display = "inline-block";
        }

        /**
         * @return {undefined}
         */
        function redrawContainer() {
            if (shouldRefreshCss) {
                handleStyles(odfcontainer, formatting, stylesxmlcss);
                shouldRefreshCss = false;
                // different styles means different layout, thus different sizes
            }
            if (shouldRerenderAnnotations) {
                if (annotationViewManager) {
                    annotationViewManager.rerenderAnnotations();
                }
                shouldRerenderAnnotations = false;
            }
            fixContainerSize();
        }

        /**
         * Collect the names of graphic styles (resolving style inheritance)
         * that request presentation text autofit, i.e. carry
         * style:shrink-to-fit="true" or draw:fit-to-size="shrink-to-fit".
         * @param {!odf.OdfContainer} container
         * @return {?Object.<string,boolean>}  null when no style asks for it
         */
        function collectAutofitStyleNames(container) {
            var rootElement = container.rootElement,
                roots = [rootElement.automaticStyles, rootElement.styles],
                directly = {},
                parent = {},
                names = [],
                result = null,
                r,
                styles,
                j,
                s,
                name,
                gp;

            /**
             * @param {string} n
             * @param {!number} depth
             * @return {!boolean}
             */
            function inherits(n, depth) {
                if (!n || depth > 32) {
                    return false;
                }
                if (directly[n] === true) {
                    return true;
                }
                return inherits(parent[n] || '', depth + 1);
            }

            for (r = 0; r < roots.length; r += 1) {
                if (!roots[r]) {
                    continue;
                }
                styles = domUtils.getElementsByTagNameNS(roots[r], stylens, 'style');
                for (j = 0; j < styles.length; j += 1) {
                    s = styles[j];
                    if (s.getAttributeNS(stylens, 'family') !== 'graphic') {
                        continue;
                    }
                    name = s.getAttributeNS(stylens, 'name');
                    if (!name) {
                        continue;
                    }
                    names.push(name);
                    parent[name] = s.getAttributeNS(stylens, 'parent-style-name') || '';
                    gp = domUtils.getElementsByTagNameNS(s, stylens, 'graphic-properties')[0];
                    if (gp && (gp.getAttributeNS(stylens, 'shrink-to-fit') === 'true'
                            || gp.getAttributeNS(drawns, 'fit-to-size') === 'shrink-to-fit')) {
                        directly[name] = true;
                    }
                }
            }
            for (j = 0; j < names.length; j += 1) {
                if (inherits(names[j], 0)) {
                    if (!result) {
                        result = {};
                    }
                    result[names[j]] = true;
                }
            }
            return result;
        }

        /**
         * Presentation text autofit ("shrink text on overflow"). WebODF gives
         * every placeholder a fixed height, so a paragraph that is too tall just
         * overflows its frame. For frames whose graphic style asks for
         * shrink-to-fit, scale the text box down (CSS zoom) so it fits. Frames
         * that already fit, and decks that never request autofit, are untouched
         * (the ratio is measured from rendered geometry, so it is independent of
         * the canvas zoom applied afterwards).
         * @param {!Element} odfbody
         * @param {!odf.OdfContainer} container
         * @param {!CSSStyleSheet} css
         * @return {undefined}
         */
        function shrinkAutofitText(odfbody, container, css) {
            var autofit = collectAutofitStyleNames(container),
                frames,
                shapes,
                i,
                j,
                frame,
                shape,
                textbox,
                paragraphs,
                r,
                minTop,
                maxBottom,
                frameHeight,
                textHeight,
                scale,
                id;
            if (!autofit) {
                return;
            }
            frames = domUtils.getElementsByTagNameNS(odfbody, drawns, 'frame');
            for (i = 0; i < frames.length; i += 1) {
                frame = frames[i];
                if (autofit[frame.getAttributeNS(drawns, 'style-name')] !== true) {
                    continue;
                }
                textbox = domUtils.getElementsByTagNameNS(frame, drawns, 'text-box')[0];
                if (!textbox) {
                    continue;
                }
                frameHeight = frame.getBoundingClientRect().height;
                textHeight = textbox.getBoundingClientRect().height;
                if (frameHeight > 0 && textHeight > frameHeight + 1) {
                    // zoom < 1 shrinks the text; clamp so it never disappears.
                    scale = Math.max(0.3, (frameHeight / textHeight) * 0.97);
                    id = "autofit" + (autofitCounter += 1);
                    textbox.setAttributeNS(webodfhelperns, 'autofitid', id);
                    css.insertRule('draw|text-box[webodfhelper|autofitid="' + id
                        + '"] { zoom: ' + scale + '; }', css.cssRules.length);
                }
            }

            shapes = domUtils.getElementsByTagNameNS(odfbody, drawns, 'custom-shape');
            for (i = 0; i < shapes.length; i += 1) {
                shape = shapes[i];
                if (autofit[shape.getAttributeNS(drawns, 'style-name')] !== true) {
                    continue;
                }
                paragraphs = domUtils.getElementsByTagNameNS(shape, textns, 'p');
                if (!paragraphs.length) {
                    continue;
                }
                frameHeight = shape.getBoundingClientRect().height;
                minTop = Infinity;
                maxBottom = -Infinity;
                for (j = 0; j < paragraphs.length; j += 1) {
                    r = paragraphs[j].getBoundingClientRect();
                    if (r.height > 0) {
                        minTop = Math.min(minTop, r.top);
                        maxBottom = Math.max(maxBottom, r.bottom);
                    }
                }
                textHeight = maxBottom > minTop ? maxBottom - minTop : 0;
                if (frameHeight > 0 && textHeight > frameHeight + 1) {
                    scale = Math.max(0.3, (frameHeight / textHeight) * 0.97);
                    id = "autofit" + (autofitCounter += 1);
                    shape.setAttributeNS(webodfhelperns, 'autofitid', id);
                    css.insertRule('draw|custom-shape[webodfhelper|autofitid="' + id
                        + '"] > text|p { zoom: ' + scale + '; }', css.cssRules.length);
                }
            }
        }

        /**
         * A new content.xml has been loaded. Update the live document with it.
         * @param {!odf.OdfContainer} container
         * @param {!odf.ODFDocumentElement} odfnode
         * @return {undefined}
         **/
        function handleContent(container, odfnode) {
            var css = /**@type{!CSSStyleSheet}*/(positioncss.sheet);
            // only append the content at the end
            domUtils.removeAllChildNodes(element);

            sizer = /**@type{!HTMLDivElement}*/(doc.createElementNS(element.namespaceURI, 'div'));
            sizer.style.display = "inline-block";
            sizer.style.background = "white";
            // The #shadowContent master-page overlay is position:absolute with
            // top/left:0, so it must be anchored to the sizer (which carries the
            // slide flow and the zoom). Without an explicit position the sizer is
            // static, so the overlay instead resolves against a higher ancestor:
            // any padding/header the host page puts around the canvas then shifts
            // every master shape relative to its slide text. Make the sizer the
            // containing block so the overlay always lines up with the slides.
            sizer.style.position = "relative";
            // When the window is shrunk such that the
            // canvas container has a horizontal scrollbar,
            // zooming out seems to not make the scrollable
            // width disappear. This extra scrollable
            // width seems to be proportional to the
            // annotation pane's width. Setting the 'float'
            // of the sizer to 'left' fixes this in webkit.
            sizer.style.setProperty("float", "left", "important");
            sizer.appendChild(odfnode);
            element.appendChild(sizer);

            // An annotations pane div. Will only be shown when annotations are enabled
            annotationsPane = /**@type{!HTMLDivElement}*/(doc.createElementNS(element.namespaceURI, 'div'));
            annotationsPane.id = "annotationsPane";
            // A "Shadow Content" div. This will contain stuff like pages
            // extracted from <style:master-page>. These need to be nicely
            // styled, so we will populate this in the ODF body first. Once the
            // styling is handled, it can then be lifted out of the
            // ODF body and placed beside it, to not pollute the ODF dom.
            shadowContent = doc.createElementNS(element.namespaceURI, 'div');
            shadowContent.id = "shadowContent";
            shadowContent.style.position = 'absolute';
            shadowContent.style.top = 0;
            shadowContent.style.left = 0;
            container.getContentElement().appendChild(shadowContent);

            modifyDrawElements(odfnode.body, css);
            cloneMasterPages(formatting, container, shadowContent, odfnode.body, css);
            modifyTables(odfnode.body, element.namespaceURI);
            modifyLineBreakElements(odfnode.body);
            hideEmptyListItems(odfnode.body, css);
            expandSpaceElements(odfnode.body);
            expandTabElements(odfnode.body);
            loadImages(container, odfnode.body, css);
            loadCharts(container, odfnode.body, css);
            loadVideos(container, odfnode.body);
            // WebODF's default graphic style lists "page" among its tags, so it
            // leaks a fill/border onto every draw:page. Reset it (low specificity)
            // so pages are transparent by default and the white canvas sizer
            // shows through; named page fills and the bitmap rules below have
            // higher specificity and still win. Without this a master-page clone
            // shows that stray fill instead of being see-through, and (for decks
            // whose slides are full-bleed bitmaps) the clone's leftover fill
            // ghosts against the slide on devices that render the absolutely
            // positioned clone slightly off the in-flow slide.
            css.insertRule('draw|page {background-color: transparent; border: 0;}',
                css.cssRules.length);
            loadHiddenDrawLayers(container, css);
            loadPageBackgrounds(container, odfnode.body, css);
            loadMasterPageBackgrounds(container, css);
            loadCustomShapes(odfnode.body, css);
            loadLines(odfnode.body, container.rootElement, css);

            sizer.insertBefore(shadowContent, sizer.firstChild);
            zoomHelper.setZoomableElement(sizer);

            // Once the slide is laid out, shrink any overflowing autofit text.
            shrinkAutofitText(odfnode.body, container, css);
        }

        /**
         * This should create an annotations pane if non existent, and then populate it with annotations
         * If annotations are disallowed, it should remove the pane and all annotations
         * @param {!odf.ODFDocumentElement} odfnode
         */
        function handleAnnotations(odfnode) {
            var annotationNodes;

            if (allowAnnotations) {
                if (!annotationsPane.parentNode) {
                    sizer.appendChild(annotationsPane);
                }
                if (annotationViewManager) {
                    annotationViewManager.forgetAnnotations();
                }
                annotationViewManager = new gui.AnnotationViewManager(self, odfnode.body, annotationsPane, showAnnotationRemoveButton);
                annotationNodes = /**@type{!Array.<!odf.AnnotationElement>}*/(domUtils.getElementsByTagNameNS(odfnode.body, officens, 'annotation'));
                annotationViewManager.addAnnotations(annotationNodes);

                fixContainerSize();
            } else {
                if (annotationsPane.parentNode) {
                    sizer.removeChild(annotationsPane);
                    annotationViewManager.forgetAnnotations();
                    fixContainerSize();
                }
            }
        }

        /**
         * @param {boolean} suppressEvent Suppress the statereadychange event from firing. Used for refreshing the OdtContainer
         * @return {undefined}
         **/
        function refreshOdf(suppressEvent) {

            // synchronize the object a window.odfcontainer with the view
            function callback() {
                // clean up
                clearCSSStyleSheet(fontcss);
                clearCSSStyleSheet(stylesxmlcss);
                clearCSSStyleSheet(positioncss);

                domUtils.removeAllChildNodes(element);

                // setup
                element.style.display = "inline-block";
                var odfnode = odfcontainer.rootElement;
                element.ownerDocument.importNode(odfnode, true);

                formatting.setOdfContainer(odfcontainer);
                handleFonts(odfcontainer, fontcss);
                handleStyles(odfcontainer, formatting, stylesxmlcss);
                // do content last, because otherwise the document is constantly
                // updated whenever the css changes
                handleContent(odfcontainer, odfnode);
                handleAnnotations(odfnode);

                if (!suppressEvent) {
                    loadingQueue.addToQueue(function () {
                        fireEvent("statereadychange", [odfcontainer]);
                    });
                }
            }

            if (odfcontainer.state === odf.OdfContainer.DONE) {
                callback();
            } else {
                // so the ODF is not done yet. take care that we'll
                // do the work once it is done:

                // FIXME: use callback registry instead of replacing the onchange
                runtime.log("WARNING: refreshOdf called but ODF was not DONE.");

                waitingForDoneTimeoutId = runtime.setTimeout(function later_cb() {
                    if (odfcontainer.state === odf.OdfContainer.DONE) {
                        callback();
                    } else {
                        runtime.log("will be back later...");
                        waitingForDoneTimeoutId = runtime.setTimeout(later_cb, 500);
                    }
                }, 100);
            }
        }

        /**
         * Updates the CSS rules to match the ODF document styles and also
         * updates the size of the canvas to match the new layout.
         * Needs to be called after changes to the styles of the ODF document.
         * @return {undefined}
         */
        this.refreshCSS = function () {
            shouldRefreshCss = true;
            redrawContainerTask.trigger();
        };

        /**
         * Updates the size of the canvas to the size of the content.
         * Needs to be called after changes to the content of the ODF document.
         * @return {undefined}
         */
        this.refreshSize = function () {
            redrawContainerTask.trigger();
        };
        /**
         * @return {!odf.OdfContainer}
         */
        this.odfContainer = function () {
            return odfcontainer;
        };
        /**
         * Set a odfcontainer manually.
         * @param {!odf.OdfContainer} container
         * @param {boolean=} suppressEvent Default value is false
         * @return {undefined}
         */
        this.setOdfContainer = function (container, suppressEvent) {
            odfcontainer = container;
            refreshOdf(suppressEvent === true);
        };
        /**
         * @param {string} url
         * @return {undefined}
         */
        function load(url) {
            // clean up
            loadingQueue.clearQueue();

            // FIXME: We need to support parametrized strings, because
            // drop-in word replacements are inadequate for translations;
            // see http://techbase.kde.org/Development/Tutorials/Localization/i18n_Mistakes#Pitfall_.232:_Word_Puzzles
            domUtils.removeAllChildNodes(element);
            element.appendChild(element.ownerDocument.createTextNode(runtime.tr('Loading') + url + '...'));
            element.removeAttribute('style');
            // open the odf container
            odfcontainer = new odf.OdfContainer(url, function (container) {
                // assignment might be necessary if the callback
                // fires before the assignment above happens.
                odfcontainer = container;
                refreshOdf(false);
            });
        }
        this["load"] = load;
        this.load = load;

        /**
         * @param {function(?string):undefined} callback
         * @return {undefined}
         */
        this.save = function (callback) {
            odfcontainer.save(callback);
        };

        /**
         * @param {!string} eventName
         * @param {!function(*)} handler
         * @return {undefined}
         */
        this.addListener = function (eventName, handler) {
            switch (eventName) {
            case "click":
                listenEvent(element, eventName, handler); break;
            default:
                addEventListener(eventName, handler); break;
            }
        };

        /**
         * @return {!odf.Formatting}
         */
        this.getFormatting = function () {
            return formatting;
        };

        /**
         * @return {gui.AnnotationViewManager}
         */
        this.getAnnotationViewManager = function () {
            return annotationViewManager;
        };

        /**
         * Unstyles and untracks all annotations present in the document,
         * and then tracks them again with fresh rendering
         * @return {undefined}
         */
        this.refreshAnnotations = function () {
            handleAnnotations(odfcontainer.rootElement);
        };

        /**
         * Re-renders all annotations if enabled
         * @return {undefined}
         */
        this.rerenderAnnotations = function () {
            if (annotationViewManager) {
                shouldRerenderAnnotations = true;
                redrawContainerTask.trigger();
            }
        };

        /**
         * This returns the element inside the canvas which can be zoomed with
         * CSS and which contains the ODF document and the annotation sidebar.
         * @return {!HTMLElement}
         */
        this.getSizer = function () {
            return /**@type{!HTMLElement}*/(sizer);
        };

        /** Allows / disallows annotations
         * @param {!boolean} allow
         * @param {!boolean} showRemoveButton
         * @return {undefined}
         */
        this.enableAnnotations = function (allow, showRemoveButton) {
            if (allow !== allowAnnotations) {
                allowAnnotations = allow;
                showAnnotationRemoveButton = showRemoveButton;
                if (odfcontainer) {
                    handleAnnotations(odfcontainer.rootElement);
                }
            }
        };

        /**
         * Adds an annotation for the annotaiton manager to track
         * and wraps and highlights it
         * @param {!odf.AnnotationElement} annotation
         * @return {undefined}
         */
        this.addAnnotation = function (annotation) {
            if (annotationViewManager) {
                annotationViewManager.addAnnotations([annotation]);
                fixContainerSize();
            }
        };

        /**
         * Stops an annotation and unwraps it
         * @param {!odf.AnnotationElement} annotation
         * @return {undefined}
         */
        this.forgetAnnotation = function (annotation) {
            if (annotationViewManager) {
                annotationViewManager.forgetAnnotation(annotation);
                fixContainerSize();
            }
        };

        /**
         * @return {!gui.ZoomHelper}
         */
        this.getZoomHelper = function () {
            return zoomHelper;
        };

        /**
         * @param {!number} zoom
         * @return {undefined}
         */
        this.setZoomLevel = function (zoom) {
            zoomHelper.setZoomLevel(zoom);
        };
        /**
         * @return {!number}
         */
        this.getZoomLevel = function () {
            return zoomHelper.getZoomLevel();
        };
        /**
         * @param {!number} width
         * @param {!number} height
         * @return {undefined}
         */
        this.fitToContainingElement = function (width, height) {
            var zoomLevel = zoomHelper.getZoomLevel(),
                realWidth = element.offsetWidth / zoomLevel,
                realHeight = element.offsetHeight / zoomLevel,
                zoom;

            zoom = width / realWidth;
            if (height / realHeight < zoom) {
                zoom = height / realHeight;
            }
            zoomHelper.setZoomLevel(zoom);
        };
        /**
         * @param {!number} width
         * @return {undefined}
         */
        this.fitToWidth = function (width) {
            var realWidth = element.offsetWidth / zoomHelper.getZoomLevel();
            zoomHelper.setZoomLevel(width / realWidth);
        };
        /**
         * @param {!number} width
         * @param {!number} height
         * @return {undefined}
         */
        this.fitSmart = function (width, height) {
            var realWidth, realHeight, newScale,
                zoomLevel = zoomHelper.getZoomLevel();

            realWidth = element.offsetWidth / zoomLevel;
            realHeight = element.offsetHeight / zoomLevel;

            newScale = width / realWidth;
            if (height !== undefined) {
                if (height / realHeight < newScale) {
                    newScale = height / realHeight;
                }
            }

            zoomHelper.setZoomLevel(Math.min(1.0, newScale));
        };
        /**
         * @param {!number} height
         * @return {undefined}
         */
        this.fitToHeight = function (height) {
            var realHeight = element.offsetHeight / zoomHelper.getZoomLevel();
            zoomHelper.setZoomLevel(height / realHeight);
        };
        /**
         * @return {undefined}
         */
        this.showFirstPage = function () {
            pageSwitcher.showFirstPage();
        };
        /**
         * @return {undefined}
         */
        this.showNextPage = function () {
            pageSwitcher.showNextPage();
        };
        /**
         * @return {undefined}
         */
        this.showPreviousPage = function () {
            pageSwitcher.showPreviousPage();
        };
        /**
         * @param {!number} n  number of the page
         * @return {undefined}
         */
        this.showPage = function (n) {
            pageSwitcher.showPage(n);
            fixContainerSize();
        };

        /**
         * @return {!HTMLElement}
         */
        this.getElement = function () {
            return element;
        };

        /**
         * @return {!gui.Viewport}
         */
        this.getViewport = function () {
            return canvasViewport;
        };

        /**
         * Add additional css rules for newly inserted draw:frame and draw:image. eg. position, dimensions and background image
         * @param {!Element} frame
         */
        this.addCssForFrameWithImage = function (frame) {
            // TODO: frameid and imageid generation here is better brought in sync with that for the images on loading of a odf file.
            var frameName = frame.getAttributeNS(drawns, 'name'),
                fc = frame.firstElementChild;
            setDrawElementPosition(frameName, frame,
                    /**@type{!CSSStyleSheet}*/(positioncss.sheet));
            if (fc) {
                setImage(frameName + 'img', odfcontainer, fc,
                   /**@type{!CSSStyleSheet}*/( positioncss.sheet));
            }
        };
        /**
         * @param {!function(!Error=)} callback, passing an error object in case of error
         * @return {undefined}
         */
        this.destroy = function(callback) {
            var head = /**@type{!HTMLHeadElement}*/(doc.getElementsByTagName('head')[0]),
                cleanup = [pageSwitcher.destroy, redrawContainerTask.destroy];

            runtime.clearTimeout(waitingForDoneTimeoutId);
            // TODO: anything to clean with annotationViewManager?
            if (annotationsPane && annotationsPane.parentNode) {
                annotationsPane.parentNode.removeChild(annotationsPane);
            }

            zoomHelper.destroy(function () {
                if (sizer) {
                    element.removeChild(sizer);
                    sizer = null;
                }
            });

            // remove all styles
            removeWebODFStyleSheet(webodfcss);
            head.removeChild(fontcss);
            head.removeChild(stylesxmlcss);
            head.removeChild(positioncss);

            // TODO: loadingQueue, make sure it is empty
            core.Async.destroyAll(cleanup, callback);
        };

        function init() {
            webodfcss = addWebODFStyleSheet(doc);
            pageSwitcher = new PageSwitcher(addStyleSheet(doc));
            fontcss = addStyleSheet(doc);
            stylesxmlcss = addStyleSheet(doc);
            positioncss = addStyleSheet(doc);
            redrawContainerTask = core.Task.createRedrawTask(redrawContainer);
            zoomHelper.subscribe(gui.ZoomHelper.signalZoomChanged, fixContainerSize);
        }

        init();
    };
}());
