/*
 * Builds a single self-contained webodf.js from the source modules, without the
 * CMake/Closure toolchain. The official "compiled" webodf.js is just the source
 * files concatenated in dependency order with IS_COMPILED_CODE flipped to true
 * (so runtime.loadClass becomes a no-op instead of fetching files over XHR).
 * This reproduces that, plus the generated webodf_css string, then minifies.
 *
 * The output carries WebODF's compiled-file license header, which adds the
 * AGPL section-7 permissions: an HTML file that merely calls into this code is
 * a separate work (so the host app is not forced under the AGPL). Because this
 * build is a *modified* version, we extend that exception to it by keeping the
 * notice, and we point recipients at the Corresponding Source (the fork).
 *
 * terser is optional; if it cannot be resolved the bundle is written unminified.
 *
 * Usage: node build-bundle.js <output.js>
 */
"use strict";

var fs = require("fs");
var path = require("path");

var LICENSE_HEADER = [
    "/*!",
    " * WebODF compiled library — https://github.com/kogmbh/WebODF/",
    " * Copyright (C) 2010-2015 KO GmbH <copyright@kogmbh.com>",
    " *",
    " * @licstart",
    " * The code in this file is free software: you can redistribute it and/or",
    " * modify it under the terms of the GNU Affero General Public License (GNU",
    " * AGPL) as published by the Free Software Foundation, either version 3 of",
    " * the License, or (at your option) any later version.",
    " *",
    " * The code in this file is distributed in the hope that it will be useful,",
    " * but WITHOUT ANY WARRANTY; without even the implied warranty of",
    " * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU Affero",
    " * General Public License for more details.",
    " *",
    " * As additional permission under GNU AGPL version 3 section 7, you may",
    " * distribute UNMODIFIED VERSIONS OF THIS file without the copy of the GNU",
    " * AGPL normally required by section 4, provided you include this license",
    " * notice and a URL through which recipients can access the Corresponding",
    " * Source.",
    " *",
    " * As a special exception to the AGPL, any HTML file which merely makes",
    " * function calls to this code, and for that purpose includes it in",
    " * unmodified form by reference or in-line shall be deemed a separate work",
    " * for copyright law purposes. In addition, the copyright holders of this",
    " * code give you permission to combine this code with free software",
    " * libraries that are released under the GNU LGPL. You may copy and",
    " * distribute such a system following the terms of the GNU AGPL for this",
    " * code and the LGPL for the libraries. If you modify this code, you may",
    " * extend this exception to your version of the code, but you are not",
    " * obligated to do so. If you do not wish to do so, delete this exception",
    " * statement from your version.",
    " *",
    " * This license applies to this entire compilation.",
    " * @licend",
    " *",
    " * @source: http://www.webodf.org/",
    " * @source: https://github.com/kogmbh/WebODF/",
    " * Corresponding Source for this modified build: https://github.com/voltara13/WebODF",
    " */"
].join("\n") + "\n";

var libDir = path.join(__dirname, "webodf", "lib");
var manifest = JSON.parse(fs.readFileSync(path.join(libDir, "manifest.json"), "utf8"));

// Only bundle what odf.OdfCanvas transitively needs (drops the editor/ops UI).
var ROOT = "odf.OdfCanvas";

function collect(name, seen, order) {
    if (seen[name]) { return; }
    seen[name] = true;
    var deps = manifest[name] || [];
    deps.forEach(function (dep) { collect(dep, seen, order); });
    order.push(name); // post-order => dependencies precede dependents
}

var order = [];
collect(ROOT, {}, order);

function classToPath(name) {
    return path.join(libDir, name.replace(/\./g, path.sep) + ".js");
}

function cssToJs() {
    var css = fs.readFileSync(path.join(__dirname, "webodf", "webodf.css"), "utf8");
    css = css.replace(/\/\*([\r\n]|.)*?\*\//g, "");
    css = css.replace(/(^\s*)|(\s*$)/gm, "");
    css = css.replace(/\r?\n/g, "");
    css = css.replace(/\\/g, "\\\\");
    css = css.replace(/'/g, "\\'");
    return "var webodf_css = '" + css + "';\n";
}

function resolveTerser() {
    // terser is not a WebODF dependency; borrow the copy in the sibling rtf.js
    // checkout if present, else fall back to a normal resolve.
    var candidates = [
        path.join(__dirname, "..", "rtf.js", "node_modules"),
        path.join(__dirname, "node_modules"),
    ];
    for (var i = 0; i < candidates.length; i += 1) {
        try {
            return require(require.resolve("terser", { paths: [candidates[i]] }));
        } catch (e) { /* try next */ }
    }
    return null;
}

function build() {
    var out = [];

    // 1. Runtime, with the compiled-code flag enabled.
    var runtime = fs.readFileSync(path.join(libDir, "runtime.js"), "utf8");
    runtime = runtime.replace("var IS_COMPILED_CODE = false;", "var IS_COMPILED_CODE = true;");
    if (runtime.indexOf("var IS_COMPILED_CODE = true;") === -1) {
        throw new Error("Failed to patch IS_COMPILED_CODE flag");
    }
    out.push(runtime);

    // 2. Generated stylesheet string.
    out.push(cssToJs());

    // 3. Modules in dependency order.
    order.forEach(function (name) {
        out.push(fs.readFileSync(classToPath(name), "utf8"));
    });

    return out.join("\n");
}

async function main() {
    var outPath = process.argv[2];
    if (!outPath) { throw new Error("Usage: node build-bundle.js <output.js>"); }

    var body = build();
    var code = body;
    var terser = resolveTerser();
    if (terser) {
        // mangle:false — WebODF's runtime inspects function names in places, so
        // renaming identifiers is unsafe; whitespace/comment + compress alone
        // still roughly halves the size.
        var result = await terser.minify(body, {
            compress: true,
            mangle: false,
            format: { comments: false },
        });
        if (result.error) { throw result.error; }
        code = result.code;
        console.log("Minified with terser.");
    } else {
        console.log("terser not found — writing unminified bundle.");
    }

    fs.writeFileSync(outPath, LICENSE_HEADER + code + "\n");

    console.log("Bundled " + order.length + " modules from root " + ROOT);
    console.log("Wrote " + outPath + " (" + Math.round(fs.statSync(outPath).size / 1024) + " KB)");
}

main().catch(function (e) { console.error(e); process.exit(1); });
