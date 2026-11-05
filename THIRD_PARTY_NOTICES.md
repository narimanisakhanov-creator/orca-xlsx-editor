# Third-party notices

The plugin ships two third-party packages: ExcelJS (vendored into `editor/assets/` and patched) and
fflate (bundled into `editor/assets/app-<hash>.js`). The vendored ExcelJS browser build is itself a
bundle of ExcelJS and its dependencies, so it carries those components too — see *Third-party code
inside the ExcelJS bundle* below. The BSD-3-Clause and MIT texts are reproduced at the end.

## ExcelJS 4.4.0

- Upstream: https://github.com/exceljs/exceljs — MIT, Copyright (c) 2014-2019 Guyon Roche.
- What is used: the browser build `node_modules/exceljs/dist/exceljs.bare.min.js` (UMD, exposes
  `window.ExcelJS`), which `build/build.mjs` copies into `editor/assets/exceljs-4.4.0.min.js`. That
  build step is the only thing that writes the file; esbuild keeps ExcelJS out of the app bundle, and
  the frame loads the vendor copy as a classic script.
- What the copy differs by: the header comment and the patches below, and nothing else.
- What is patched: two kinds of dynamic code construction, six spans in total, both refused by the
  editor frame's CSP (`default-src 'self'`) and by `orca plugin check`:
  1. one span in the bundled `setImmediate` polyfill, which built a function object from the
     argument's source text when the argument was not a function. The replacement installs a
     callable that throws `TypeError('setImmediate expects a function')`; ExcelJS only ever passes
     real functions, so no shipped path reaches that branch.
  2. five spans of lodash's global probe `Function("return this")()`, each replaced with
     `globalThis`. The expression sat behind a short-circuit that always resolves `window` or `self`
     first, so the replacement produces the same object.
- Size: of 862 631 upstream bytes, the patched body is 862 604 (−27: +48 for the `setImmediate`
  replacement, −75 for the five probes). The shipped file is 863 518 bytes — a 914-byte header, then
  that body. The patch is deterministic, so rebuilding reproduces the file byte for byte.

## Third-party code inside the ExcelJS bundle

`exceljs.bare.min.js` is a Browserify build, so ExcelJS's own dependencies and the Node shims it
needs travel inside it. They are not separate files and this plugin does not call them directly;
they are listed because they are distributed inside the file we ship. Among them:

- JSZip v3.10.1 (its banner is still in the file), and pako, which JSZip's own banner names;
- the `lodash.*` modules, fast-csv, uuid, dayjs, readable-stream, xmlchars, saxes, safe-buffer and
  ieee754;
- Browserify's shims for Node built-ins — `buffer`, `events`, `process`, `setimmediate`, `stream` —
  and the crypto stack JSZip's browser build pulls in (`asn1.js`, `bn.js`, `browserify-aes`,
  `browserify-cipher`, `browserify-rsa`, `browserify-sign`, `create-hash`, `create-hmac`, `elliptic`,
  `hash.js`, `md5.js`, `public-encrypt`, `randombytes`, `randomfill`, `sha.js` and their
  dependencies).

Licences: MIT for the great majority (ExcelJS's dependencies, the Browserify shims, buffer,
fast-csv, uuid, dayjs, readable-stream, xmlchars, safe-buffer, lodash); pako under MIT with a Zlib
part; ISC for saxes; BSD-3-Clause for ieee754; MIT or GPL-3.0-or-later, at your choice, for JSZip.

The bundle keeps every licence banner it was built with — the buffer module, ieee754, JSZip and
safe-buffer — byte for byte. The BSD-3-Clause text that ieee754's banner refers to is reproduced
below, as is the MIT text.

## fflate 0.8.3

- Upstream: https://github.com/101arrowz/fflate — MIT, Copyright (c) 2026 Arjun Barrett.
- What is used: the browser ESM build (`fflate`, resolved to `esm/browser.js`) for the raw-DEFLATE
  `deflateSync` / `Inflate` pair in `src/write/zip.js`. Its unprefixed codecs are DEFLATE without a
  zlib or gzip wrapper, which is what ZIP entries carry; the `*RawSync` names do not exist in 0.8.3.
  esbuild bundles it, unmodified, into `editor/assets/app-<hash>.js`.

## BSD-3-Clause license text (ieee754)

```
Copyright 2008 Fair Oaks Labs, Inc.

Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:

1. Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.

2. Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.

3. Neither the name of the copyright holder nor the names of its contributors may be used to endorse or promote products derived from this software without specific prior written permission.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
```

## MIT license text

```
The MIT License (MIT)

Copyright (c) 2014-2019 Guyon Roche            (ExcelJS)
Copyright (c) 2026 Arjun Barrett                (fflate)

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```
