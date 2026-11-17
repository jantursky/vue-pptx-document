# Security policy

## Supported versions

Only the latest release of `vue-pptx-document` gets fixes.

## Reporting a vulnerability

Please don't open a public issue. Use GitHub's [private vulnerability reporting](https://github.com/jantursky/vue-pptx-document/security/advisories/new), or email jan.tursky@gmail.com with a description and, if you can, a minimal reproduction.

You'll get an answer within a week. Once a fix is released, the advisory is published with credit, unless you'd rather stay anonymous.

Issues in the file writer should also go to [PptxGenJS](https://github.com/gitbrent/PptxGenJS/security).

## Things to know

- Presentations are rendered from the data you pass in. Text from users ends up in the file as text; it isn't executed.
- Image `src` values that aren't data URIs are loaded when the file is written: in Node, http(s) URLs are fetched (10 s timeout, 10 MB cap, `image/*` responses only) and anything else is read as a file path; in the browser, PptxGenJS fetches them. When a deck is built from untrusted input, pass `allowFiles: false` so Node never reads the file system, and a `loadImage` function that only fetches from hosts you trust. See [Images and untrusted input](https://vue-pptx-document-docs.vercel.app/guides/server#images-and-untrusted-input).
- The preview only turns `http`, `https`, `mailto`, `tel` and relative links into clickable `<a>` elements; other schemes, such as `javascript:`, are shown as plain text.
- The importer reads untrusted files: XML is parsed without external entities, and text, attributes and URLs are escaped in the generated component. It holds the whole file and its pictures in memory, so cap the upload size if you run it on a server.
