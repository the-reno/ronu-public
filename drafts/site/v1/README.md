# Ronu — current-content site preview v1

A separate implementation of the approved dark library layout. This directory does not change the production website, its navigation, DNS, or GitHub Pages settings.

## Open the preview

https://raw.githack.com/the-reno/ronu-public/main/drafts/site/v1/index.html

This is a public development preview served through raw.githack, not the production site. The host may show a repository confirmation on the first visit and cache updates for several minutes. No links were added to the main website. `noindex, nofollow, noarchive` requests are included, but are not access control. Do not add private material here.

## Content included

The eight topics linked from the Ronu homepage, checked on 21 September 2026:

- Triathlon: Endurance; Body Mechanics; Arm Mechanics; From Chaos to Motion.
- Science: The Complexity of Prediction; The Spark; Phantom Traffic Jam; The Dice and the Storm.

Markets and Maker keep their section introductions without invented articles or projects. The Dice and the Storm is explicitly labeled Introduction only, matching the existing source. Unlisted experimental pages and private data are excluded.

The source repository is `the-reno/the-reno.github.io`. Article imports are pinned to commit `23df86d3f7c8ab99fc22c3c2072c16b6abc533e5`, the production main-branch revision at the time of this preview. The source content was written before the snapshot date; the snapshot date is not an article publication date.

Articles are fetched from public raw GitHub source, with a raw.githack fallback. Original article bodies, reference lists, tables, and inline diagrams are imported into a scoped reader. Duplicate page titles and old global navigation are replaced by the new shell. Editorial summaries are used on cards and reader headers. This is a presentation migration, not a new scientific or factual review.

Interactive pages are loaded only after selecting Load interactive. They use the existing public ronu.one URLs in an iframe; the models, controls, and original inner-page design are not rewritten. Open separately remains available if a browser blocks framing. Leaving the topic removes the frame. These live interactive pages may change independently of the pinned articles.

## Files and editing

- `index.html`: page shell, navigation, dialogs, accessibility labels.
- `content.js`: source revision, section introductions, topic catalogue, source paths.
- `styles.css`: approved dark visual system.
- `reader.css`: source-reader and embedded-model layout additions.
- `art.js`: decorative schematic illustrations; not data or simulation output.
- `reader.js`: article import, contents links, source fallback, model frames.
- `app.js`: routes, search, format filters, grid/list preference, browsing.

No installation, package manager, analytics, or login is required. JavaScript and internet access are required for imported articles and original models. Only the grid/list display preference is stored locally in the browser.

For local review from this directory:

```sh
python -m http.server 8000 --bind 127.0.0.1
```

Then open http://127.0.0.1:8000/ in a browser. Keep all seven runtime files together.

## Validation and limits

The local Chromium checks covered eight-topic navigation, section/format filters, search and reset, list/grid views, accurate empty states, article import and contents links, the introduction-only label, opt-in frame mounting/unloading, expand/close/Escape, and horizontal overflow at 390px and 320px. JavaScript syntax checks passed.

The test environment blocks browser network navigation. Local reader tests used representative source fragments, and model tests verified the frame controls rather than the remote simulations. They are not evidence of full live-browser or numerical-model validation. Read the live preview in a normal browser before production migration, including the full From Chaos to Motion diagrams and references. Existing source wording has not been independently reviewed here.

## Publishing later

Keep revisions here while the design is reviewed. Migration to the main website is a separate, explicitly approved change. Nothing in this directory automatically replaces production content.
