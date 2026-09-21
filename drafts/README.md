# Ronu — draft workspace

Public layout prototypes, kept separate from the live Ronu website.

- [Open the draft hub](https://raw.githack.com/the-reno/ronu-public/main/drafts/index.html)
- [Open Endurance v1](https://raw.githack.com/the-reno/ronu-public/main/drafts/endurance/v1/index.html)

These browser previews use raw.githack, a third-party development proxy for public GitHub files. It may show a confirmation screen on the first visit. Updates can be cached for several minutes. This is a review link, not the production website. No hosting settings or files in `the-reno/the-reno.github.io` have been changed.

## Structure

```text
drafts/
  index.html                Draft hub and responsive preview controls
  README.md                 This workflow
  templates/
    README.md               How to start another topic or version
  endurance/
    v1/
      index.html            Page structure
      content.js            Chapters, topics, and explanatory copy
      styles.css            Dark visual system and responsive layout
      app.js                Navigation, search, dialogs, and interactions
      NOTES.md              Decisions, open questions, and review checklist
```

## Phone and desktop

Open the hub in Safari, Chrome, or another modern browser. “Open draft” uses the full available screen width. In the hub, Responsive / Phone / Desktop changes the width of the embedded page. A large preview is scaled down when necessary; it is an overview, not a substitute for testing on a real device.

The prototype uses a section drawer on small screens and persistent navigation on wider screens. Both use the same content and files. No installation, package manager, account, or build is required to run the page. JavaScript must be enabled.

## Working together

1. Review a version using its browser link. Reference the chapter, topic, and screen size when suggesting a change.
2. Edit `content.js` for copy, `styles.css` for presentation, `app.js` for behavior, or `index.html` for page structure.
3. Update `NOTES.md` when a design decision changes. Test phone and desktop before committing.
4. Keep small revisions in the same version and use Git history. For a substantially different direction, copy the version to `v2` and update the hub links, embedded preview, and labels together.
5. Publish to the live website only as a separate, explicitly approved change. Nothing here automatically replaces a production page.

## Privacy

This repository and these previews are public. Do not commit credentials, private drafts, health records, client material, or personal datasets. Private working material belongs in `the-reno/ronu-private`, not here. `noindex` is a request to search engines, not access control. The page stores only a local focus-view preference and adds no analytics, external fonts, or tracking scripts. The preview host handles network requests under its own terms.

## Local preview

From the repository root:

```sh
python -m http.server 8000 --bind 127.0.0.1
```

Open `http://127.0.0.1:8000/drafts/index.html`. Keep all files together. Double-clicking the draft HTML also works, but a local server is preferable for testing navigation and assets.

## References

- Responsive design: https://developer.mozilla.org/en-US/docs/Learn_web_development/Core/CSS_layout/Responsive_Design
- Preview service and caching: https://raw.githack.com/

The Endurance text is illustrative layout content, not an approved or validated physiology article.
