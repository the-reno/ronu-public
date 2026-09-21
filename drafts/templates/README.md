# Start a new draft

Use a separate directory for each topic, and a separate version directory for each major layout direction.

```text
drafts/<topic>/v1/
  index.html
  content.js
  styles.css
  app.js
  NOTES.md
```

Copy `drafts/endurance/v1/` as the working template. The `CHAPTERS` array in `content.js` drives sections and topics. It contains public sample copy, which must be replaced or explicitly kept as placeholder content.

Update the document title, description, headings, navigation labels, chapter count, and local-storage key. Change topic IDs to stable, unique slugs. Check both ordinary and Focus view. When changing the number of chapters, also update the static labels and progress bar limits in `index.html`; do not assume every label is generated.

Add an entry to `drafts/index.html`, with working version and notes links. The hub's embedded preview currently shows Endurance v1; update its iframe `src` when choosing a new featured draft.

Use `NOTES.md` for status, changes, and questions. Avoid `final-final` filenames: small changes use Git history; substantial alternatives use `v2`, `v3`, and so on. Keep confidential content in the private repository.
