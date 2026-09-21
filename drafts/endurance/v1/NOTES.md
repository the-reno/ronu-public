# Endurance — layout v1

Status: draft for layout review. Created 2026-09-21.

## Starting point

Preserves the dark editorial prototype supplied in this conversation. The single HTML prototype is now split into page structure, content, styling, and behavior. Added links back to the draft hub, scoped the local focus preference to this version, and enlarged small controls for touch input.

## Current structure

Header → section navigation → introduction → topic tabs → explanation and conceptual visual → optional detail → next section.

Desktop: persistent left navigation and a wider reading area. Phone: collapsible section menu and stacked content. Focus view: removes the large introduction and sidebar, keeping a chapter selector available. Search: all chapters and topics. Deep links: chapter and topic in the URL fragment.

## Review next

- Is section-by-section navigation preferable to one continuous article?
- Is the balance between explanation and visual comfortable on both screen sizes?
- Should the introduction be more compact?
- Which technical details should stay visible rather than inside an expandable row?

## Checks before each revision

- No page-level horizontal overflow at 360, 390, 768, 1024, and 1440 CSS pixels.
- Menus, topic tabs, conceptual steps, disclosures, search, and Focus view work with touch and keyboard.
- Browser back/forward and copied topic links return to the expected content.
- The navigation drawer closes correctly after selection and after a viewport change.
- Reduced-motion preferences are respected; text can be zoomed.
- Source links and “All drafts” return links remain correct.

The preview size controls simulate widths, not real devices. Real iOS/Android browser testing remains part of review. Content and conceptual diagrams are placeholders; review scientific accuracy and sources separately before publication.
