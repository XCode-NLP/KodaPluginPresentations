// The design brief the agent reads before building or restyling a deck.

import { SIZE_SCALE } from "./themes.js";

export const DESIGN_GUIDE = `
# Workflow
1. Plan before drawing. Write the storyline: audience, the one thing they must remember, then one
   message per slide. Pick a layout for each message from the catalog. Typical arc: title → agenda
   (decks over 8 slides) → problem → insight/numbers → solution → how it works → proof (quote,
   stats, chart) → plan/timeline → closing with the ask.
2. Pick a theme that matches audience and mood (set_presentation_theme). Keep it for the whole deck.
3. Build slides with add_layout_slide. Use raw objects (apply_presentation_operations) only for
   small additions; then use role names, never invented hex colors or fonts.
4. Check: run check_presentation and fix every error. Then render_slides and look at the images;
   fix what looks wrong (crowding, weak hierarchy, awkward wrapping, empty areas). Repeat once.
5. Put details and talking points into speaker notes (updateSlide {notes}), not onto the slide.

# Writing
- Titles state the conclusion ("Churn fell 30% after onboarding v2"), not the topic ("Churn").
- One idea per slide. Under 40 words on a content slide. Points ≤12 words, no full stops.
- Numbers beat adjectives: use big-number, stats or chart when there is a figure.
- Same grammatical form across points and card titles. No orphan single bullet.
- Write in the language of the conversation.

# Visual rules
- Hierarchy: at most 3–4 text sizes per slide from the theme scale.
- Color: background, text and muted carry the slide; accent marks the one thing to look at. Use
  accent2 and chart colors only for data. Never more than one accent block competing per slide.
- White space is structure: keep the 64/48-unit margins, align to existing left edges, keep equal
  gaps between repeated elements.
- Rhythm: alternate dense and airy slides; do not run three slides of the same layout in a row.
  Use section (accent tone) to separate parts in decks over 10 slides.
- Images: PNG from the project. Prefer real screenshots and product shots over decoration.
- Contrast: body text ≥ 4.5:1, large text ≥ 3:1 (check_presentation reports violations).

# Tokens
Color roles: background, surface, text, muted, line, accent, accent2, onAccent, chart1–chart5.
Font roles: heading, body. Size roles (for a 540-unit-high slide): ${Object.entries(SIZE_SCALE)
  .map(([key, value]) => `${key} ${value}`)
  .join(", ")}.
In addObject/updateObject any color, font or size field accepts a role name; objects keep the role
and follow later theme switches. A literal hex or number detaches the field from the theme.
A new object may be partial: {kind: "text", text, x, y, w, h, size: "h3", color: "accent"}.
Tones: "default", "accent" (accent background) and "inverse" (dark text-colored background) can be
passed to add_layout_slide to override a layout's own tone, e.g. a statement on an accent slide.
`.trim();
