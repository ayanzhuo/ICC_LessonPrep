# ICC Lesson Prep

## Current Structure

- `Lesson01/lesson01.html`
  - 5.1, Lesson 01.
  - Focus: problem awareness in the AI era, invention and innovation, scenario cards, pain points, and four-question refinement.
- `Lesson02/lesson02.html`
  - 5.1, Lesson 02.
  - Focus: design thinking, first principles, invention topic draft, AI-assisted questioning, novelty search, invention log, and presentation.
- `assets/`
  - Shared local assets and styles.
  - `common.css`: shared layout, navigation, logo, typography, print, and responsive styles.
  - `logo.svg`: course logo.
- `templates/`
  - `lesson_template.html`: base template for later lessons.
- `Course_Production_Guide.md`
  - How to add lessons and separate shared changes from lesson-specific changes.
- `Course_Overview.md`
  - Overall plan for the six lessons and follow-up project sessions.

## Project Directions

Current group project directions:
- Refrigerator food detection
- Orchard fruit disease inspection
- Office ineffective lighting monitoring
- Vehicle blind-spot visual reminder
- Rare bird smart observation system

## Collaboration Rules

- When improving Lesson 01, edit only `Lesson01/lesson01.html` unless a shared visual rule requires `assets/common.css`.
- When improving Lesson 02, edit only `Lesson02/lesson02.html` unless a shared visual rule requires `assets/common.css`.
- For shared visual issues, update `assets/common.css`.
- For new lessons, copy `templates/lesson_template.html` into a new folder such as `Lesson03/lesson03.html`.
- Do not regenerate existing completed HTML files from older outer files.
- Older files outside `ICC_LessonPrep` should be treated as history copies.
- Keep Chinese HTML files in UTF-8. Use `apply_patch` for small text edits and avoid whole-file PowerShell rewrites unless UTF-8 read/write is explicit and verified.
