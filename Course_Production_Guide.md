# ICC Lesson Production Guide

## Directory Rules

Each lesson should live in its own folder. Each folder should contain only that lesson's HTML file and lesson-specific assets if needed.

Recommended structure:

```text
ICC_LessonPrep/
  assets/
    common.css
    logo.svg
  templates/
    lesson_template.html
  Lesson01/
    lesson01.html
  Lesson02/
    lesson02.html
  Lesson03/
    lesson03.html
  Lesson04/
    lesson04.html
```

## Shared vs. Lesson-Specific Changes

Shared visual issues go in `assets/common.css`:
- Side navigation
- Logo position
- Page numbers
- Background texture
- Typography scale
- Cards, buttons, images, question pages, and cover styles
- Print styles
- Narrow-screen preview styles

Lesson-specific issues go in the relevant lesson HTML:
- Lesson theme
- Lesson route
- Classroom stories
- Case content
- Student tasks
- Lesson outputs
- Homework
- Lesson navigation anchors

## Encoding Safety Rules

Chinese lesson HTML files must stay UTF-8. Treat encoding as a high-risk rule because one unsafe write can turn the full lesson into mojibake.

- Prefer `apply_patch` for small HTML edits, especially `<title>`, headings, copy, and navigation labels.
- Do not use PowerShell `Set-Content` / `Out-File` / line-array rewrites on a whole HTML file unless the command explicitly reads and writes UTF-8.
- If a whole-file rewrite is unavoidable, use an explicit UTF-8 writer and verify immediately with `Get-Content -Encoding UTF8`.
- After editing a Chinese HTML file, check the first screen text and the browser tab title for mojibake before continuing.
- If mojibake appears, restore from git or the latest clean copy first, then reapply only the intended small edit.

## Creating a New Lesson

1. Copy `templates/lesson_template.html`.
2. Create a new folder, such as `Lesson03/`.
3. Rename the copied file to `lesson03.html`.
4. Update `<title>`, side-navigation title, cover title, and lesson route.
5. Replace placeholder content in the template.
6. Remove unused pages or duplicate existing page patterns when useful.
7. Check every `href="#..."` in the side navigation has a matching `id="..."`.
8. Open the HTML preview and check text, images, logo, page numbers, and scrolling navigation.

## Recommended Lesson Structure

Each lesson does not need the same number of pages, but it should usually include:
- Cover: topic and one-sentence goal
- Core question page: one big question for classroom discussion
- 停顿回答页：先展示问题或观察任务，参考分析放在下一页。
- Concept or method page: explain the key method clearly
- HMW bridge page when relevant: change the main problem into a "How Might We..." question before asking for solutions
- Case or story page: show a real scenario
- Student work page: give students time to think, write, discuss, or collect evidence
- Output or homework page: connect naturally to invention log, novelty search, making, or presentation materials

## Avoid

- Do not regenerate a completed lesson from old outer files.
- Do not write large blocks of shared CSS inside a single lesson HTML.
- Do not merge all lessons into one HTML file.
- Do not put teacher preparation notes directly on student-facing slides.
- Do not put a question and its full reference answer on the same slide when students should answer first.

## Current Six-Lesson Mapping

- 5.1 Lesson 01: AI-era problem discovery, invention and innovation, scenario pain points.
- 5.1 Lesson 02: HMW bridge, first principles, invention topic draft, and AI-assisted questioning.
- 5.2 Lesson 03: how to learn technology in the AI era, hardware sensing basics.
- 5.2 Lesson 04: simple embedded systems and simple modeling exercises.
- 5.3 Lesson 05: Taobao material selection, parts, and BOM.
- 5.3 Lesson 06: project system design, solution presentation, and making tasks.

Follow-up sessions on 5.16, 5.23, and 5.30 can continue using the same folder pattern.
