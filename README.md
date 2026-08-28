TrainingSchedule

TrainingSchedule is a SharePoint Framework (SPFx) web part that provides a corporate training dashboard. It lets employees view available trainings, enroll in courses, track progress, and lets admins create and manage trainings and enrollments.

Key features
- Dashboard view of trainings and enrollments
- Employee "My Courses" view with progress tracking
- Admin board to create trainings and manage enrollments
- Modular React + TypeScript codebase
- Styles extracted into small TypeScript style modules (no SCSS changes)

Architecture
- Frontend: React components (functional components) written in TypeScript
- SPFx: Web part wiring and SharePoint access using PnPjs
- Styling: Inline styles were moved into small TypeScript modules under src/webparts/helloWorld/components/styles
- Main composition: TrainingDashboard composes smaller components: Header, TrainingList, TrainingCard, MyCourses, AdminBoard, EnrollmentModal, TrainingForm.

Tech Stack
- TypeScript (~5.x)
- React 17
- SPFx (SharePoint Framework) Web Part
- @fluentui/react for UI controls
- @pnp/sp for SharePoint REST access
- Gulp for SPFx build tasks

Repository structure (important paths)
- src/index.ts — web part entry
- src/webparts/helloWorld/ — main web part source
  - HelloWorldWebPart.ts — SPFx web part class
  - components/ — React components used by the web part
    - ui/ and component files (e.g., TrainingCard.tsx, TrainingDashboard.tsx)
    - styles/ — TypeScript style modules (e.g., TrainingCard.styles.ts)
- config/ — SPFx build configuration and deploy settings
- gulpfile.js, package.json, tsconfig.json — build and toolchain

How it was built (high level)
1. The UI was refactored from large monolithic files into many smaller components to improve readability and maintainability.
2. Large inline style objects were extracted into small TypeScript modules so the UI looks the same but styles are easier to manage.
3. TypeScript checks (npx tsc --noEmit -p tsconfig.json) were run frequently to catch typing and duplicate-definition issues.
4. Commits and pushes were made to a feature branch named sonal.

Getting started (developer)
Prerequisites:
- Node.js (14+ or 16+ depending on SPFx version)
- npm or yarn
- Gulp installed globally if you run gulp tasks directly: npm i -g gulp-cli

Install dependencies:
Run: npm install

Type-check the project:
Run: npx tsc --noEmit -p tsconfig.json

Build (SPFx bundle):
Run: gulp bundle --ship

Serve locally (if configured):
Run: gulp serve

Common developer tasks
- Run type-check: npx tsc --noEmit -p tsconfig.json
- Run build: gulp bundle
- Run local dev server: gulp serve
- Commit & push: git add -A && git commit -m "your message" && git push origin sonal

Notes about the refactor
- Style extraction was done using small TypeScript objects exported from components/styles/*.styles.ts files.
- Some dynamic style helpers in style modules were typed loosely (any) to avoid friction with Fluent UI's IStyle typing.
- Duplicate code artifacts and unused imports were removed during the process.

Where to look first
- UI composition and main logic: src/webparts/helloWorld/components/TrainingDashboard.tsx
- Training card (single item): src/webparts/helloWorld/components/TrainingCard.tsx
- Styles: src/webparts/helloWorld/components/styles/

Want me to continue?
If you'd like I can:
- Extract remaining style modules (AdminBoard, EnrollmentModal, TrainingForm, Header) the same way.
- Clean up duplicate ui/ or modular/ folders if you want a single consolidated location.

---

(Generated with assistance from an automated refactor process.)
