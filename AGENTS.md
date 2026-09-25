# Repository Guidelines

## Project Structure & Module Organization

This repository is a professional portfolio built with Next.js and TypeScript. Keep application code in `src/`: routes and layouts belong in `src/app/`, reusable UI in `src/components/`, shared utilities in `src/lib/`, and typed configuration in `src/data/`. Store generated images and other display content in the repository-level `assest/` directory and import assets directly into Next.js components. Use descriptive filenames such as `secure-delivery-hero.png`; do not add generic stock imagery. Co-locate small, route-specific components with their route and promote components only when reused. Place automated tests beside the code they cover or in `tests/` for end-to-end scenarios.

## Visual Assets & Content

Treat `assest/` as the canonical source for site imagery and display media. New visuals should be original, match the tokens and visual language in `DESIGN.md`, and represent the surrounding content rather than decorate it. Prefer purpose-built illustrations over stock photography. Optimize large assets before shipping, write meaningful alt text at the usage site, and remove replaced files after confirming they have no remaining references.

## Build, Test, and Development Commands

Use the scripts defined in `package.json`:

- `npm install` installs locked dependencies.
- `npm run dev` starts the local development server.
- `npm run build` creates the production build and catches compilation errors.
- `npm run start` serves the production build locally.
- `npm run lint` runs ESLint checks.
- Add `npm test` when a unit test runner is introduced; no test suite is configured yet.

The hosting configuration currently expects static output in `dist/`; keep the Next.js export settings and `.openai/hosting.json` aligned if deployment output changes.

## Coding Style & Naming Conventions

Use TypeScript, two-space indentation, semicolons, and single quotes unless automated formatting specifies otherwise. Prefer functional React components and server components by default; add `'use client'` only when browser state or APIs require it. Name components in PascalCase (`ProjectCard.tsx`), hooks with a `use` prefix (`useTheme.ts`), and utility files in camelCase. Route folders and URL segments should be lowercase kebab-case. Run the formatter and linter before submitting changes.

## Icon System

Use `@phosphor-icons/react` for every interface icon. Import from `@phosphor-icons/react/dist/ssr` in server components and from `@phosphor-icons/react` in client components. Keep icon weights and sizing consistent within each interface region, provide accessible labels for icon-only controls, and mark decorative icons with `aria-hidden="true"`. Do not mix icon libraries or hand-draw SVG paths when a suitable Phosphor icon exists.

## Testing Guidelines

Add focused tests for shared utilities, interactive components, navigation, and project-data rendering. Name unit tests `*.test.ts` or `*.test.tsx`; name browser tests `*.spec.ts`. Every bug fix should include a regression test when practical. Until coverage thresholds are configured, prioritize meaningful coverage of user-visible behavior and critical links.

## Commit & Pull Request Guidelines

No commit history is available yet. Use concise Conventional Commit messages, such as `feat: add homelab project page` or `fix: correct LinkedIn link`. Keep each commit focused. Pull requests should explain the change, list verification performed, link relevant issues, and include screenshots or recordings for visual changes. Call out configuration, dependency, accessibility, or deployment impacts explicitly.

## Security & Configuration

Never commit credentials, private hostnames, tokens, or personal infrastructure details. Keep local secrets in `.env.local`, document required variables in `.env.example`, and expose browser-side values only when intentionally prefixed with `NEXT_PUBLIC_`.
