# Contributing to clear-excel

Thanks for helping improve `clear-excel`. Bug reports, focused fixes, tests,
and documentation improvements are welcome.

## Before opening an issue

- Search existing issues first.
- Use the appropriate issue form.
- Do not upload confidential or production workbooks. Prefer a minimal,
  generated fixture that reproduces the behavior.
- Report suspected vulnerabilities privately as described in
  [SECURITY.md](./SECURITY.md).

## Development setup

Requirements:

- Node.js 22 or newer;
- pnpm 11 (the exact version is pinned in `package.json`).

```sh
git clone https://github.com/Vlad1slav1vanov/clear-excel.git
cd clear-excel
corepack enable pnpm
pnpm install --frozen-lockfile
pnpm run check
```

Useful commands:

```sh
pnpm test
pnpm run check:types
pnpm run lint
pnpm run format
pnpm run build
pnpm run test:package
```

`pnpm run build` cleans `dist` and uses one NodeNext TypeScript compilation to
emit ESM modules, declarations, and JavaScript maps with embedded sources.
Run it before `pnpm run test:package`, which installs the tarball in an isolated
consumer, exercises both APIs on all three supported formats, and checks a
NodeNext TypeScript consumer against the published declarations.

Add regression tests for behavior changes. Tests should use generated fixtures
where practical so the repository never contains real user workbooks or
sensitive spreadsheet data.

## Pull requests

Keep each pull request focused on one change. In the description, explain the
problem, the chosen approach, compatibility or security implications, and the
commands used for verification. `pnpm run check` must pass before review.

Changes to public exports, options, errors, supported formats, or minimum
runtime versions must also update the README and changelog.

## Release process

Releases follow Semantic Versioning and are published automatically after a
version change reaches `main`:

1. Update the version in `package.json` and `pnpm-lock.yaml`.
2. Move the release notes into `CHANGELOG.md` and run `pnpm run check`.
3. Merge the release commit into `main`.
4. The publish workflow checks whether that exact version exists in npm. New
   versions are verified and published with a provenance attestation; existing
   versions are skipped.
5. After a successful publish, create a GitHub Release tagged
   `v<package-version>` to keep the repository release history complete.

The npm package must trust `.github/workflows/release.yml` in
`Vlad1slav1vanov/clear-excel`. Configure this once in the package's npm Trusted
Publisher settings with permission to publish.

No long-lived npm token is required or expected in GitHub Actions.
