# Contributing to SOC Cockpit

Thanks for your interest in contributing! This document explains how to
propose changes.

## Developer Certificate of Origin (DCO)

All contributions must be signed off under the
[Developer Certificate of Origin](https://developercertificate.org/). Add a
`Signed-off-by` line to each commit:

```
git commit -s -m "Your commit message"
```

By signing off, you certify that you wrote the change or otherwise have the
right to submit it under the project's Apache-2.0 license.

## Development setup

Prerequisites: Node.js >= 22.5.0 (the server uses the built-in `node:sqlite`
module).

```bash
npm install
cp .env.example .env   # then fill in values as needed
npm run dev            # starts the Vite web app and the API together
```

- Frontend dev server: http://127.0.0.1:5173/
- API server: http://127.0.0.1:3001/

To produce a production build and run it:

```bash
npm run build
npm start
```

## Coding guidelines

- Keep changes focused; avoid unrelated refactors in the same PR.
- Match the existing code style. Run `npm run build` before submitting to
  ensure the TypeScript project compiles.
- Do not commit secrets, `.env`, or contents of `data/`, `uploads/`, or
  `artifacts/`.

## Submitting changes

1. Fork the repository and create a feature branch.
2. Make your change with clear, signed-off commits.
3. Ensure the build passes.
4. Open a pull request describing the change and its motivation.

## Reporting issues

Use GitHub Issues for bugs and feature requests. For security issues, follow
[SECURITY.md](SECURITY.md) instead of opening a public issue.
