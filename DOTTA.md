# dotta fork

This fork of [openooxml/betteroffice](https://github.com/openooxml/betteroffice) ships the
pptx editor and the `betteroffice-pptx` Python binding that dotta pins. Nothing here is
published to npm or PyPI: every build is a GitHub Release of this repository.

## Branches

- `main` mirrors upstream `main` and carries no fork changes.
- `dotta/main` is upstream plus a short stack of fork commits. Each one is a single commit
  whose subject ends in a label and whose body has a `Dotta-Patch:` line:
  - `TO-UPSTREAM`: meant for upstream, not sent yet.
  - `SENT #<n>`: open upstream as openooxml/betteroffice#n.
  - `DOTTA-ONLY`: fork tooling (this file, `dotta/`, `.github/workflows/dotta.yml`).

List the stack with `git log --grep='^Dotta-Patch:' --oneline`. The commit below the stack is
the upstream base, and every release names it in its notes.

## Taking an upstream release

1. `git fetch upstream` and fast-forward `main` to the upstream release commit.
2. `git checkout -b dotta/rebase-<upstream version> dotta/main`, then
   `git rebase --onto <upstream commit> <old base>`; resolve conflicts commit by commit and
   keep one commit per change.
3. Drop a fork commit when upstream has the same behaviour: its tests pass on upstream
   alone, or upstream merged the PR named by `SENT #n`. Relabel a commit `SENT #n` when its
   upstream PR opens.
4. Push the rebased stack as `dotta/rebase-<upstream version>` and run the `Dotta` workflow
   on it (workflow_dispatch). Once green, a maintainer replaces `dotta/main` with it using
   `git push --force-with-lease`; this is the only force push, and earlier releases stay
   reachable through their tags.
5. Tag `pptx-v<upstream pptx version>-dotta.<n>` on `dotta/main` (`n` counts from 1 for each
   upstream version) and push the tag. The workflow checks the tag against
   `packages/pptx/package.json`, reruns the gate and creates the release.

## What CI runs

`.github/workflows/dotta.yml`, on pushes and PRs to `dotta/main` and on release tags:

- Rust: fmt, clippy and tests for the pptx and shared OOXML crates, plus the Python binding.
- TypeScript: the pptx Wasm build, a check that the committed `.d.ts` match it, typecheck
  and tests for `pptx-i18n`, `pptx`, `pptx-react` and `shared`.
- Python: the binding's tests and `dotta/roundtrip`, which builds a deck with python-pptx,
  edits it with the binding and with the editor engine, and reads it back with python-pptx
  (font names, inherited placeholders, merged cells, column widths, row heights, untouched
  parts byte for byte, a layout ellipse drawn as an ellipse).

Tool versions are pinned at the top of the workflow: Rust, bun, wasm-pack, maturin, Python,
and binaryen through `.github/actions/wasm-toolchain`. Upstream's own workflows expect
upstream's runners; disable them in this repository's Actions settings.

## Using a release from dotta

Each release carries `@betteroffice/pptx`, `@betteroffice/pptx-react` and
`@betteroffice/pptx-i18n` as `.tgz` (version `<upstream>-dotta.<n>`), manylinux wheels of
`betteroffice-pptx` for x86_64 and aarch64 (upstream binding version), and `SHA256SUMS`.

```jsonc
// package.json
{
  "dependencies": {
    "@betteroffice/pptx-react": "https://github.com/yokasaka/betteroffice/releases/download/pptx-v0.2.0-dotta.1/betteroffice-pptx-react-0.2.0-dotta.1.tgz"
  },
  "overrides": {
    "@betteroffice/pptx": "https://github.com/yokasaka/betteroffice/releases/download/pptx-v0.2.0-dotta.1/betteroffice-pptx-0.2.0-dotta.1.tgz",
    "@betteroffice/pptx-i18n": "https://github.com/yokasaka/betteroffice/releases/download/pptx-v0.2.0-dotta.1/betteroffice-pptx-i18n-0.2.0-dotta.1.tgz"
  }
}
```

`overrides` works for npm and bun; pnpm uses `pnpm.overrides`, yarn uses `resolutions`.

```text
# requirements.txt for the Lambda image (pick the wheel for its architecture)
betteroffice-pptx @ https://github.com/yokasaka/betteroffice/releases/download/pptx-v0.2.0-dotta.1/betteroffice_pptx-0.2.0-cp39-abi3-manylinux_2_17_x86_64.manylinux2014_x86_64.whl#sha256=<from SHA256SUMS>
```

The repository is public, so the URLs need no token.
