# Publish to crates.io

`ferroxl` and `ferroxl-mcp` are published from CI, in that order. Both are on crates.io
now; [ferroxl](https://crates.io/crates/ferroxl) and
[ferroxl-mcp](https://crates.io/crates/ferroxl-mcp), 0.1.5 through 0.1.9.

`release.yml` and `publish.yml` are separate on purpose. The tag builds the binaries; the
published *release* publishes the crates. A failed cross-compile must not put a crate on
crates.io, and a crate is not something you want to withdraw.

## One-time setup

Create a crates.io API token with the **`publish:new`** scope on the `ferroxl-2`
organisation:

1. Sign in at <https://crates.io/settings/tokens>.
2. **New token**, name it `ferroxl-ci`.
3. Tick **publish:new**. Do not tick `publish:update` — a token that can replace a
   published version is a footgun, and none of this needs it.
4. Copy it once; crates.io shows it only at creation.

Add it as a repository secret:

```sh
gh secret set CARGO_REGISTRY_TOKEN --repo SV-stark/ferroxl
```

`scripts/publish.ps1` does the same thing from this machine if you would rather publish
locally.

## How it runs

`.github/workflows/publish.yml` fires when a GitHub release is published, so a build that
fails does not put a crate on crates.io. It can also be started by hand from the Actions
tab, which is how the first publish happened and how a half-finished one is retried.

The workflow does not check out the tag. It reads the version out of the tag name and lets
the published source be whatever produced the release; checking out would be more precise
but adds a failure mode, since a tag that cannot be fetched would stop the publish. The
`release.yml` tag-versus-`Cargo.toml` check is the guard that matters, and it runs before
either workflow.

The two crates go up one after the other, never in parallel. `ferroxl-mcp` depends on
`ferroxl` by version, and crates.io will not accept a package whose dependency it cannot
resolve — so publishing both at once fails with

```
no matching package named `ferroxl` found
```

which reads like a version number is wrong rather than a race. The workflow waits for the
registry index to list `ferroxl` before it starts on the server, for the same reason
`scripts/publish.ps1` does.

## Why a dry run comes first

`cargo publish --dry-run --locked` packages and compiles the crate exactly as an upload
would, without sending anything. It is the cheapest check that the published artifact builds
on its own rather than only inside this workspace.

## If a publish fails part-way

crates.io does not allow replacing a published version, so a retry cannot fix a bad
upload — the version has to be bumped. `scripts/publish.ps1` refuses to run against a
version that already exists rather than letting it fail obscurely.

## What is published

Only the library and the server. `tools/` and `scripts/` are outside both crate directories
and so are never packaged; `exclude` covers `.github`, `target` and
`examples/test_nrlm_read.rs`. That example is a scratch runner pointing at a local
directory, and a maintainer's absolute path is not something to publish.