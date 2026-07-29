# Releasing

Releases are cut by pushing a `v*` tag. The `publish` job in
`.github/workflows/CI.yml` then uploads to PyPI, but only after `lint`,
`test` (5 platform/Python combinations), `msrv`, all three wheel jobs and
`build-sdist` have passed.

## One-time setup: PyPI Trusted Publishing

**This must be done before the first release, and it is the only step that
cannot be automated from CI.**

The workflow uses [Trusted Publishing][tp], so there is **no API token to
create, store or rotate** — PyPI verifies a short-lived OIDC token that
GitHub mints for this specific repository and workflow. Nothing is added to
this repository's Actions secrets.

Because `rvgsrust-xlsxwriter` does not exist on PyPI yet, register a
*pending* publisher:

1. Sign in to PyPI and go to
   <https://pypi.org/manage/account/publishing/>
2. Under "Add a new pending publisher", fill in:

   | Field | Value |
   |---|---|
   | PyPI Project Name | `rvgsrust-xlsxwriter` |
   | Owner | `Godse4ever` |
   | Repository name | `rvgsrust-xlsxwriter` |
   | Workflow name | `CI.yml` |
   | Environment name | *(leave blank)* |

3. Save. The project is created automatically on the first successful
   upload.

Leave the environment field blank — the `publish` job does not declare a
GitHub environment, and PyPI requires the two to agree.

[tp]: https://docs.pypi.org/trusted-publishers/

## Cutting a release

1. **Bump the version in all three files.** They must agree; a release has
   already shipped where they did not:
   - `Cargo.toml` → `version`
   - `pyproject.toml` → `version`
   - `python/rvgsrust_xlsxwriter/__init__.py` → `__version__`
2. Add a `CHANGELOG.md` entry.
3. Merge to `main` and confirm CI is green.
4. Tag and push:

   ```bash
   git tag -a v0.2.1 -m "v0.2.1"
   git push origin v0.2.1
   ```

The `publish` job verifies the tag matches all three declared versions
before uploading, and fails the release if any disagree. **A PyPI version
number can never be reused, even after deletion**, so a mismatched tag
would burn it permanently — hence the guard.

## If `publish` fails

The tag does not need to be recreated. Fix the cause and re-run the failed
job from the Actions UI, or:

```bash
gh run rerun <run-id> --failed
```

Common causes:

- **`Trusted publishing exchange failure`** — the pending publisher above
  was not registered, or one of its fields does not match exactly
  (including the blank environment).
- **Version mismatch** — the guard fired; the error names which file
  disagrees.
- **Missing artifacts** — an upstream build job failed, or an
  `actions/upload-artifact` call timed out. The latter is transient
  infrastructure; just re-run.
