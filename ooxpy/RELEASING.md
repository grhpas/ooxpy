# Releasing ooxpy

Releases are built and published by `.github/workflows/release.yml` with
PyPI's trusted publishing: no API token is stored in the repository or in
GitHub secrets.

## One-time setup

1. **Name.** Confirm that `ooxpy` is still free on [PyPI](https://pypi.org/project/ooxpy/)
   and [TestPyPI](https://test.pypi.org/project/ooxpy/) (a search on
   2026-11-16 found no project of that name, nor one called `oox`).
2. **Repository URL.** Once the repository is public, add to `pyproject.toml`:
   ```toml
   [project.urls]
   Homepage = "https://github.com/<owner>/ooxpy"
   Documentation = "https://<owner>.github.io/ooxpy/"
   Source = "https://github.com/<owner>/ooxpy"
   Issues = "https://github.com/<owner>/ooxpy/issues"
   ```
   and the same `repo_url`/`site_url` in `mkdocs.yml`.
3. **Trusted publishers.** On TestPyPI and on PyPI, add a *pending publisher*
   for project `ooxpy` (account → Publishing): owner and repository as on
   GitHub, workflow `release.yml`, environment `testpypi` (TestPyPI) or
   `pypi` (PyPI).
4. **Environments.** In the GitHub repository settings, create the
   environments `testpypi` and `pypi`. Requiring a reviewer on `pypi` makes
   every real upload wait for a click.

## Each release

1. **Office evidence.** The Office check corpus (`tools/officecheck/`) has
   been run on the Office versions the release claims, with the results in
   `tools/officecheck/reports/README.md`.
2. **API review.** `git diff <last tag> -- tests/api/public_api.txt` shows
   every change to the public API. Removals and renames follow
   `docs/STABILITY.md` (from 1.0: deprecated first).
3. **Version and changelog.** Set `__version__` in `src/oox/__init__.py` (the
   package version is read from there) and move the "Unreleased" entries in
   `CHANGELOG.md` under the new version and date.
4. **Local check.**
   ```bash
   python -m pytest
   ruff check src tools tests examples && pyright
   mkdocs build --strict
   python -m build && python -m twine check --strict dist/*
   ```
5. **Rehearse on TestPyPI.** Actions → Release → Run workflow (target
   `testpypi`). Then, in a fresh environment:
   ```bash
   python -m pip install -i https://test.pypi.org/simple/ --extra-index-url https://pypi.org/simple/ ooxpy==X.Y.Z
   cd /tmp && python /path/to/ooxpy/tools/smoke_install.py
   ```
6. **Publish.** `git tag vX.Y.Z && git push origin vX.Y.Z`. The workflow
   checks that the tag equals `__version__`, tests the wheel on Python 3.10
   and 3.14 (Linux and Windows) and the sdist's test suite, then uploads to
   PyPI.
7. **Documentation.** `mkdocs gh-deploy` (or a Pages workflow) publishes the
   site from `docs/`.

## What goes where

- **Wheel:** the `oox` package only: runtime, generated grammar tables and
  factories, `.pyi` stubs and `py.typed`. About 1.1 MB compressed, 8 MB installed.
- **sdist:** the package plus `tests/`, `tools/` (the tests import the
  Office-check kit), `examples/`, `docs/` and the workflows, so the test
  suite runs from it (`MANIFEST.in`).
