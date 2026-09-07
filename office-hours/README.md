# GitHub Actions office hour: August and September 2026

## Session plan

- 40 minutes: product updates and three short demonstrations
- 5 minutes: GitHub Actions performance tips
- 15 minutes: Q&A

The 18-slide presentation-ready deck with formatted presenter notes and
clickable feature references is
`GitHub-Actions-Office-Hours-Aug-Sep-2026.pptx`.

## Demo 1: reusable workflow identity

1. Open `.github/workflows/office-hours-demo.yml` and `.github/workflows/reusable-source.yml`.
2. In the repository **Actions** tab, select **Office hours demo** and choose **Run workflow**.
3. Open the `reusable-identity` job summary.
4. Compare `github.workflow_ref`, which identifies the caller, with `job.workflow_ref`, which identifies the reusable workflow defining the job.

**Point to land:** a shared workflow can now report and validate its own repository, file, ref, and SHA without caller-provided inputs.

## Demo 2: least-privilege Dependabot access

1. In `.github/workflows/office-hours-demo.yml`, highlight the job-level permission:

   ```yaml
   permissions:
     contents: read
     vulnerability-alerts: read
   ```

2. Open the `dependabot-alerts` job summary from the same workflow run.
3. Show that the workflow can count open Dependabot alerts with `GITHUB_TOKEN`; no PAT or broad `security-events` permission is needed.

**Prerequisite:** Dependabot alerts must be enabled for the repository. An empty result is still a successful demonstration. If organization policy blocks access, use the YAML as a code walkthrough and explain the policy boundary.

## Demo 3: Code Quality attribution

1. In `.github/workflows/office-hours-demo.yml`, highlight the `actions: read` permission on the `code-quality-attribution` job.
2. Open the job summary and compare the two actors and expected dynamic workflow paths:
   - Code Quality: `github-code-quality` and `dynamic/github-code-quality/codeql`
   - Code scanning: `github-advanced-security` and `dynamic/github-code-scanning/codeql`
3. Point out that usage reports, dashboards, and API consumers can now distinguish quality analysis from security scanning.

**Prerequisite:** GitHub Code Quality must be enabled and have completed a run to show an observed Code Quality path. A zero count demonstrates the query and expected attribution; keep a screenshot or successful run from an enabled repository as a fallback.

## Pre-session checklist

- Run the workflow once before the session.
- Keep the successful run open in a browser tab as a fallback.
- Confirm Dependabot alerts are enabled for the repository.
- Confirm GitHub Code Quality has a completed run, or prepare a fallback screenshot.
- Open both workflow files in adjacent editor tabs.
- Set browser and editor zoom to at least 125%.
- Download a local copy of the deck.

## Sources

- [Early September 2026 GitHub Actions updates](https://github.blog/changelog/2026-09-03-github-actions-early-september-2026-updates)
- [Actions retention expansion](https://github.blog/changelog/2026-08-27-actions-retention-will-cover-checks-workflow-runs-and-statuses)
- [Windows 11 Arm64 Visual Studio 2026 image](https://github.blog/changelog/2026-08-20-windows-11-arm64-vs2026-image-generally-available)
- [Separate GitHub Actions path for GitHub Code Quality](https://github.blog/changelog/2026-08-20-separate-github-actions-path-for-github-code-quality)
