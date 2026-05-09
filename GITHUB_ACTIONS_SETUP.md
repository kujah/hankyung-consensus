# Hankyung Consensus Automation

## What this does

- Runs `stocks.py` on a GitHub Actions schedule.
- Uses the latest available Hankyung Consensus report date by default.
- Publishes `reports_mobile/index.html` to GitHub Pages for phone viewing.
- Uploads PDFs, JSON files, and Excel files as a workflow artifact.

## Files

- Workflow: `.github/workflows/hankyung-consensus.yml`
- Python dependencies: `requirements.txt`
- Mobile output: `reports_mobile/index.html`

## One-time GitHub setup

1. Push this folder to a GitHub repository.
2. In the repository, add a secret named `OPENAI_API_KEY`.
3. In repository settings, enable Pages and set the source to `GitHub Actions`.
4. Run the workflow once with `workflow_dispatch` to verify permissions and the first deployment.

## Schedule

- The workflow runs at `22:15 UTC` on `Sunday-Thursday`.
- In Korea time, that is `07:15 KST` on `Monday-Friday`.

## Manual run

- Open the Actions tab.
- Run `Hankyung Consensus Daily`.
- Optionally set `target_date` as `YYYY-MM-DD`.

## How to check on phone

- After a successful deployment, open the repository's GitHub Pages URL.
- The page shows the mobile summary UI from `reports_mobile/index.html`.

## How to download the full output

- Open the latest workflow run in GitHub Actions.
- Download the artifact named `hankyung-consensus-output`.

## Notes

- GitHub Actions does not persist files between runs unless you upload them as artifacts or commit them.
- The mobile page is rebuilt each run and deployed from the latest output.
