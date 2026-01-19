# integrity
Here's how to take the updated version through to being live for users:
1. Make changes in Replit.
2. In Replit Shell: scripts/push-staging.sh "Describe what changed"
   - this commits changes on staging, pushes to GitHub when then auto-deploys to Fly staging
   - can watch the process at https://github.com/admin-ops-stellaris/integrity/actions/workflows/deploy-staging.yml
   - (just click on the latest workflow, then click on deploy)
3. Test the staged version at https://integrity-staging.fly.dev
4. When ready to go live go to https://github.com/admin-ops-stellaris/integrity/pulls
   - New pull request
   - base:main compare:staging
   - Title "Deploy staging to main" (or whatever)
   - Create pull request
   - Merge pull request
   - Confirm merge
   - can watch the log of it going live at https://fly.io/apps/integrity-prod/monitoring
5. To view the live version go to https://integrity-prod.fly.dev

**Prompts for code health**
Periodically (every few major features):
"Review the codebase for code health. Are there duplicated patterns that should be consolidated? Functions over 50 lines that should be split? Inconsistent patterns that should be standardized?"
At session start (when continuing work):
"Before we begin, please check replit.md for current architecture. Follow existing patterns. If you see something that should be refactored, mention it before proceeding."