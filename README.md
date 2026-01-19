# Deployment
Here's how to take the updated version through to being live for users:
1. Make changes in Replit until ready to deploy new version.
2. In Replit Shell: scripts/push-staging.sh "Describe what changed"
   - this commits changes on staging, pushes to GitHub which then auto-deploys to Fly staging
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

# Prompts for code health
- Periodically (every few major features):
   - Review the codebase for code health. Are there duplicated patterns that should be consolidated? Functions over 50 lines that should be split? Inconsistent patterns that should be standardized? Finally, compare the current file structure and logic patterns against replit.md. If we have introduced new architecture (like new modules or service files) that isn't documented there, please rewrite the relevant sections of replit.md to keep it accurate.
- At session start (when sitting down to work):
   - Before we begin, strictly read replit.md to understand the current architecture, file structure, and preferred patterns (like the Modal Overlay system). Do not deviate from these patterns. If you see a code request that conflicts with the documented architecture, warn me before proceeding.
- At session end (when wrapping up a session):
   - We are done for now. Please scan the changes we made today. If we added any new features, database tables, or changed how a core system works, please update replit.md with a summary so the next session starts with fresh context.