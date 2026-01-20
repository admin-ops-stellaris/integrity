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

**At Session Start (The "Locator" Prompt)**
"Before we begin:

Read replit.md to load the current architecture and file structure.

Locate the Logic: If I ask for a feature (e.g., 'Update the email composer'), first identify which module handles that logic (e.g., email.js) and work strictly within that file. Do not dump new logic into app.js unless it is global orchestration.

Pattern Check: Remind yourself of the 'Modal Overlay' pattern and the IntegrityState object for state management. Do not create new global variables."

**Periodically / Code Health (The "Integrity" Prompt)**

"Please review the codebase for code health:

Module Boundaries: Are there any functions sitting in app.js that actually belong in a specific feature module? If so, move them.

State Hygiene: Are we correctly using window.IntegrityState for shared data, or have we accidentally created loose global variables?

Performance Check: Ensure we haven't broken the 'Lazy Loading' pattern (e.g., check that we aren't accidentally fetching full deep-graph records in simple list views).

Documentation: Compare the current file structure against replit.md and update it if we've added new modules."

**At Session End (The "Debt" Prompt)**

"We are done for now. Please scan the changes we made today:

Documentation: Update replit.md with any new features, database changes, or module additions.

Tech Debt Check: Did we leave any 'TODOs' or temporary hacks (like hardcoded IDs or bypassed checks) to get things working? List them now so I know what to clean up next time.

File Size: Briefly confirm that app.js hasn't bloated back up."