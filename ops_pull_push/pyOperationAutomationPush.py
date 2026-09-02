cd "C:\Users\vaibh\Documents\IntelliBI Automation\IntelliBI_Operations_Automation"
$msg = "Operations: <describe your changes here>`n`nCo-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>`nClaude-Session: https://claude.ai/code/session_01B7oes8NoE3Q9mKZgP1qgis"

Remove-Item .git\index.lock -Force -ErrorAction SilentlyContinue

git checkout dev
git add -A
git reset -q -- "*_to_delete/*"          # keep backup folders out of the commit
git commit -m $msg
git push origin dev

git checkout main
git merge dev --no-edit
git push origin main

git checkout prod
git merge main --no-edit
git push origin prod

git checkout dev