# ================= OPERATIONS =================
cd "C:\Users\pc\PycharmProjects\IntelliBI_Operations_Automation"
Remove-Item .git\index.lock -Force -ErrorAction SilentlyContinue
git fetch origin --prune
git checkout dev  ; git reset --hard origin/dev
git checkout main ; git reset --hard origin/main
git checkout prod ; git reset --hard origin/prod
git checkout prod        # leave it on the branch this machine runs