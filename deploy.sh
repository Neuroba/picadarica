#!/bin/bash
set -e

SITE_DIR="$(cd "$(dirname "$0")" && pwd)"
REPO_NAME="picadarica"

cd "$SITE_DIR"

# Инициализация git (один раз)
if [ ! -d .git ]; then
    git init -b main --quiet
    echo ".netlify" >> .gitignore
fi

# Создание GitHub репо + Pages (один раз)
if ! git remote get-url origin &>/dev/null; then
    GH_USER=$(gh api user --jq .login)
    gh repo create "$REPO_NAME" --public --description "PicadaRica — Premium Beef Jerky Buenos Aires" --confirm 2>/dev/null || true
    git remote add origin "https://github.com/$GH_USER/$REPO_NAME.git"
    git add -A
    git commit -m "initial deploy" --quiet
    git push -u origin main --quiet
    sleep 3
    gh api "repos/$GH_USER/$REPO_NAME/pages" \
        -X POST -f "source[branch]=main" -f "source[path]=/" --silent 2>/dev/null || true
    echo "Сайт создан: https://$GH_USER.github.io/$REPO_NAME"
    exit 0
fi

# Деплой изменений
git add -A
if git diff --staged --quiet; then
    echo "Нет изменений"
    exit 0
fi

git commit -m "update $(date '+%Y-%m-%d %H:%M')" --quiet
git push origin main --quiet

GH_USER=$(gh api user --jq .login)
echo "Задеплоено: https://$GH_USER.github.io/$REPO_NAME"
