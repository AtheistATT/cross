#!/bin/bash

if [[ -z "$GIT_USERNAME" || -z "$GIT_TOKEN" ]]; then
  echo "Ошибка: переменные окружения GIT_USERNAME и GIT_TOKEN не установлены."
  exit 1
fi

REPO_URL="https://${GIT_USERNAME}:${GIT_TOKEN}@github.com/AtheistATT/cross"

git add .

git ls-files --ignored --exclude-standard | while read file; do
  git rm --cached "$file"
done

read -p "Введите сообщение для коммита: " commit_message

git commit -m "$commit_message"

git push $REPO_URL master --force

cowsay "Коммит выполнен и изменения отправлены в репозиторий."

read -p "Нажмите Enter для выхода..."
