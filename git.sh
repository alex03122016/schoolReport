#!/usr/bin/env bash

git status
git add .
git commit -m "added Fehlzeiten for example only Deutsch"
git remote add origin https://github.com/alex03122016/schoolReport.git
git push -u origin master
