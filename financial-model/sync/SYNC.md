# Как обновить снимок

Скрипт `sync_snapshot.sh` **только читает** локальный проект и пишет только в указанную папку назначения. Он ничего не меняет в `D:\TableMInd`, не трогает его git-репозитории и рабочие деревья.

1. Сделайте свежий клон репозитория **вне** `D:\TableMInd\TableMind_EB` и его деревьев, например `D:\TableMInd\_gh_upload\project-management-tablemind`:
   `git clone https://github.com/Asteliaa/project-management-tablemind.git`
2. Создайте ветку: `git checkout -b financial-model-snapshot-YYYYMMDD`
3. В Git Bash выполните:
   `bash financial-model/sync/sync_snapshot.sh /d/TableMInd financial-model`
   Скрипт пересоздаёт `model/`, `journals/`, `scripts/`, `governance/`, `audit/`, `MANIFEST.sha256`, `SNAPSHOT_TIME.txt`. Отсутствующие в источнике файлы исчезнут из снимка — это нормально, история остаётся в git.
4. Проверьте `git status` и `git diff --stat`; убедитесь, что в изменениях нет lock/tmp-файлов и посторонних данных.
5. Закоммитьте, отправьте ветку и откройте PR (прямой push в `main` не используется).
6. Обновите таблицу «Статус снимка» в `README.md`, если изменилось, какие блоки заполнены (смотрите `governance/PROJECT_STATE.md`).

Что копируется: `*_TableMind.xlsx/.docx` из `TableMind_Model/TableMind/` и `_recalc/`; файлы из `_журналы/` (кроме `*.tmp`); `*.py/*.sh` из `_scripts/`; `CLAUDE.md`, `PROJECT_MAP.md`, `PROJECT_STATE.md`, `OPEN_QUESTIONS.md`, `TableMind_Model/README.md`. Что не копируется: `Модель (1)`, `_backup`, `.~lock*`, `lu*.tmp`, `TableMind_EB`.
