# VBA Inventor: СПДС A3 рамка/штамп + IDW-альбом из Excel

## Что делает проект
Репозиторий содержит исходники VBA-модулей для Autodesk Inventor, которые:
- создают/обновляют рамку A3 (`BorderDefinition`) по СПДС;
- создают/обновляют штамп формы 3 (`TitleBlockDefinition`) с prompted-полями;
- строят/обновляют альбом `ALB_*` листов в `.idw` по списку моделей;
- поддерживают сборку из Excel (`ALBUM`) и заполнение штампа по колонкам Excel.

## Структура
- `src/RKM_EntryPoints.bas` — публичные точки входа.
- `src/RKM_IdwAlbum.bas` — сборка/обновление альбома, листов и видов.
- `src/RKM_Excel.bas` — чтение Excel (late binding).
- `src/RKM_FrameBorder.bas` — рамка A3.
- `src/RKM_TitleBlockPrompted.bas` — штамп и prompt-логика.
- `src/RKM_Utils.bas` — общие утилиты, диалоги выбора файлов.
- `docs/excel-format.md` — формат Excel для альбома.
- `docs/manual-test-checklist.md` — ручной чек-лист.

## Prompt-поля штампа (фактические)
- `CODE`
- `PROJECT_NAME`
- `DRAWING_NAME`
- `ORG_NAME`
- `STAGE`
- `SHEET`
- `SHEETS`

Порядок и дефолты собраны в helper-функциях в `RKM_TitleBlockPrompted.bas`.

## Как запустить в Inventor
1. Откройте Inventor.
2. `Alt+F11` -> VBA Editor.
3. `File -> Import File...` и импортируйте все `.bas` из `src/`.
4. `Debug -> Compile VBAProject`.

### Запуск макросов
- Только рамка/штамп на активном листе:
  - `Rkm_CreateOrApplyA3Frame`
- Альбом по моделям из Workspace:
  - `Rkm_BuildOrUpdateAlbum`
- Альбом по Excel в активный DrawingDocument:
  - `Rkm_BuildAlbumFromExcel_OnActiveDrawing`
- Альбом по Excel в новый DrawingDocument + `SaveAs`:
  - `Rkm_BuildAlbumFromExcel_AndSaveAs`

## Идемпотентность
Повторный запуск:
- не плодит дубликаты рамки/штампа на листе;
- переиспользует существующие листы `ALB_*`;
- удаляет устаревшие `ALB_*` листы, которых нет в текущем источнике моделей.

## Формат Excel
См. `docs/excel-format.md`.

## Важное замечание по кодировке
В VBA-модулях русские строки стараемся хранить через `RuText(ChrW...)`, чтобы снизить риск искажений при переносе `.bas` между разными кодировками.
