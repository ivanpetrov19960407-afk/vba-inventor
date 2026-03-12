# Prompt-поля штампа СПДС форма 3

В `RKM_SPDS_A3_FORM3_V17` используются prompted `TextBox` через маркеры `<Prompt>...</Prompt>`.

## Фактический порядок PromptStrings
При `Sheet.AddTitleBlock` массив `PromptStrings` формируется строго в порядке:
1. `CODE`
2. `PROJECT_NAME`
3. `DRAWING_NAME`
4. `ORG_NAME`
5. `STAGE`
6. `SHEET`
7. `SHEETS`

Порядок централизован в `GetPromptOrder()` (`src/RKM_TitleBlockPrompted.bas`).

## Правила заполнения
1. Значения из Excel (если колонка есть и ячейка не пустая).
2. Для `SHEET`/`SHEETS` при пустых значениях — авторасчёт по индексу листа и общему количеству.
3. Для остальных полей — дефолты из `DefaultPromptMap()`.

## Сопоставление PDF -> prompt-поля
- `№ док.` -> `CODE`
- Наименование объекта/проекта -> `PROJECT_NAME`
- Описание изделия/листа -> `DRAWING_NAME`
- Организация -> `ORG_NAME`
- Стадия -> `STAGE`
- Лист -> `SHEET`
- Листов -> `SHEETS`

## Почему выбран вариант через PromptStrings
Выбран путь передачи `PromptStrings` в `Sheet.AddTitleBlock`, потому что:
- порядок prompt-полей фиксирован и задаётся кодом создания `TitleBlockDefinition`;
- этот путь проще и стабильнее для идемпотентного пере-применения штампа;
- не требуется пост-обход `TextBox` и вызовы `SetPromptResultText`.
