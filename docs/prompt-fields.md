# Prompt-поля штампа СПДС форма 3

В текущей версии `RKM_SPDS_A3_FORM3_TITLEBLOCK` использует prompted `TextBox` через маркеры `<Prompt>...</Prompt>`.

При вставке штампа (`Sheet.AddTitleBlock`) макрос обязательно передаёт `PromptStrings` в корректном порядке (по порядку добавления prompted TextBox в определении):
1. `DOC_NAME`
2. `OBJ_NAME`
3. `STAGE`
4. `SHEET`
5. `SHEETS`

Если изменить состав/порядок prompted-полей в `TitleBlockDefinition.Edit`, нужно синхронно обновить массив в `BuildSpdsPromptStrings`.
