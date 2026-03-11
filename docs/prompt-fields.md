# Prompt-поля штампа

В текущей версии макроса prompt-поля не используются.

Штамп `RKM_A3_TITLEBLOCK` вставляется со статическими подписями.
Если потребуется вернуть prompt-поля, нужно:
1. Добавить `TextBoxes.AddByRectangle(..., kPromptedEntryText)` в `TitleBlockDefinition.Edit`.
2. Передавать массив `PromptStrings` в `Sheet.AddTitleBlock`.
