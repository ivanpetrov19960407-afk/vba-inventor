# Prompt-поля штампа: точный порядок

Inventor сопоставляет значения `PromptStrings` строго по позиции.
Ниже приведен **фиксированный порядок из 15 полей**, реализованный в модуле `RKM_TitleBlockPrompted` и в дефолтном массиве `DefaultPromptValues`.

> Язык prompt-меток: **русский** (целевой режим).

| # | Prompt label (в `<Prompt>...</Prompt>`) | Внутренний смысл | Пример значения |
|---|---|---|---|
| 1 | Заказчик | Customer | ООО Ромашка |
| 2 | Обозначение | Designation | RKM-001 |
| 3 | Описание объекта 1 | ObjectDescription1 | Жилой комплекс |
| 4 | Описание объекта 2 | ObjectDescription2 | Корпус 2 |
| 5 | Описание объекта 3 | ObjectDescription3 | Стадия П |
| 6 | Заголовок раздела 1 | SectionTitle1 | Архитектурные решения |
| 7 | Заголовок раздела 2 | SectionTitle2 | Пояснительная записка |
| 8 | Заголовок раздела 3 | SectionTitle3 | Общие данные |
| 9 | Стадия | Stage | П |
| 10 | Лист | SheetNumber | 1 |
| 11 | Листов | TotalSheets | 12 |
| 12 | Наименование листа | SheetName | Общие данные |
| 13 | Организация | Organization | ООО Проект |
| 14 | Разработал ФИО | DeveloperName | И.И. Иванов |
| 15 | Дата разработал | DeveloperDate | 01.01.2026 |

## Важно
- Не меняйте порядок без одновременного изменения:
  1. порядка `TextBoxes` с `<Prompt>` в `RKM_TitleBlockPrompted.bas`;
  2. массива в `RKM_Utils.bas::DefaultPromptValues`;
  3. этой документации.
