# Лабораторна робота №1

**Тема:** Робота з файлами та даними у C# (CSV/JSON/XML/XLSX, звіти DOCX/XLSX, графіки)

**Студент:** Мілецький Ілля Валерійович, КН-21

**Варіант / набір даних:** 37 / Air-Quality-Index, [Kaggle](https://www.kaggle.com/datasets/dnkumars/air-quality-index)

**Дата:** 30.09.2025

---

## 1. Мета та завдання

**Мета:**  
Опанувати роботу з файлами у форматах CSV/JSON/XML/XLSX, реалізувати обробку та збереження даних через інтерфейси, навчитися будувати графіки та формувати звіти у форматах DOCX та XLSX.

**Завдання:**
- Імпорт даних із CSV/JSON/XML/XLSX
- Обробка/валідація/агрегація
- Експорт у всі формати (CSV/JSON/XML/XLSX)
- Генерація звітів XLSX та DOCX
- Візуалізація (мінімум 2 графіки) зі стороннім компонентом
- Архітектура у 3 збірках (Domain/Data/UI), діаграма класів

---

## 2. Опис датасету

**Походження/джерело:** [Kaggle – Air Quality Index dataset](https://www.kaggle.com/datasets/dnkumars/air-quality-index)

**Структура полів:**

| Поле | Тип | Приклад значення | Опис/правила |
|------|-----|------------------|--------------|
| Rank | int | 1 | Порядковий номер міста за рівнем забруднення |
| CityCountry | string | "Delhi, India" | Назва міста та країни |
| AverageAQI | int | 168 | Середній індекс якості повітря (чим вище - тим гірше) |
| MonthlyData | List&lt;MonthlyAQI&gt; | {Jan=165, Feb=160...} | Показники AQI за 12 місяців |

**Обсяг:** 
- Рядків: 5377
- Стовпців: 15 (Rank, CityCountry, AverageAQI, 12 місяців)

**Проблеми якості:**
- Можливі пропуски у місячних значеннях (замінюються на 0)
- Некоректні значення (відфільтровуються)
- Дублікати відсутні
- Кодування: UTF-8

---

## 3. Модель предметної області

**Виділені сутності:**

1. **CityAirQuality** — основна сутність, що представляє місто з показниками якості повітря
   - Властивості: Rank (int), CityCountry (string), AverageAQI (int), MonthlyData (List&lt;MonthlyAQI&gt;)

2. **MonthlyAQI** — сутність для зберігання місячних показників
   - Властивості: Month (string), Value (int)

3. **CityAirQualityManager** — менеджер колекції міст для управління даними
   - Відповідає за операції над колекцією міст

**Зв'язки:**
- `CityAirQuality` **1..*** → `MonthlyAQI` (одне місто має дані за 12 місяців)

**Діаграма класів:**

<img width="489" height="232" alt="Діаграма класів" src="https://github.com/user-attachments/assets/f372e88d-f468-4bf2-9456-1953b931cd28" />

**Обґрунтування:**
Обрано таку модель, оскільки датасет логічно розділяється на міста (основні сутності) та їх місячні показники. Зв'язок 1-до-багатьох: кожне місто має рівно 12 місячних записів. Така структура дозволяє легко агрегувати дані, будувати графіки, формувати звіти та читати і записувати данні.

---

## 4. Архітектура застосунку

**Збірки/шари:**
### Domain (Core/Model)
- `CityAirQuality.cs` — основна сутність
- `MonthlyAQI.cs` — місячні дані
- `CityAirQualityManager.cs` — менеджер колекції

### Data (Infrastructure)
- **Outputs/** - папка для збереження графіків
- **Providers/**
  - `IProvider.cs` - інтерфейс для всіх провайдерів
  - `CsvProvider.cs` - робота з CSV
  - `JsonProvider.cs` - робота з JSON
  - `XmlProvider.cs` - робота з XML 
  - `XlsxProvider.cs` - робота з Excel 
  - `TextProvider.cs` - збереження логів
- **Reports/**
  - `XlsxReportService.cs` - генерація XLSX-звітів
  - `DocxReportService.cs` - генерація DOCX-звітів

### UI (Presentation)
- `MainForm.cs` — головна форма
- `ImportPreview.cs` — попередній перегляд даних перед імпортом
- `AddCityForm.cs` — форма для додавання нового міста

**Дерево рішення:**
```
/Task2
├── /Domain              # Збірка Domain.csproj
│   ├── CityAirQuality.cs
│   ├── MonthlyAQI.cs
│   └── CityAirQualityManager.cs
├── /Data                # Збірка Data.csproj
|   ├── /Outputs
│   ├── /Providers
│   │   ├── IProvider.cs
│   │   ├── CsvProvider.cs
│   │   ├── JsonProvider.cs
│   │   ├── XmlProvider.cs
│   │   ├── XlsxProvider.cs
│   │   └── TextProvider.cs
│   └── /Reports
│       ├── XlsxReportService.cs
│       └── DocxReportService.cs
└── /Task2               # Збірка UI.csproj
    ├── MainForm.cs
    ├── ImportPreview.cs
    ├── AddCityForm.cs
    └── Program.cs
```

**Залежності/NuGet:**
- **CsvHelper**
- **System.Text.Json**
- **System.Xml.Linq**
- **EPPlus**
- **DocumentFormat.OpenXml**
- **OxyPlot.WindowsForms**

---

## 5. Імпорт та чистка даних

**Підтримувані формати імпорту:** CSV, JSON, XML, XLSX

**Налаштування імпорту:**
- **Кодування:** UTF-8 за замовчуванням
- **Роздільник CSV:** кома (,), підтримка лапок для полів
- **Формат чисел:** ціле число (int) для Rank, AverageAQI та місячних значень
- **Попередній перегляд:** перші 5 рядків через форму ImportPreview

**Перетворення та валідація:**

1. **Правила валідації:**
   - Rank, AverageAQI, та місячні значення мають бути цілими числами
   - CityCountry не може бути порожнім в іншому випадку буде викликано Exception
   - Якщо було введено некоректні дані для Rank буде викликано Exception

2. **Обробка пропусків:**
   - Пропущені місячні значення замінюються на 0
   - Пропущені avg aqi підраховуються на основі місячних значень

3. **Усунення дублікатів:**
   - Не застосовується, оскільки кожне місто унікальне за Rank

**Результат "до/після":**
- До: велика кількість пропусків у місячних даних в кінці файлу
- Після: усі пропуски замінені на 0

---

## 6. Операції над даними

**Фільтрація/пошук/сортування:**
- **UI:** Панель фільтрів з полями:
  - Вибір поля (City, Country, Rank, AverageAQI)
  - Вибір умови (>, <, =, >=, <=, !=)
  - Введення значення
  - Кнопка "Застосувати"
- **Логіка:** Фільтрація на основі обраних критеріїв, а саме пошук міста або міст за країною та фільтрування на основі AverageAQI або Rank.
- **Сортування:** Клік по заголовку стовпця DataGridView. Також сортування відбувається після редагування, видалення або додавання міст, з присвоєнням відповідного Rank за вирахуваним AverageAQI.

**Групування та агрегування:**
- **Поля:** AverageAQI(середнє значення якості повітря), місячні значення AQI, та Rank
- **Метрики:**
  - Rank - місце серед інших міст, які є доданими або завантаженими в програму.
  - AverageAQI - середній AQI.

**Редагування:**
- **Додавання:** додавання міста через форму, з полями для обрання назви та aqi для кожного місяця,
- **Видалення:** кнопка "Видалити" для вибраного запису
- **Перевірки цілісності:** Валідація числових полів перед завантаженням у програму, та експортом. Також при редагуванні та додавані міста.

---

## 7. Експорт даних

**Підтримувані формати експорту:** CSV, JSON, XML, XLSX

**Параметри експорту:**
- **CSV:**
  - Кодування UTF-8
  - Роздільник: кома
  - Є поля видулені лапками
- **JSON:** 
  - Форматування з відступами
  - UTF-8
- **XML:** 
  - Структурований з тегами `<CityAirQuality>` та `<MonthlyAQI>`
  - UTF-8, автоматичні відступи
- **XLSX:** 
  - Аркуш "AirQuality"
  - Заголовки: Rank, City, Avg, Jan...Dec

**Перевірка якості експорту:**
- Усі файли зберігаються коректно
- Формати можна повторно імпортувати без втрати даних
- Контрольні приклади:
CSV

<img width="802" height="482" alt="image" src="https://github.com/user-attachments/assets/05d4c1ad-2fc3-45b7-8a8d-7ad9e7296c6a" />
<img width="611" height="484" alt="image" src="https://github.com/user-attachments/assets/9172c286-21bf-45bb-925e-002b37f8b81a" />
<img width="661" height="729" alt="image" src="https://github.com/user-attachments/assets/4079149e-3900-477b-b62e-ec5dd189e851" />

JSON

<img width="802" height="482" alt="image" src="https://github.com/user-attachments/assets/da21cefc-3422-4611-bfac-d61b99b93b02" />
<img width="611" height="484" alt="image" src="https://github.com/user-attachments/assets/ec023a0a-d76d-4ac2-bc96-e8516e466179" />
<img width="309" height="675" alt="image" src="https://github.com/user-attachments/assets/e449d56d-fff3-406a-bcbf-51f6b7ea59cb" />

XML

<img width="802" height="482" alt="image" src="https://github.com/user-attachments/assets/ad28d56b-da53-4460-9832-79dae50f8355" />
<img width="611" height="484" alt="image" src="https://github.com/user-attachments/assets/e4d4963e-f8fe-4872-a459-4a065009e39e" />
<img width="546" height="829" alt="image" src="https://github.com/user-attachments/assets/0e6f540b-abf9-4e53-a417-9c7bbd680a0b" />

XLSX

<img width="802" height="482" alt="image" src="https://github.com/user-attachments/assets/63bcd1df-add8-4614-9214-14cd5cc4c1ac" />
<img width="611" height="484" alt="image" src="https://github.com/user-attachments/assets/926f9158-0618-420c-bbba-6f24bd7b9259" />
<img width="1232" height="634" alt="image" src="https://github.com/user-attachments/assets/f30d1858-ae87-45ae-acc2-18a35ff4fda9" />

---

## 8. Звіти

### 8.1 XLSX-звіт

**Аркуш Summary:**
- Таблиця з колонками: City/Country, Average AQI, Min AQI, Max AQI
- **Умовне форматування:** Top 3 міста за Average AQI виділені кольором
- **Стилізація:** Сірий фон заголовків, жирний шрифт, авто-ширина стовпців

<img width="565" height="811" alt="image" src="https://github.com/user-attachments/assets/ef9d98cb-bd50-4813-8b65-234fa1234714" />

**Аркуш Charts:**
- Містить кругову діаграму з AverageAQI по містам для кожної країни

<img width="860" height="810" alt="image" src="https://github.com/user-attachments/assets/eb6236b4-9494-4ae8-9d9e-c4bd4654d540" />

### 8.2 DOCX-звіт
**Структура:**
1. **Титульна сторінка:**
   - ПІБ: Мілецький Ілля Валерійович
   - Варіант: 37
   - Джерело: Kaggle Air Quality Index
   - Дата створення звіту

2. **Таблиця з метриками:**
   - Перші 20 міст
   - Колонки: Rank, City/Country, Average AQI

3. **Вставлене зображення діаграми:**
   - PNG з експортованою діаграмою
   - Розмір: 6000000L x 3500000L (EMU)

4. **Висновки**
   - Медіана AQI
   - Діапазон (мін–макс)
   - Топ-5 найзабрудненіших і найчистіших міст.
   - Кількість міст із добрим / поганим повітрям за пороговими значеннями.

<img width="691" height="918" alt="image" src="https://github.com/user-attachments/assets/abaed622-33be-48a0-a439-5bc85f1637db" />

<img width="694" height="918" alt="image" src="https://github.com/user-attachments/assets/15bb7d6f-e6fd-4efe-9923-b8046c12d98d" />

## 9. Візуалізація

**Компонент:** OxyPlot

**Побудовані графіки (мінімум 2):**
### 1. Лінійна діаграма (Line Chart)
- **Що зображено:** Місячні значення AQI для обраного міста
- **X:** Місяці (Jan-Dec)
- **Y:** Значення AQI
- **Призначення:** Показати динаміку зміни якості повітря протягом року
- **Інтерпретація:** Піки вказують на найбільш забруднені місяці
   - Для Rishra, India
  
<img width="802" height="482" alt="image-14" src="https://github.com/user-attachments/assets/777c6c09-5cf8-4fd4-88ce-e6237f32f2fd" />

### 2. Стовпчикова діаграма (Bar Chart)
- **Що зображено:** Порівняння місячних AQI для міста
- **X:** Місяці
- **Y:** Значення AQI
- **Призначення:** Порівняти рівні забруднення між місяцями
- **Інтерпретація:** Візуально видно найгірші та найкращі місяці
   - Для Rishra, India

<img width="802" height="482" alt="image-15" src="https://github.com/user-attachments/assets/653da8eb-260f-4e58-bd3e-485150041c71" />

### 3. Кругова діаграма (Pie Chart)
- **Що зображено:** Розподіл середнього AQI між містами однієї країни або країнами з датасету
- **Series:** Міста Країни
- **Призначення:** Показати відносний внесок кожного міста у загальне забруднення країни або всіх країн у загальне забруднення повітря
- **Інтерпретація:** Найбільші сектори — найбільш забруднені міста
   - Для всіх країн

<img width="802" height="482" alt="image-16" src="https://github.com/user-attachments/assets/7766cdc1-31f0-49f2-bb47-06499a0b18a0" />

   - Для країни China

<img width="802" height="482" alt="image-17" src="https://github.com/user-attachments/assets/6e9a9790-c0f6-404c-a570-facf75604da5" />

**Експорт у PNG:**
   - Є можливість екпорту всіх видів графіків у PNG
   - Графіки, які були експортовані для звіту зберігаються за шляхом "\Data\Outputs\"

<img width="1571" height="1028" alt="image-18" src="https://github.com/user-attachments/assets/a3151668-21d9-4f4c-bcce-d957174a7189" />

## 10. Інтерфейс користувача

**Екрани/форми:**

### MainForm (головна форма)

- **TabControl з вкладками:**

  - **Вкладка "Головна"**

    <img width="840" height="487" alt="image" src="https://github.com/user-attachments/assets/63741274-33a4-4d5d-a213-8354ad3f6568" />

  - **Вкладка "Графіки"**

    <img width="863" height="485" alt="image-20" src="https://github.com/user-attachments/assets/4897aed1-14a7-4c54-adc4-071efee4715c" />

  - **Вкладка "Логи"**

    <img width="863" height="485" alt="image-21" src="https://github.com/user-attachments/assets/e1bdd217-fc4d-4c77-84c9-d1124d365c3d" />

### ImportPreview (попередній перегляд)

   <img width="402" height="482" alt="image-22" src="https://github.com/user-attachments/assets/2dd9b662-9873-4525-a1bc-d1959c2e1855" />

### AddCityForm (додавання міста)

   <img width="237" height="523" alt="image-23" src="https://github.com/user-attachments/assets/6f335fa1-2d29-47b8-9c39-e058137f228b" />

**Сценарії використання:**

1. **Імпорт даних:**
   - Файл → Відкрити → Вибрати файл (CSV/JSON/XML/XLSX)
   - Переглянути ImportPreview
   - Дані завантажені у таблицю

2. **Фільтрація:**
   - Вибрати поле (City/Country/Rank/AverageAQI)
   - Вибрати умову (>/</=/>=/<=/!=)
   - Ввести значення
   - Натиснути "Застосувати"

3. **Побудова графіка:**
   - Перейти на вкладку "Графіки"
   - Ввести місто або країну
   - Вибрати тип графіка
   - "Побудувати" → "Експорт у PNG"

4. **Генерація звіту:**
   - Звіти → Згенерувати XLSX/DOCX
   - Вибрати шлях збереження
   - Файл створено

**Повідомлення про помилки:**
- Некоректний формат файлу: `"Error reading file"`

<img width="580" height="88" alt="image-27" src="https://github.com/user-attachments/assets/db437dc6-bf7f-469a-b335-efc99779affa" />

<img width="796" height="474" alt="image-25" src="https://github.com/user-attachments/assets/f2f2972e-4a03-4535-abd6-592ba8120bde" />

<img width="580" height="107" alt="image-26" src="https://github.com/user-attachments/assets/0396a154-bb42-43a1-85e4-3ca21e2f9a10" />

<img width="790" height="469" alt="image-28" src="https://github.com/user-attachments/assets/2405ef65-55ba-4c7f-9bd3-eb9a1e432466" />

**Логи:**
- **Де зберігаються:** TextBox на вкладці "Логи"
- **Формат:** [{НН:mm:ss}] повідомлення
- **Експорт:** Через TextProvider у файл з фоматом txt за обраним шляхом

---

## 11. Перевірка продуктивності та надійності

**Обсяг даних:** 5377 рядків

**Час виконання операцій:**
- Імпорт CSV: ~ 0 сек

  <img width="235" height="133" alt="image" src="https://github.com/user-attachments/assets/5594c7f2-86ab-4756-a67a-1c467986ca9d" />


- Експорт XLSX: ~1 сек

  <img width="283" height="133" alt="image" src="https://github.com/user-attachments/assets/276b9991-7284-4ec3-ac0f-7513fbf39a5b" />

- Побудова графіка: миттєво
- Генерація DOCX-звіту: ~ 0.1 сек

  <img width="272" height="133" alt="image" src="https://github.com/user-attachments/assets/97bb17af-4e86-49ed-8a43-f732811a2f17" />


**Обробка великого файлу:**
- Інтерфейс може на мить підвисати під час завантаження даних

**Тестові кейси:**
1. **Некоректний формат:**
   - Спроба відкрити .txt як CSV → Exception з повідомленням
   
2. **Відсутній файл:**
   - Перевірка `File.Exists()` перед читанням
   
3. **Заблокований файл:**
   - Try/catch блоки перехоплюють IOException
   
4. **Пропуски у даних:**
   - Заміна на 0, логування у ImportPreview

---

## 12. Використання AI
- Які інструменти AI застосовано: ChatGPT
- Для чого саме: створено приклад формування діаграм OxyPlot; код адаптовано відповідно до наявного набору даних  
- Що змінено/дороблено власноруч: реалізовано побудову всіх типів діаграм (лінійна, стовпчикова, кругова) відповідно до даних, додано обробку помилок та налаштування осей і підписів
- **Підтвердження:** _“Усі фрагменти коду у фінальній версії я повністю розумію та можу пояснити на захисті.”_

---

## 13. Висновки

У рамках лабораторної роботи реалізовано повнофункціональний застосунок для аналізу якості повітря з підтримкою множинних форматів даних.

**Що вдалось:**
- Повна реалізація імпорту/експорту (CSV/JSON/XML/XLSX)
- Чиста 3-рівнева архітектура з розділенням відповідальностей
- Звіти з умовним форматуванням та вбудованими графіками
- Інтуїтивний інтерфейс з фільтрацією та візуалізацією
- Можливість перегляду логів та експорту їх

**Що покращити:**
- Додати прогрес-бар для великих файлів
- Додати можливість експорту графіків у інші формати (SVG, PDF)

**Ідеї для розвитку:**
- Інтеграція з API для отримання актуальних даних про якість повітря
- Додаткові типи візуалізацій (теплові карти, boxplot)
- Порівняльний аналіз між країнами/регіонами
