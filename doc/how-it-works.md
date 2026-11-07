**English** | [한국어](how-it-works.ko.md)

# How SlideFill works

## One path for preview and report

Preview and Generate both call `core.run`. Preview fills a copy in memory and throws it away. So the list you check is what the report will hold, not an estimate.

## Placeholders split across runs

PowerPoint often stores `{{Sales}}` as several pieces of text, such as `{{`, `Sal`, `es}}`, because of spell checking or edits. SlideFill joins the adjacent pieces in a paragraph, finds the placeholder, writes the value into the piece where it starts, and removes what is left. The value takes the formatting of that first piece.

## Showing numbers as Excel does

A cell stores `0.853` and a format such as `0.0%`. Excel shows `85.3%`. `numfmt.py` reads the format code and does the same, rounding half up on the exact decimal text, as Excel does (`2.675` with `0.00` shows `2.68`).

Some of this depends on the PC:

- `/` in a date format means "the system's date separator". On Korean Windows, it shows as `-`.
- Built-in format 38 shows `-5` on Korean Windows and `(5)` on US Windows. The file only stores the number 38.
- Built-in format 14 follows the system's short date.

SlideFill reads these from Windows, so it matches Excel on the same PC.

If a format is not understood, SlideFill uses the plain value and warns. It does not guess.

A format can also colour a section, as `#,##0;[Red]-#,##0` shows a loss in red. The colour goes on the slide with the value: the placeholder's run is split so only the value is red, and the words around it keep their look; in a table the cell's text takes it. `[Blue]`, `[Green]` and the other named colours, `[Color1]` to `[Color56]` from Excel's standard palette, and the Korean names (`[빨강]`) are read. `[Black]` and `[White]` are left to the template, since white on a white slide would hide the number. A value put into a sentence is not coloured: only the cell's own text carries it.

Numbers are always written with `,` and `.`, whatever the PC's region. Excel draws them in the region's marks, but the file stores no marks at all, and SlideFill's sentences use `,` and `.`. A report that changed with the PC it was made on could not be checked against another.

## SmartArt

SmartArt keeps its words in a data part, and PowerPoint draws it from a second copy, sized for the words it had. SlideFill fills the placeholders in the data part, and then drops the drawn copy of that SmartArt. With no drawn copy, PowerPoint lays the SmartArt out again from its data when it opens the report, so a long value fits its box (checked with `tools/smartart_check.ps1`). Kept, the old copy would show the value at the size picked for the short placeholder, spilling out of the box. A SmartArt with nothing to fill is left exactly as it was. A drawn copy that cannot be matched to its data part is filled too. A placeholder that sits outside the text, where no value can go, is reported with its slide.

## Names that grow with their data

A defined name such as `OFFSET(Sheet!$B$1,COUNTA(Sheet!$B:$B)-1,0)` or `Sheet!$B$2:INDEX(Sheet!$B:$B,COUNTA(Sheet!$B:$B))` is worked out the way Excel would: `OFFSET`, `INDEX`, `COUNTA`, `COUNT`, `ROWS`, `COLUMNS`, whole numbers, `+` and `-`. In a range of one row, `INDEX(B2:G2, 3)` counts across the row as Excel does, and a number outside the range is not read. Anything else in a name, such as `SUM`, is reported as a name that is not cells.

## When a value cannot be determined

A missing name, a missing sheet, a formula with no saved result, or a table whose columns do not match are all errors. The placeholder or table stays exactly as it was, so a gap can be seen in the report and cannot pass as data. An empty cell is a warning: blank is the true value, but you may want to know.

## Finding the newest rows

`{{latest:Daily!A,C}}` picks a row by the date in column A, not by its position.

1. Every cell in the date column is read. A cell counts when it is a real Excel date or text that spells a full date or a year and month. A year alone, such as `2021` or `2021년`, counts only in a column that is mostly years, because it is also an ordinary number. Rows such as `소계` and `총계` drop out here.
2. A row must also hold a value in the columns asked for. A date filled in ahead with nothing beside it is not the latest. Where those columns hold figures, a row with only text is not filled either: a sheet that prints each coming month and its queue name ahead has no data for them yet. A formula with no saved result does count, so it is reported rather than quietly replaced by an older row. If a newer row has figures in other columns only, it is still being filled in: the older value is used with a warning, because the report may then mix two days.
3. Rows are sorted by date, so a sheet with the newest row on top or in mixed order gives the same answer.

If the date is only on the first row of each group, as with merged cells, the rows below it are not read, and SlideFill warns when it sees this in two or more places.

A workbook that was not updated still gives its newest row, and the list shows that date so you can spot it.

## Figures SlideFill works out

`facts.py` holds the arithmetic. It uses no model and no network, so the same workbook always gives the same answer.

- **Same weekday last week** looks for the row exactly seven days earlier. A nearby day would compare a Monday with a Sunday, which mostly measures the weekend.
- **The recent mean** leaves out the newest value, so a value is not compared with itself.
- **A change from 0** has no meaning, so it is reported instead of shown.
- **Monthly, quarterly and yearly data** get no weekly or weekday figures, and their dates read as months, quarters or years.
- **The same date on two rows** stops every comparison, wherever it is in the sheet. Two tables on one sheet would otherwise be compared with each other. The date and how many are affected are listed.
- **A period not over when the file was saved** gets a warning. The month, quarter or year must have ended by the day the file was last saved. A daily value entered on its own day only warns if it is also at least 15% below its weekday's usual level, so an ordinary evening report does not warn every day.
- **Part of a year among whole years**, such as `2022.10.` after `2021`, gets a warning, and sentences compare it per month instead.
- **A figure in a sentence** uses the cell's own format. A cell with no format gets a sensible number of digits, while the value itself still shows as Excel shows it.

## Data as people really keep it

Nothing below is guessed: each one is something you write in the placeholder. Rows are only added when you ask.

| Layout | Write | Example |
|---|---|---|
| Several rows per date | `sum(D:E)`, `count(A)`, `sum(D*F)` | Monthly passengers by direction |
| Only some rows | `sum(C, B=name)`: rows whose B is exactly the name, or contains it if no cell is exact | One queue of several |
| A mean, a share of rows | `mean(S)`, `rate(B=Yes)` | Mean pay and attrition by department |
| Year and month in separate columns | `yearColumn+monthColumn`: `A+B,D` | A table kept as year, month, value |
| Periods across a row | Row numbers: `Sheet!5,7` | A year laid out in columns |
| Two tables on one sheet | A range: `A10:A475,C10:C475` | An index and its revisions |
| One series on many sheets | `Year*!A,C`, since `*` cannot be in a real sheet name | One sheet per year |
| No dates | The same `reading`, with `top5` and `share5` for tables | Revenue by line |

Dates are also read as `2026 Aug`, `Aug. 2026`, `Aug-26`, `26년 8월`, `2026 Q3`, `3Q26`, `01월` under a merged `2024년`, and `January` under a heading row `Year 2026`. Months with no year (`1월` to `12월`, each once and in order) take the one year in the sheet's title or name. That year is dropped when months that had not begun on the day the file was saved already hold figures: `tourStat_2021-07-28`, saved that day with August to December filled, was named when it was exported, not for the year of its months. With no year, they are a series named as the sheet writes them (`12월`), never given a year no one wrote. A sheet with two or more years on other lines, a table of years by months, is not read this way.

A cell such as `362(37%)`, `1,234` written as text, or `15건` gives its leading number with a warning. `78.5%` written as text, `-` and `해당없음` give none. Commas count as thousands only every three digits: `1,2,3` and `12,34` are not one number, so they are cells it could not read, and it asks. Spaces that group thousands, as a statistics yearbook writes `39 925`, are read the same way; `12 34` and `2019 2020` are still not one number. So it does for a cell between the first and last date that holds digits but no date it can read, such as `2026_`: its row falls in no period.

When summing:

- **A date on the first row of a group** covers the rows below it, until the next date or a row such as `Total`. When `서울`, `부산` and `소계` sit under a merged date, the automatic choice uses only the `소계` line: adding all three would count everything twice. By item, the `소계` line is left out and `서울` is set against `부산`.
- **The same date on two sheets** is refused, because it would count twice.
- **A line headed differently on another sheet** is refused, because row 7 may be something else there.
- **Warnings:** the newest date has a different number of rows than usual, or its last time of day is over an hour earlier than usual. Times need the date cell to hold them.

For a table with no dates:

- **The whole** is the total the sheet states. Besides `합계` and `총계`, a line whose name ends in `전체`, `합계` or `총계`, such as `공원관리사업소 전체`, is a total line and is left out of the items, as are `전국` over regions and `전산업` over industries. Without one, the items are added and the sentence says "the items' sum", with a warning, because SlideFill cannot know that every column can be added.
- **A stated total that differs from the items' sum** by over 0.5% means rows overlap, so no shares are given.
- **Rates, indexes and means** get no shares. This is judged from headings and the sheet's name, so a column with no such word can still be added wrongly.
- **`소계`** rows are left out of the items, and a catch-all such as `기타` is ranked last. Closing lines merged in from the left column, such as `유료` and `무료`, count as subtotals too.
- **In a table with a kind on each row**, if the kind column has a line meaning all of them (`FY Subtotal`, `00:소계`, `5합계`), only that line is used. A line with another name counts too if its value equals the other lines added up, date after date (within 2%, at least 90% of the time). The line chosen is written beside the figure's name. Adding the parts to the whole would double the value. When kinds split each name (`월계` and `누계`, `건수` and `연면적`), only the first is used. A column headed `지표` or `Measures` holds different units, so it is split in tables by date as well.
- **Tables stacked on one sheet** (the same heading and names coming back further down, or a table with headings of its own starting under a blank row): only the first is read, and it asks before making the report. **Lines indented as in a pivot table** are parts of the line above, so only the outermost level is compared. So it is when the levels stand in columns of their own, with one name to a line (시·도, then its 시·군·구 in the next column, then their 행정동): adding every level counts each hectare two or three times. A group's name written beside the first of its lines (서울 beside 강남) names the group, and that line is one of the inner ones.
- **A table beyond an empty column**: when a column of names of its own stands beyond the gap, its figures belong to that table (`수량` beside `제품` is not a figure of `지점`). An empty column with no names beyond it is read as one table.
- **A row equal to the sum of the rows under it** takes them as its parts and leaves them out of the comparison. When the sheet does not mark them as parts, by indenting, a leading space or bold, the sum may match by chance (300 = 100 + 200), so it asks first. When their names say whose parts they are, `611-01` under `611` or `서울 강남` under `서울`, it does not ask. When it asks, it names the row and the rows it seemed to add up.
- **An average over time for an amount**, such as `분기별 평균 수출금액`, adds up across items. `평균 단가` and `평균 점수` are averaged.
- **Identifiers and codes** are not chosen as figures: columns whose name says so (`ABN`, `ISBN`, `FILE NUMBER`, postcodes, latitude), columns that repeat one to one with the words beside them (a species code, an area repeated for each region), and whole numbers of nine or more digits that are all different. Neither are places in an order (`발생차수`, rank, a running number that starts again), dates stored as numbers (`Group End Date`), columns where every line is a year, or lists that mix 0.003 with 6,378,137 in one column. Nor is a column where every item has the same value (a fee of 0 for each member): there is no highest or lowest to tell, and the total line under them may be all it holds. Ages are not added; people are counted instead. If no figure is left, rows are counted or the file is said to be unreadable. This is judged from names and contents, so it can miss one or drop the wrong one, and the window says which figure was taken.
- **A column saying when each item came to be** (`설치년도`, `준공`, `건립`, `설립`, `취득`, `출생`, `지정일자`, installed, built, established, designated) is not used as the dates: the list is read as a table of named items. Read by the year each street light was put up, a ledger becomes a series that "fell every year".
- **A heading that repeats** takes the merged heading right above it (`일반수출신고 건수`), so a figure for one part does not read as the whole.
- **A heading merged down over the first line of figures**, such as `구 분` over A3:C5 with the sheet's total on row 5, is not that line's name, so the total is not ranked as one more item.
- **Lines sharing one merged name**: when several lines with figures of their own sit under one name merged down over them, such as switchboards under a road, each is compared under that name, and one switchboard's value can read as the whole road. It asks before making the report, naming the road and how many lines share it. When the lines are added up under the road and the sheet's own `소계` lines agree with that sum, it does not ask: no one line stands for the road then.
- **A ledger named only by merged names**: when the column that should name each line holds nothing but `소계` and `총계`, as a street-light ledger with no switchboard names does, it starts no table of its own, and the road merged down on the left names the lines. If the sheet's own `소계` lines equal the lines above them (two or more, and none differs), the column is added up by road: 흥안대로 404, as the sheet says. Without that, lines of named things are counted, as before, since the column could be an age or a price.
- **A data dictionary** (`Field Name`, `Field Type`, `Field Width`, or `컬럼명` and `데이터 타입`) describes another table's columns, so it is not read as figures.
- **A count of 1 on every line** (`Count` beside each record) is something to add, not a level, so it is not asked about.
- **People added across kinds that can overlap**: a number of people (`인원`, `환자`, `persons`) added across a column of kinds one person can be in more than once, such as inpatient, outpatient and pharmacy, may count the same person twice, so it asks first. Kinds that split people apart, such as sex, age, region or year, are added without asking.
- **A comparison of items under a period condition** says the period first: "2022년 5월 기준, ..." or "In May 2022, ...", and `2025-26 기준` for a financial year.
- **One name on several lines** is told apart by the column beside it that differs between them: a rate table's `일반용` becomes `일반용 2,001이상`.
- **`수입 계` and `지출 계`** (a `계` set apart, in four words or fewer) are total lines; `서울 소계` is a subtotal. The words before the `계` must hold two letters or more, so `회  계` and `통  계`, spaced out to fill their cells, are still words, not totals. With several total lines, the items are not checked against a total of another kind.
- **A heading row of bare years** (`항목 | 2024 | 2025`, the years written as numbers) heads the columns, so the figure is called `2025`, not `C`, and the change between the two years is stated.
- **Mean, maximum and minimum under each month**, on lines with empty date cells: those lines belong to the month above them, and the mean line is chosen rather than all three added or averaged, also when it is spaced out to fill its cell (`평 균`).
- **An `.xls` that is really an HTML table** has its `rowspan` and `colspan` laid out as Excel does, so the columns under a wide heading stay in place.
- **Two periods** name the largest rise in percent, the largest in amount if different, and the largest fall. Items under 1% of the largest are not named for moving most. Rates move in points.

- **Actuals beside a budget**: when the chosen figure is what happened (`2024-25 Actual`) and another column is a budget, target or plan for the same year, the reading sets the actual against that budget, in words that say so: "예산보다 가장 크게 웃돈 항목", not "가장 크게 늘어난 항목". With several budgets it takes the one a `Variation to published budget` column names, then a published or original one. A column headed 계획 or 목표 (plan, target) is called that, not a budget. A budget for another year (2026 예산 beside 2025 실적) is not set against the actuals.
- **Several total lines of different things**, as an operating statement has for income and expenses (`Total income from transactions`, `Total expenses from transactions`, also after a heading such as `Administered Income: Total administered income`): the lines between them are not ranked against one another, no top-five table or chart is made, and the sentences state the total lines, against the budget when there is one ("Total income from transactions 실적 33,037, 예산 28,120 대비 +17.5%"). Group totals of one kind are not this: when one total line is the sum of the others (서울 합계 30 and 부산 합계 12 under 총계 42), the items are compared as usual.
- **A whole left on a nameless line**: when a figure is the same on every item, so it is not chosen, and a line with no name right under them holds something else (each member's `Total expenditure $` of 0, the committee's $1,910.43 below), it asks before making the report, giving that value.
- **Level numbers** (`I.`, `1.`, `1-1.`, `(1)`), in two or more kinds, mean only the top level is compared and no shares are given, because adding a parent to its parts means nothing.
- **A negative base** gives no percentage, so that item is stated apart with its amounts.

A series that looks like rates, scores or change rates speaks in points rather than percent. This is judged from the heading, the sheet's name, words such as `percentage change` above the table, and the cell's format. A column headed `증감` alone changes by an amount.

## What stays on the slide

Warnings appear only in the app's list. So where a figure must be read differently, `note` and `reading` say so in the sentence.

- **A period with fewer rows than usual**: instead of comparing totals, the sentence says it looks incomplete and gives a per-row mean. That does not allow for the season of the missing rows, so with a month column, a monthly series through `A+B` is more accurate.
- **Newest values that are words** (`(결측)`, `집계중`): it says how many periods were passed over and which value was used.
- **A period not over when the file was saved**: stated first, and not compared with other periods. The reading comes from the periods that have ended, since three months set against whole years reads as a collapse however well warned. If the file gives a date such as `Data as at 30-Sep-2025` and it falls within that period, that date is used instead of the day it was saved.
- **A day cut short, a 0 filled in ahead**: stated first. `학년도` and `회계연도` are not calendar years, so they are not called unfinished.
- **Financial years** (`2025-26`) are read as the ending year only for sorting. A sentence says them as the sheet does. Financial quarters such as `Q1 12/13`, `Q1 FY24` and `FY24 Q1` are sorted as if the year starts in July, and written as the sheet writes them. When the year ends is read from the workbook's own words, at the top and bottom of every sheet: `year ending 30 June`, `FY ended March 31`, a whole year of months such as `July to June` or `1 April 2024 to 31 March 2025`, `4월~3월`, `7월 1일부터 6월 30일`, `3월 결산`. That end decides whether the newest year is over, and where a financial quarter falls on the calendar (`Q1 FY26` is October to December 2025 in a year that runs October to September). December says nothing about a `2025-26` year, so it is not taken. With no such words, or two that disagree, June is assumed, and calling a period unfinished can then be a few months off.
- **Different kinds in one column**: if the value's name says nothing, such as `값` or `value`, and a column of kinds such as 항목 or 구분 sits beside it, only the first kind is used. A named figure whose kinds differ in size by over 100 times is asked about.
- **A number column beside a year**: it reads as months only if its heading says 월, 기간 or month, and as quarters if it says 분기 or quarter or holds `Q4` or `4분기`. A heading naming something else, such as a branch number, is not part of a date.
- **Numbers that look like years**: a column of numbers such as 1987 reads as years only when the heading right above them (not a sheet title further up) is a word such as 연도, year or period, when there is no heading, or when they run a year at a time or step evenly by 2, 3, 4, 5 or 10 years as a census does. 1987 to 2036 under `건수` are counts. A date written as text or digits is read only for the years 1800 to 2199, so `1001.5` is not a date. `[2023]`, as a catalogue writes a year it worked out, is 2023.
- **Year-like numbers along a row** follow the same rule as a column: they read as years only after words such as year or 구분 at the start of the row, or when they run a year at a time. 2,026, 1,957 and 2,072 on a line of a statement are amounts.
- **Digits kept as text with a leading zero** (`02`, `0715`) are codes for a district or an office and are not chosen as figures. So is a column of digits all one width with leading zeros, such as industry codes `01` to `99`.
- **A period with most of its values missing**: if a past period that a comparison uses has under half the cells its line usually has (3 months of 12), it asks before making the report, since a change from missing months means nothing.
- **Over 100 in a per-cent column**: a line over 100 in a column whose heading or format says per cent, such as `Population estimates (n)`, may be a count of people among the shares, so it asks.
- **Last year and this year**: with `전년 실적` beside `당년 실적`, this year's is chosen. `전년 대비` and `전년비` are changes up to now, so they stay.
- **Blank cells with only a format** past the last value are not read, such as the one a format given to a whole column leaves on row 1,048,576. A merged cell keeps the area it is drawn over.
- **`률` and `율`**: only at the end of a word do they mean a rate. `법률상담건수` is not one.
- **Prices per item**: a heading ending in `수수료`, `요금`, `가격`, `단가`, `price` or `fee` is ranked, not added. `수수료 수입` and `총수수료` are money taken, so they add up. Writing `합계( )` adds them anyway.
- **Period columns it cannot read**: with four or more columns differing only in their digits, such as `H1 21/22`, choosing the first asks before making the report, since it may not be the newest.
- **Values that cannot be added**: a name meaning a level, such as a heating value, a concentration or a temperature, is averaged rather than added. A column of kinds, such as 항목 or type, narrows it to one kind first.
- **A sheet in per cent**: if it says so, as `Per cent` under `Unit:` or `(단위: %)`, values are written in per cent, changes in points, and no shares are given. A unit note covers the table under it, or, with no figures under it, is a footnote to what is above. A % merely somewhere on the sheet is asked about.
- **A figure that stops early**: if the first choice has no value for the sheet's newest period, the largest total that does is used. The newest period is the latest where most actual figures are filled in. A figure named as a target, plan or budget is not chosen while an actual one exists, and months with only a target written ahead do not count as the newest.

## Reading rules

`{{reading:...}}` decides how to read a figure with fixed rules. The usual level of a weekday is the median of its previous 8 values, so one holiday does not move it.

**Daily data:**

| Rule | When it speaks |
|---|---|
| Outside the usual range | At least 3 usual spreads and 7% from the weekday's median |
| Large, within the range | At least 15% from the median, on a weekday that swings widely. The rules state the move, and only ask to check the last period if the value is also the lowest of those weekdays |
| Last week was odd | Last week's same day was itself far from its usual level, so the usual level is the fair base |
| A weekend in the mean | The same-weekday and recent-mean comparisons point different ways, by at least 3%, and weekends differ from weekdays by at least 10% |
| Weekday against weekend | The previous value is across the weekend line |
| One day moved the mean | A recent day stands out, and leaving it out moves the comparison by at least 2 points |
| Weeks in a row | The 7-value total has moved the same way for 3 weeks |
| Highest or lowest | Of the newest 30 |
| Ordinary | None of the above, and within 5% of the median |
| Nothing stands out | None of the above, under 15% from the median or the previous value, with at least 14 values. A day with no finding still gets a sentence, so no placeholder is left in the report |

**Monthly and quarterly data** read by the same period last year first, because that leaves the season out:

| Rule | When it speaks |
|---|---|
| Same period last year | Whenever there is one |
| The comparisons disagree | Against the previous period and against last year point different ways, each by at least 3%. Read it by last year |
| No year before | There is no such value, so the season cannot be separated out. States the previous period and the rank |
| The period before was extreme | It was the highest or lowest of the latest 12, so the change against it reads large or small |
| In a row against last year | At least 3 periods in a row above or below |
| Same period over the years | With 3 earlier years: the highest, the lowest, or under 70% of the best |
| A running total | It only rises within each year, and each January falls to under half of December. Only last year is a fair base |

A fall of at least 15% against the previous period also asks to check that the last period is fully counted.

**Yearly data** states the year before, years in a row, and the highest or lowest of at least 5. Part of a year states the change per month instead.

With 3 to 13 values and nothing else to say, a sentence states the change and says the usual level cannot be judged. With fewer than 3, the placeholder stays.

The thresholds for daily data were set by running the rules on every day of 2025 in two public datasets, then checked on data that did not set them: three station sheets of the same year, and 2022 to 2024 of the other dataset. Firing rates were close. On the 8th-line ridership, 13 days were named as standing out, and 11 of them were public holidays by a list written by hand, plus May 1 (Labor Day) and one Sunday. On the sewage inflow, which swings with events the data does not show, the "one day moved the mean" rule speaks on about a third of days, so rules are noisier on series like that. Two datasets is not many, so expect to adjust these.

The monthly, yearly and large-move rules came from a case study of 17 public workbooks, so those workbooks cannot prove them. They were then read on 10 series the changes had not looked at, other columns of the same files: 8 read correctly, one named only one of its spikes, and one was a running total read wrongly, which is why that rule exists. On widely swinging daily series, such as online sales, the large-move sentence appears on about half of days and "one day moved the mean" on over half, so those rules are noisy there.

## Insight candidates

Three stages, with a person at the end. `core.run` never calls a model, and does not even import the module that can. So the preview and the report always agree, and a report never changes because a model answered differently.

1. **Code calculates.** `facts.py` works out the figures. It also finds what is worth noticing: comparisons that point different ways, a value outside the usual range for its weekday, the day least like its usual level.
2. **The router picks a model.** It counts Hangul and Latin letters in the template, and gives each sheet name and each cell of words in the top corner of a sheet one vote, so long English names do not outvote short Korean headings. A rule does this rather than another model, because it needs no memory and has been right so far.
3. **The model writes, and checks filter.** The model cites the facts it used, from a list fixed by the output format. A sentence is left out if it holds a number its cited facts do not, makes a comparison they do not (a change against last week written as "a year earlier"), gives the sum or another item's value as one item's, names the figure from a sheet's title or by a guess ("households"), gives a change or a low in a period that had not ended without saying so, judges ("highly concentrated"), says "highest" or "most" when no cited fact does (one line's +10.3% called the highest against budget), states a cause, leaks an internal name, is in the wrong language, or makes a known misreading. A possible cause goes in a separate field, is shown as something to check, and never enters the report.

The person's choice is saved with a fingerprint of the calculated results. When the figures change, the fingerprint no longer matches, and the report goes back to the calculated sentence and says so.

### Why these models

Tested on real public data with 8GB of video memory as the limit, reading every output (18 rounds each, about 50 sentences each). The error counts are of clear errors in the Korean sentences, counted by hand, so treat them as rough.

| Model | Memory | Mean time | Clear errors in Korean |
|---|---|---|---|
| A.X 4.0 Light | 5.0GB | 6.6s | about 6 of 34 |
| Qwen 3.5 4B | 3.1GB | 8.7s | about 9 of 36 |
| EXAONE 3.5 7.8B | 5.7GB | 9.5s | about 14 of 33 |
| Kanana 1.5 8B | 5.8GB | 12.5s | about 12 of 24 |

A.X 4.0 Light is used for Korean. Qwen 3.5 4B is used for English, where its sentences were the most accurate. Passing the automatic checks did not predict quality: vague sentences with few numbers pass easily. That is why a person chooses.

## Tables

Rows are added by copying an existing body row, so new rows look like the others. If the table has PowerPoint's "Total Row" turned on, the last row stays last. If cells are merged down only in the heading rows, rows come and go at the bottom and the merges stay. A table merged down in its body cannot change its row count, because copying part of a merge would break it. A table that grows as far as a shape that was below it gets a warning.

## Charts

SlideFill replaces the chart's data and its embedded workbook, which is what PowerPoint's "Edit Data" opens. The chart's own formatting is not touched.

## Several files

With two or more files, each is read on its own and gets its own slides, and one overview slide goes first.

The overview's first sentence comes from the code, not the model. It takes each file's lead figure and its change from a year earlier (or from the period before), and says whether they all rose, all fell, or moved in different directions. If the files end at different times, or compare against different periods, it says so and does not compare them.

Whether one file affects another cannot be calculated, so it is never said. In a test with the real model, leaving the comparison to the model gave about 3 sentences out of 24 that truly compared the files, and wrong readings passed the checks. So the code compares, the model only reads the result and suggests sentences, and only a sentence someone chooses goes in the report.
