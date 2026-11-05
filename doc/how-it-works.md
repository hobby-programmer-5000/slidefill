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

Dates are also read as `2026 Aug`, `Aug. 2026`, `Aug-26`, `26년 8월`, `2026 Q3`, `3Q26`, `01월` under a merged `2024년`, and `January` under a heading row `Year 2026`. A month with no year anywhere is not guessed.

A cell such as `362(37%)`, `1,234` written as text, or `15건` gives its leading number with a warning. `78.5%` written as text, `-` and `해당없음` give none.

When summing:

- **A date on the first row of a group** covers the rows below it, until the next date or a row such as `Total`.
- **The same date on two sheets** is refused, because it would count twice.
- **A line headed differently on another sheet** is refused, because row 7 may be something else there.
- **Warnings:** the newest date has a different number of rows than usual, or its last time of day is over an hour earlier than usual. Times need the date cell to hold them.

For a table with no dates:

- **The whole** is the total the sheet states. Without one, the items are added and the sentence says "the items' sum", with a warning, because SlideFill cannot know that every column can be added.
- **A stated total that differs from the items' sum** by over 0.5% means rows overlap, so no shares are given.
- **Rates, indexes and means** get no shares. This is judged from headings and the sheet's name, so a column with no such word can still be added wrongly.
- **`소계`** rows are left out of the items, and a catch-all such as `기타` is ranked last. Closing lines merged in from the left column, such as `유료` and `무료`, count as subtotals too.
- **In a table with a kind on each row**, if the kind column has a line meaning all of them (`FY Subtotal`, `00:소계`, `5합계`), only that line is used. A line with another name counts too if its value equals the other lines added up, date after date (within 2%, at least 90% of the time). The line chosen is written beside the figure's name. Adding the parts to the whole would double the value. When kinds split each name (`월계` and `누계`, `건수` and `연면적`), only the first is used. A column headed `지표` or `Measures` holds different units, so it is split in tables by date as well.
- **Tables stacked on one sheet** (the same heading and names coming back further down): only the first is read. **Lines indented as in a pivot table** are parts of the line above, so only the outermost level is compared.
- **An average over time for an amount**, such as `분기별 평균 수출금액`, adds up across items. `평균 단가` and `평균 점수` are averaged.
- **Identifiers and codes** are not chosen as figures: columns whose name says so (`ABN`, `ISBN`, `FILE NUMBER`, postcodes, latitude), columns that repeat one to one with the words beside them (a species code, an area repeated for each region), and whole numbers of nine or more digits that are all different. Neither are places in an order (`발생차수`, rank, a running number that starts again), dates stored as numbers (`Group End Date`), columns where every line is a year, or lists that mix 0.003 with 6,378,137 in one column. Ages are not added; people are counted instead. If no figure is left, rows are counted or the file is said to be unreadable. This is judged from names and contents, so it can miss one or drop the wrong one, and the window says which figure was taken.
- **A heading that repeats** takes the merged heading right above it (`일반수출신고 건수`), so a figure for one part does not read as the whole.
- **Two periods** name the largest rise in percent, the largest in amount if different, and the largest fall. Items under 1% of the largest are not named for moving most. Rates move in points.

- **Level numbers** (`I.`, `1.`, `1-1.`, `(1)`), in two or more kinds, mean only the top level is compared and no shares are given, because adding a parent to its parts means nothing.
- **A negative base** gives no percentage, so that item is stated apart with its amounts.

A series that looks like rates, scores or change rates speaks in points rather than percent. This is judged from the heading, the sheet's name, words such as `percentage change` above the table, and the cell's format. A column headed `증감` alone changes by an amount.

## What stays on the slide

Warnings appear only in the app's list. So where a figure must be read differently, `note` and `reading` say so in the sentence.

- **A period with fewer rows than usual**: instead of comparing totals, the sentence says it looks incomplete and gives a per-row mean. That does not allow for the season of the missing rows, so with a month column, a monthly series through `A+B` is more accurate.
- **Newest values that are words** (`(결측)`, `집계중`): it says how many periods were passed over and which value was used.
- **A period not over when the file was saved**: stated first, and not compared with other periods. The reading comes from the periods that have ended, since three months set against whole years reads as a collapse however well warned. If the file gives a date such as `Data as at 30-Sep-2025` and it falls within that period, that date is used instead of the day it was saved.
- **A day cut short, a 0 filled in ahead**: stated first. `학년도` and `회계연도` are not calendar years, so they are not called unfinished.
- **Financial years** (`2025-26`) are read as the ending year only for sorting. A sentence says them as the sheet does. Financial quarters such as `Q1 12/13`, `Q1 FY24` and `FY24 Q1` are taken to start in July for sorting, and written as the sheet writes them. The file does not say when its year starts, so calling a period unfinished can be a few months off.
- **Different kinds in one column**: if the value's name says nothing, such as `값` or `value`, and a column of kinds such as 항목 or 구분 sits beside it, only the first kind is used. A named figure whose kinds differ in size by over 100 times is asked about.
- **A number column beside a year**: it reads as months only if its heading says 월, 기간 or month, and as quarters if it says 분기 or quarter or holds `Q4` or `4분기`. A heading naming something else, such as a branch number, is not part of a date.
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
2. **The router picks a model.** It counts Hangul and Latin letters in the template and the sheets. A rule does this rather than another model, because it needs no memory and has been right so far.
3. **The model writes, and checks filter.** The model cites the facts it used, from a list fixed by the output format. A sentence is left out if it holds a number its cited facts do not, makes a comparison they do not (a change against last week written as "a year earlier"), gives the sum or another item's value as one item's, names the figure from a sheet's title or by a guess ("households"), gives a change or a low in a period that had not ended without saying so, judges ("highly concentrated"), states a cause, leaks an internal name, is in the wrong language, or makes a known misreading. A possible cause goes in a separate field, is shown as something to check, and never enters the report.

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
