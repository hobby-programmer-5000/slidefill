"""All UI strings, English and Korean."""

STRINGS = {
    "en": {
        "window_title": "SlideFill",
        "tagline": "Choose an Excel workbook and get a PowerPoint report.",
        "own_template": "Use my own PowerPoint template (optional)",
        "add_files": "Add files…",
        "remove_files": "Remove",
        "drop_hint": "Drag Excel and Word files here, or press Add files.  .xlsx .xls .xlsb .ods .csv .tsv .docx",
        "err_doc_needs_workbook": "A Word document is read alongside figures, so add an Excel file too.",
        "err_wrong_type_document": "This kind of document cannot be read. Save it as .docx.\n{path}",
        "drop_refused": "Not a kind of file SlideFill reads: {names}",
        "report_many": "{first} and {more} more",
        "report_many_file": "{first}_and_{more}_more",
        "auto_ready_unread": "Read the workbooks and chose the contents: {what}. Nothing could be read from: {names}. Check the values, then create the report.",
        "auto_ready_doubt": "Read the workbook and chose the contents: {what}. {count} thing(s) need checking; you will be asked before the report is made. To show something else, press Ctrl+R.",
        "doubt_title": "Please check this figure",
        "doubt_intro": "SlideFill chose the figures by itself, and some signs say it may have chosen wrongly. No report has been made yet.",
        "doubt_report_note": "※ Not confirmed. {reason}",
        "doubt_chosen": "Chosen: {what}",
        "doubt_change": "Choose the figures…",
        "doubt_keep": "Create it as it is",
        "doubt_cancel": "Cancel",
        "doubt_example_sheet": "This sheet's name says it is an example or a template, so it may not hold real data.",
        "doubt_count_only": "Nothing here looked safe to add up, so it counts rows. A count may not be what this file is for.",
        "doubt_sum_of_levels": "{figure} is added up across rows, but it looks like a level, such as a heating value or a concentration. Adding those up means nothing.",
        "doubt_first_table": "This sheet seems to hold several tables, and only the first was read.",
        "doubt_heading_below": "A heading appears below the first figure in this column, so the name {figure} may belong to another table.",
        "doubt_dates_sparse": "Many figures have no date beside them, so they were left out of the totals by date.",
        "doubt_maybe_percent": "A note on the sheet says per cent, so {figure} may be a percentage, not a count.",
        "doubt_several_tables": "This sheet seems to hold several tables one under another, read here as one list.",
        "doubt_total_differs": "The total written in the sheet differs from the items added up, so some rows may overlap.",
        "doubt_period_columns": "{figure} looks like the first of several period columns, so it may not be the newest.",
        "doubt_mixed_kinds": "{figure} adds up different kinds of row, and their sizes are far apart. They may be different measures, which mean nothing added together.",
        "doubt_figures_not_read": "Some cells hold digits that could not be read as figures, so they were left out of what is shown.",
        "doubt_count_among_shares": "{figure} is a column of per cents, but {item} is over 100. It may be a count of people or cases ranked among the shares.",
        "doubt_sparse_period": "Only {filled} of the usual {usual} cells are filled for {period}. A comparison with it may be with a period whose values are missing.",
        "doubt_merged_names": "{count} lines with figures of their own share the merged name {item}, and each was compared under that name. One line's value may read as the whole of {item}.",
        "doubt_people_across_kinds": "{figure} was added up across {kind} ({values}). One person can be in more than one of them, so the sum may count people twice and is not a number of people.",
        "doubt_nameless_total": "{item} is {same} for every line, so it was not chosen. A line with no name under them holds {total}, which may be the whole of what the sheet lists.",
        "doubt_parts_by_sum": "In {figure}, {parent} equals the sum of the rows under it ({parts}), so those rows were taken as its parts and left out of the comparison. If the sum matches by chance, they are missing.",
        "template_first_file": "A template reads one workbook, so only {name} is used. Untick the template to report on all the files.",
        "insight_about": "About",
        "insight_all_files": "All the files together ({files})",
        "contents_figures": "What goes in",
        "contents_preview": "The slide it makes ({n} of {total})",
        "contents_preview_none": "Nothing chosen",
        "contents_detail": "Change how it is calculated…",
        "contents_count": "{slides} slide(s), {figures} figure(s)",
        "contents_change": "Now: {before}.  If you use this: {after}.",
        "contents_same": "This is the report as it is now: {now}. Tick or untick to change it.",
        "contents_none": "Tick at least one figure.",
        "contents_limit": "A sheet can show up to {count} figures on its slide.",
        "contents_skipped": "{count} large sheet(s) are not listed, so this window stays quick.",
        "auto_ready": "Read the workbook and chose the contents: {what}. Check the values, then create the report. To show something else, press Ctrl+R.",
        "err_auto_nothing": "No dates or named figures could be read from this workbook, so SlideFill cannot lay out a report by itself.\n\nTick \"Use my own PowerPoint template\" to say what goes where. Press F1 to see how.",
        "helper_button_auto": "Change contents (Ctrl+R)",
        "helper_title_auto": "Change what the report shows",
        "auto_ready_left": "Read the workbook and chose the contents: {what}. Left out: {figures} more figure(s) and {sheets} more sheet(s). Check the values, then create the report. To show something else, press Ctrl+R.",
        "helper_intro_auto": "Tick the figures on the left. The right side shows the slide the report will really have, and changes as you tick. Nothing changes in the report until you press Use these contents.",
        "helper_use": "Use these contents",
        "helper_starter_auto": "Save as a template I can edit",
        "template": "Template (.pptx)",
        "excel": "Files",
        "output": "Save to folder",
        "browse": "Browse…",
        "preview": "Preview (F5)",
        "generate": "Generate report (Ctrl+G)",
        "open_result": "Open report",
        "open_folder": "Open folder",
        "help": "Help (F1)",
        "language": "한국어",
        "pick_template": "Choose the PowerPoint template",
        "pick_excel": "Choose the Excel workbook",
        "pick_output": "Choose where to save reports",
        "filetype_pptx": "PowerPoint presentation",
        "filetype_excel": "Excel, data and Word files",
        "col_status": "Status",
        "col_item": "Item",
        "col_value": "Value",
        "col_source": "From",
        "col_where": "Where",
        "status_ok": "OK",
        "status_warn": "Check",
        "status_error": "Problem",
        "kind_text": "Text",
        "kind_table": "Table",
        "kind_chart": "Chart",
        "kind_unsupported": "Other",
        "where_slide": "Slide {n}",
        "where_notes": "Notes {n}",
        "where_master": "Master {n}",
        "where_layout": "Layout: {n}",
        "where_part": "{n}",
        "start_hint": "Choose a workbook to see what the report will show.",
        "working": "Reading files…",
        "generating": "Creating the report…",
        "summary": "Ready {ok} · Check {warn} · Problems {error}",
        "no_placeholders": "No {{placeholders}} found in this template. Press F1 to see how to add them.",
        "saved": "Saved: {path}",
        "select_hint": "Select a row to see details.",
        "confirm_title": "Some items could not be filled",
        "confirm_body": "{count} item(s) have problems and will stay as {{{{placeholders}}}} in the report.\n\nCreate the report anyway?",
        "error_title": "SlideFill",
        "missing_inputs": "Choose a workbook first.",
        "help_title": "How SlideFill works",
        "help_body": """Choosing a workbook is enough. SlideFill reads it, chooses the contents, and shows every value before you create the report. If it took the wrong figure, press Ctrl+R to change it.

The rest of this page is for your own company template. Tick "Use my own PowerPoint template" and mark in it what goes where.

1. In PowerPoint, type placeholders into your template.

Text (anywhere: text boxes, table cells, notes, chart titles, slide master)
    {{Sales}}            a named cell in Excel (Formulas > Define Name)
    {{Summary!B3}}       a sheet and cell
    {{'My Sheet'!B3}}    quotes are optional when the sheet name has spaces

Tables and charts: right-click the table or chart > View Alt Text, and type
    {{table:Summary!A1:D10}}     or  {{table:SalesTable}}
    {{chart:Summary!A1:D5}}      or  {{chart:TrendRange}}

Names can be a defined name or, for ranges, an Excel table name.

Tables fill from the top-left cell, so include the header row in the range. Rows are added or removed to fit. The column count must match.

Charts read the first row as series names and the first column as categories.

Data that grows every day: name the date column first, and SlideFill finds the newest rows by date. Subtotal rows are skipped, and so are future dates with nothing in them yet.
    {{latest:Daily!A,C}}              column C on the newest date in column A
    {{latest:Daily!A}}                the newest date itself
    {{table:latest7:Daily!A,C,D}}     the newest 7 rows, oldest first
    {{chart:latest7:Daily!A,C,D}}     the same rows as a chart
A table keeps the header row your template already has. A blank one, and chart series names, come from the sheet.

Figures SlideFill works out itself, from the date column and one value column:
    {{week:Daily!A,C}}     change from the same weekday last week
    {{prev:Daily!A,C}}     change from the previous value
    {{avg7:Daily!A,C}}     change from the mean of the previous 7 values
    {{peak30:Daily!A,C}}   the date of the highest of the newest 30 ({{low30:...}} for the lowest)
    {{note:Daily!A,C}}     a report sentence with these figures
    {{reading:Daily!A,C}}  how to read the figure, by fixed rules: a weekend inside a mean, a holiday-like day, a value outside its usual range, the same month last year
Monthly, quarterly and yearly data are recognized from the dates. A year alone, such as 2021, works as a date in a column of years.

You do not have to write these by hand. Press Ctrl+R: SlideFill reads the workbook, shows what each one gives, and copies it or saves a starter template.

How the data is laid out, in any of the references above:
    latest:Daily!A,sum(D:E)        add up the rows of each date. Rows are only added when you write sum
    latest:Daily!A,sum(D, B=East)  only rows whose B contains East
    latest:Daily!A,count(C)        how many rows each date has
    latest:Daily!A,sum(D*E)        quantity times price
    latest:Daily!A,mean(D)         the mean of each date's rows
    latest:Daily!A,rate(B=Yes)     the share of each date's rows whose B contains Yes
    latest:Monthly!A+B,C           the year in A, the month in B
    latest:Monthly!5,7             periods across row 5, values in row 7
    latest:Monthly!C2:N2,C3:N3     a limited range, to keep two tables on one sheet apart
    latest:Year*!A,C               one series across every sheet named Year...

A table with no dates, where the first column names the items:
    {{reading:Table!A,B}}    the largest and its share, or with A,B,C what moved most between two periods
    {{top1:Table!A,B}}       the name of the largest ({{topvalue1:...}} its value, {{share1:...}} its share)
    {{table:top5:Table!A,B}} the largest five ({{table:share5:...}} adds shares; {{chart:top5:...}} works too)
If there is nothing to compare with, the placeholder stays and is listed as a problem.

A sentence you choose, with help from a model on this PC:
    {{insight:Daily!A,C}}    press Ctrl+I, read the candidates, and choose
Until you choose, the report uses the calculated sentence. When the figures change, you choose again. This needs Ollama and its models. Without them, everything else still works.

2. Choose the template and workbook here. The list shows every value before anything is created.

3. Press Generate. Your template and workbook are never changed. Each report is a new file.

Good to know
- Values are the ones Excel last saved. Save the workbook first.
- Numbers appear as Excel displays them: 85.3%, 1,234,567, 2026-11-03.
- Anything that cannot be filled stays as {{placeholder}} so you can see it. Blank means the Excel cell is blank.
- Password-protected or DRM-protected files cannot be read.

Shortcuts
    F5  Preview          Ctrl+G  Generate
    Ctrl+T  Template     Ctrl+E  Workbook
    Ctrl+L  Language     F1  Help""",
        "err_file_missing": "File not found:\n{path}",
        "err_file_locked": "Windows would not let SlideFill open this file. Close it in other programs and try again.\n{path}",
        "err_file_unreadable": "This file could not be read. It may be damaged.\n{path}\n\n{detail}",
        "err_file_protected": "This file is protected by a password or by document security (DRM), or it is an old-format file. Remove the protection or save it as a normal Office file, then try again.\n{path}",
        "err_file_not_office": "This is not a normal Office file. It may be wrapped by document security (DRM).\n{path}",
        "err_wrong_type_excel": "This kind of file cannot be read. SlideFill reads .xlsx .xlsm .xls .xlsb .ods .csv and .tsv files.\n{path}",
        "err_wrong_type_template": "The template must be a .pptx file. In PowerPoint, use Save As to convert it.\n{path}",
        "err_cannot_save": "The report could not be saved. Check the folder.\n{path}\n\n{detail}",
        "err_unexpected": "Something unexpected went wrong.\n\n{detail}",
        "issue_missing_name": "No defined name or Excel table called \"{name}\".",
        "issue_missing_sheet": "No sheet called \"{sheet}\".",
        "issue_bad_ref": "\"{ref}\" is not a valid cell address.",
        "issue_range_for_text": "{source} is several cells. Text needs a single cell. For a table, use alt text.",
        "issue_name_not_cells": "The name \"{name}\" is a formula or a constant, not cells. Point it at cells instead.",
        "issue_multi_area": "The name \"{name}\" points to several places. Use one range.",
        "issue_formula_no_cache": "{cell} has a formula with no saved result. Open the workbook in Excel and save it.",
        "issue_error_value": "{cell} shows the error {error} in Excel.",
        "issue_table_error_value": "{cell} shows the error {error} in Excel. It was copied as it is.",
        "issue_empty_cell": "{cell} is empty, so this became blank.",
        "issue_format_fallback": "The number format of {cell} ({format}) is not supported, so the plain value was used.",
        "issue_marker_in_text": "Table and chart markers go in the alt text of a table or chart, not in text.",
        "issue_marker_wrong_shape": "This {kind} marker is on \"{shape}\", which is not a {kind}.",
        "issue_column_mismatch": "The range has {range} columns but the table has {table}.",
        "issue_merged_resize": "Cells merged down in this table's body, or a heading with no body row below it, mean rows cannot be added or removed. Match the row count.",
        "issue_figures_not_read": "{count} cell(s) with digits could not be read as figures and were left out, such as \"{text}\".",
        "issue_table_covers": "At {rows} rows the table reaches \"{name}\" below it and will likely cover it. Leave more room under the table in the template.",
        "issue_latest_partial_rows": "{count} of the rows shown have some columns empty, such as {first}. If those periods have no figures yet, they show as blanks.",
        "issue_cross_backwards": "\"{second}\" looks older than \"{first}\", but changes are worked out from the first column to the second, so a rise would read as a fall. Swap the two columns in the reference.",
        "issue_merged_value_lost": "\"{value}\" falls on a merged cell and was not shown.",
        "issue_table_overflow": "With {rows} rows and its wrapped text, the table will likely run past the bottom of the slide.",
        "issue_chart_too_small": "A chart range needs at least 2 rows and 2 columns.",
        "issue_chart_unsupported": "Bubble charts are not supported.",
        "issue_chart_not_number": "\"{value}\" is not a number, so the chart leaves a gap there.",
        "issue_chart_header_number": "\"{value}\" in the first row became a series name, so that row is not plotted. If it is data, include the header row in the range.",
        "issue_unsupported_location": "Slide {slide} has a placeholder in a part SlideFill cannot read (SmartArt saved in an unusual shape, or the like), so it stays as written. Move it to a text box.",
        "issue_latest_bad": "\"{key}\" is not a valid reference. Write it like latest:Daily!A,C, latest7:Daily!A,C,D or latest:Monthly!5,7 (rows), with the dates first, or sum(D:E) and count(A) for the values.",
        "issue_latest_no_dates": "Column {column} of sheet \"{sheet}\" has no dated row with a value in it.",
        "issue_latest_ambiguous": "{date} has {count} rows with values, so SlideFill cannot choose one. For a total per date, write sum(C). For one kind of row, sum(C, B=name). For two tables on one sheet, limit it to one: A10:A400,C10:C400.",
        "issue_latest_fewer": "Asked for {wanted} rows but only {found} have dates and values.",
        "issue_latest_older": "{newest} is not filled in here yet, so this uses {used}. Other values in the report may be from a newer day.",
        "issue_latest_duplicates": "{count} date(s), up to {date}, appear on more than one row, so these figures cannot be compared. For two tables on one sheet, limit it to one: A10:A400,C10:C400. For a total per date, write sum(C).",
        "issue_latest_sheets_overlap": "{date} appears on both \"{first}\" and \"{second}\", so adding them up would count it twice. Name the sheets so that they do not overlap.",
        "issue_latest_sheets_differ": "{line} is headed differently on \"{first}\" and \"{second}\", so it may be a different item on each sheet.",
        "issue_latest_zero": "The newest value, on {date}, is 0 after values that were not. If 0 was filled in ahead, as a formula on an empty period gives, this is not the latest figure.",
        "issue_latest_text_newer": "{count} newer period(s) say \"{text}\" instead of a figure, so this uses {used}.",
        "issue_cross_nested": "Some rows are the sum of the rows under them, so only the {count} rows that stand side by side are compared.",
        "issue_cross_outline": "The names are numbered by level, so only the {count} top-level items are compared, and no shares are given.",
        "issue_cross_several_totals": "The table has {count} total lines of different things ({names}), as a statement has for income and expenses, so its lines are not ranked against one another and only the total lines are stated.",
        "issue_latest_group_size": "{date} has {found} row(s) where each date usually has {usual}, so its total may be incomplete or include another row.",
        "issue_latest_partial_day": "The last entry on {date} is at {last}, while days usually run to {usual}, so that day may be incomplete.",
        "issue_number_from_text": "{count} cell(s) hold a number written as text, such as {cell}: \"{text}\". The leading number was used.",
        "issue_cross_no_items": "Column {column} of sheet \"{sheet}\" has no dates and fewer than two named rows with figures.",
        "issue_cross_total_differs": "The total stated in the sheet, {stated}, differs from the items' sum, {added}. The rows may overlap, so no shares are given.",
        "issue_cross_total_found": "\"{label}\" equals the sum of the other rows, so it was treated as their total, not an item.",
        "issue_cross_no_total": "The sheet states no total, so shares are of the items' sum, {added}. If these figures cannot be added, such as an index or a score, the shares mean nothing.",
        "issue_cross_unlabeled": "{count} row(s) with figures have no name and were left out.",
        "issue_cross_not_additive": "\"{header}\" is a rate, an index or a mean, or has no whole to divide by, so it has no share of a total.",
        "issue_latest_grouped": "In {count} places column {column} has a date only on the first row of a group, so the rows below it are not being read.",
        "issue_latest_partial_year": "The newest value looks like a part of a year, up to {date} ({months} months). Comparing it with whole years overstates the change.",
        "issue_latest_unfinished": "The period of {date} had not ended when this data was last saved ({saved}), so the value may be incomplete.",
        "issue_calc_columns": "\"{key}\" needs exactly two columns: the dates, then the values. For example week:Daily!A,C.",
        "issue_calc_no_base": "There is nothing to compare {date} with: the day is not in the sheet, there are too few rows, or the base is 0.",
        "issue_insight_not_chosen": "No sentence chosen yet, so the calculated sentence was used. Press Ctrl+I to see candidates.",
        "issue_insight_stale": "The figures changed since a sentence was chosen, so the calculated sentence was used. Press Ctrl+I to choose again.",
        "helper_button": "Suggest references (Ctrl+R)",
        "helper_no_excel": "Choose a workbook first. The helper reads it and suggests what to put in the template.",
        "helper_title": "Suggest references",
        "helper_intro": "SlideFill read this workbook and tried each of these. What you see is what the report will show. Copy one into your template, or save a starter template with them already in place.",
        "helper_sheet": "Sheet",
        "helper_base": "By",
        "helper_value": "Figure",
        "helper_how": "Calculate",
        "helper_group1": "Only where",
        "helper_word1": "is",
        "helper_group2": "and where",
        "helper_word2": "is",
        "helper_none": "(all rows)",
        "helper_no_name": "(no heading)",
        "helper_every_row": "Every row added up",
        "helper_counted": "How many rows",
        "helper_how_plain": "As they are",
        "helper_how_sum": "Add up",
        "helper_how_mean": "Mean",
        "helper_how_count": "Count rows",
        "helper_how_rate": "Share matching the condition",
        "helper_col_what": "What",
        "helper_col_result": "What the report will show",
        "helper_col_place": "Where it goes",
        "helper_place_text": "Text",
        "helper_place_alt": "Table or chart alt text",
        "helper_where_text": "Paste it into any text in your template.",
        "helper_where_alt": "Right-click a table or chart in your template, choose View Alt Text, and paste it there.",
        "helper_copy": "Copy (Ctrl+C)",
        "helper_copied": "Copied: {reference}",
        "helper_starter": "Save a starter template",
        "helper_starter_saved": "Saved and set as the template: {path}",
        "helper_working": "Reading the workbook. A large one can take a few minutes.",
        "helper_pick_row": "Select a row to see where it goes.",
        "helper_nothing": "No dates or named figures could be read from this sheet. Try another sheet, or press F1 to write references by hand.",
        "helper_rate_needs_condition": "A share needs a condition. Choose which rows count under Only where.",
        "helper_shape_dates": "\"{sheet}\": dates in {spec}, {first} to {last}, {count} periods.",
        "helper_shape_names": "\"{sheet}\": named items in {spec}, {count} names, no dates.",
        "helper_shape_joined": "Continues across the sheets {pattern}.",
        "helper_note_first_table": "This sheet holds {count} tables, so only the first is used.",
        "helper_note_last_form": "The dates change form partway down, such as years and then quarters, so only the last and finest part is used.",
        "helper_note_repeat": "{date} appears twice, so only the {count} periods after it are used. Check that date in the workbook.",
        "err_helper_empty_sheet": "No dates or named figures could be read from \"{sheet}\".",
        "suggest_latest": "Newest value",
        "suggest_note": "Report sentence",
        "suggest_reading": "How to read it",
        "suggest_change": "What moved between two periods",
        "suggest_prev": "Change from the previous",
        "suggest_week": "Change from last week",
        "suggest_avg": "Change from the recent mean",
        "suggest_peak": "When the recent 30 peaked",
        "suggest_insight": "Chosen insight",
        "suggest_cross": "Across the files",
        "suggest_doc_note": "The document and the figures",
        "suggest_doc_quote": "From the document",
        "suggest_doc_reading": "How far that goes",
        "suggest_table_doc": "What the document records, and the figures around it",
        "suggest_table": "Table of the newest 6",
        "suggest_chart": "Chart of the newest 12",
        "suggest_top": "Name of the largest",
        "suggest_topvalue": "Value of the largest",
        "suggest_share": "Share of the largest",
        "suggest_table_share": "Table of the top 5, with shares",
        "suggest_table_top": "Table of the top 5",
        "suggest_chart_top": "Chart of the top 5",
        "suggest_by_item": "By item, newest period",
        "suggest_table_by_item": "Table by item, newest period",
        "insight_button": "Suggest insights (Ctrl+I)",
        "insight_title": "Suggest insights",
        "insight_warning": "An insight is one sentence that says what a figure means, not just what it is. Whoever reads the report gets the point without working it out from the table.\nChoose one and it goes in the report. Choose none and the sentence that is there now stays. Those marked calculated come from the program's own results. Those marked model are written by a model on this PC, and nothing leaves this PC. Some were wrong in testing even after the checks, so read each against its basis first.",
        "insight_candidates": "Suggested sentences",
        "insight_mark_calculated": "[calculated] ",
        "insight_mark_model": "[model, check it] ",
        "insight_add": "Add to sentence",
        "insight_again": "Make more",
        "insight_chosen": "Sentence for the report (you can edit it; leave it empty to use the calculated sentence)",
        "insight_apply": "Use in report",
        "insight_close": "Close",
        "insight_basis": "Calculated results this one cites:",
        "insight_hypotheses": "Ideas to check yourself (these do not go in the report):",
        "insight_working": "Writing, round {run} of {runs}. {count} so far.",
        "insight_done": "{count} sentences. The checks left out {rejected} in the latest round.",
        "insight_none": "No sentence passed the checks. {rejected} were left out. Press Make more, or use the calculated sentence.",
        "insight_err_ollama_missing": "Could not reach Ollama on this PC. Install or start Ollama, then press Make more. The report still works with the calculated sentence.",
        "insight_err_model_missing": "The model is not installed. In a terminal, run: ollama pull {model}",
        "insight_err_timeout": "The model took too long to answer. Press Make more to try again.",
        "insight_err_other": "The model could not answer: {detail}",
    },
    "ko": {
        "window_title": "SlideFill",
        "tagline": "엑셀을 고르면 PowerPoint 보고서를 만듭니다.",
        "own_template": "내 PowerPoint 템플릿 사용 (선택)",
        "add_files": "파일 추가…",
        "remove_files": "빼기",
        "drop_hint": "엑셀과 Word 파일을 여기로 끌어다 놓거나 파일 추가를 누르세요.  .xlsx .xls .xlsb .ods .csv .tsv .docx",
        "err_doc_needs_workbook": "Word 문서는 수치와 맞춰 읽으므로 엑셀 파일도 함께 넣어 주세요.",
        "err_wrong_type_document": "읽을 수 없는 종류의 문서입니다. .docx로 저장해 주세요.\n{path}",
        "drop_refused": "읽을 수 없는 종류의 파일입니다: {names}",
        "report_many": "{first} 외 {more}개",
        "report_many_file": "{first}_외_{more}개",
        "auto_ready_unread": "엑셀을 읽고 내용을 골랐습니다: {what}. 읽지 못한 파일: {names}. 값을 확인한 뒤 보고서를 만드세요.",
        "auto_ready_doubt": "엑셀을 읽고 내용을 골랐습니다: {what}. 확인이 필요한 점이 {count}개 있어 보고서를 만들 때 먼저 여쭙니다. 다른 내용으로 바꾸려면 Ctrl+R을 누르세요.",
        "doubt_title": "수치를 확인해 주세요",
        "doubt_intro": "SlideFill이 수치를 스스로 골랐는데, 잘못 골랐을 수 있는 징후가 있습니다. 보고서는 아직 만들지 않았습니다.",
        "doubt_report_note": "※ 확인되지 않은 값입니다. {reason}",
        "doubt_chosen": "고른 내용: {what}",
        "doubt_change": "수치 고르기…",
        "doubt_keep": "이대로 만들기",
        "doubt_cancel": "취소",
        "doubt_example_sheet": "시트 이름이 예시나 양식이라 실제 자료가 아닐 수 있습니다.",
        "doubt_count_only": "더해도 되는 수치를 찾지 못해 줄 수를 셉니다. 이 파일의 목적과 다를 수 있습니다.",
        "doubt_sum_of_levels": "{figure}을(를) 줄마다 더했는데, 발열량이나 농도처럼 더하면 뜻이 없는 값으로 보입니다.",
        "doubt_first_table": "이 시트에 표가 여러 개 쌓인 것으로 보여 첫 표만 읽었습니다.",
        "doubt_heading_below": "이 열에는 첫 숫자 아래에 다른 머리글이 있어, {figure}라는 이름이 다른 표의 것일 수 있습니다.",
        "doubt_dates_sparse": "날짜가 없는 줄이 많아 날짜별 합계에서 빠졌습니다.",
        "doubt_maybe_percent": "시트에 퍼센트라는 말이 있어 {figure}이(가) 개수가 아니라 비율일 수 있습니다.",
        "doubt_several_tables": "한 시트에 표 여러 개가 이어져 있는데 한 목록으로 읽었습니다.",
        "doubt_total_differs": "시트에 적힌 합계와 항목을 더한 값이 달라 겹치는 줄이 있을 수 있습니다.",
        "doubt_period_columns": "{figure}은(는) 여러 기간 열 중 첫 열로 보여 최신 값이 아닐 수 있습니다.",
        "doubt_mixed_kinds": "{figure}을(를) 종류가 다른 줄까지 더했는데 크기 차이가 큽니다. 서로 다른 지표라면 더한 값은 뜻이 없습니다.",
        "doubt_figures_not_read": "숫자가 들어 있지만 수치로 읽지 못한 칸이 있어 결과에서 빠졌습니다.",
        "doubt_count_among_shares": "{figure}은(는) 퍼센트 열인데 {item}은(는) 100이 넘습니다. 사람 수나 건수가 비율 사이에 섞였을 수 있습니다.",
        "doubt_sparse_period": "{period} 줄은 평소 {usual}칸 중 {filled}칸만 채워져 있습니다. 값이 빠진 기간과 비교했을 수 있습니다.",
        "doubt_merged_names": "수치가 따로 있는 {count}줄이 병합된 이름 {item} 하나를 함께 쓰고, 줄마다 그 이름으로 비교했습니다. 한 줄의 값이 {item} 전체처럼 읽힐 수 있습니다.",
        "doubt_people_across_kinds": "{figure}을(를) {kind}({values})를 건너 더했습니다. 한 사람이 여러 {kind}에 들 수 있어, 합은 같은 사람을 두 번 셀 수 있고 사람 수가 아닐 수 있습니다.",
        "doubt_nameless_total": "{item}은(는) 줄마다 {same}이라 고르지 않았습니다. 그 아래 이름 없는 줄에 {total}이 있어, 이 값이 시트 전체의 합계일 수 있습니다.",
        "doubt_parts_by_sum": "{figure}에서 {parent} 값이 바로 아래 행({parts})의 합과 같아, 그 행들을 {parent}의 일부로 보고 비교에서 뺐습니다. 합이 우연히 맞은 것이라면 그 행들이 빠졌습니다.",
        "template_first_file": "템플릿은 파일 하나만 읽으므로 {name}만 씁니다. 모든 파일로 보고서를 만들려면 템플릿 사용을 끄세요.",
        "insight_about": "대상",
        "insight_all_files": "모든 파일 함께 보기 ({files})",
        "contents_figures": "넣을 내용",
        "contents_preview": "이렇게 만들어집니다 (슬라이드 {n}/{total})",
        "contents_preview_none": "고른 것이 없습니다",
        "contents_detail": "계산 방법 바꾸기…",
        "contents_count": "슬라이드 {slides}장, 수치 {figures}개",
        "contents_change": "지금: {before}.  바꾸면: {after}.",
        "contents_same": "지금 보고서 그대로입니다: {now}. 체크를 바꾸면 오른쪽 슬라이드가 바로 바뀝니다.",
        "contents_none": "수치를 하나 이상 체크하세요.",
        "contents_limit": "슬라이드 한 장에는 시트당 수치 {count}개까지 넣을 수 있습니다.",
        "contents_skipped": "큰 시트 {count}개는 창이 느려지지 않도록 목록에서 뺐습니다.",
        "auto_ready": "엑셀을 읽고 내용을 골랐습니다: {what}. 값을 확인한 뒤 보고서를 만드세요. 다른 내용으로 바꾸려면 Ctrl+R을 누르세요.",
        "err_auto_nothing": "이 엑셀에서 날짜나 이름 붙은 수치를 읽지 못해 보고서를 자동으로 구성할 수 없습니다.\n\n\"내 PowerPoint 템플릿 사용\"을 켜면 무엇을 어디에 넣을지 직접 정할 수 있습니다. F1을 눌러 방법을 확인하세요.",
        "helper_button_auto": "내용 바꾸기 (Ctrl+R)",
        "helper_title_auto": "보고서 내용 바꾸기",
        "auto_ready_left": "엑셀을 읽고 내용을 골랐습니다: {what}. 넣지 않은 것: 수치 {figures}개, 시트 {sheets}개. 값을 확인한 뒤 보고서를 만드세요. 다른 내용으로 바꾸려면 Ctrl+R을 누르세요.",
        "helper_intro_auto": "왼쪽에서 넣을 수치를 체크하세요. 오른쪽은 보고서에 실제로 들어갈 슬라이드이고, 체크를 바꾸면 바로 바뀝니다. \"이 내용으로 하기\"를 누르기 전에는 보고서가 바뀌지 않습니다.",
        "helper_use": "이 내용으로 하기",
        "helper_starter_auto": "수정할 수 있는 템플릿으로 저장",
        "template": "템플릿 (.pptx)",
        "excel": "엑셀 파일",
        "output": "저장 폴더",
        "browse": "찾아보기…",
        "preview": "미리보기 (F5)",
        "generate": "보고서 만들기 (Ctrl+G)",
        "open_result": "보고서 열기",
        "open_folder": "폴더 열기",
        "help": "도움말 (F1)",
        "language": "English",
        "pick_template": "PowerPoint 템플릿 선택",
        "pick_excel": "엑셀 파일 선택",
        "pick_output": "보고서 저장 폴더 선택",
        "filetype_pptx": "PowerPoint 프레젠테이션",
        "filetype_excel": "엑셀, 데이터, Word 파일",
        "col_status": "상태",
        "col_item": "항목",
        "col_value": "값",
        "col_source": "가져오는 곳",
        "col_where": "위치",
        "status_ok": "정상",
        "status_warn": "확인",
        "status_error": "문제",
        "kind_text": "텍스트",
        "kind_table": "표",
        "kind_chart": "차트",
        "kind_unsupported": "기타",
        "where_slide": "슬라이드 {n}",
        "where_notes": "노트 {n}",
        "where_master": "마스터 {n}",
        "where_layout": "레이아웃: {n}",
        "where_part": "{n}",
        "start_hint": "엑셀을 고르면 보고서에 들어갈 내용을 보여 줍니다.",
        "working": "파일을 읽는 중…",
        "generating": "보고서를 만드는 중…",
        "summary": "정상 {ok} · 확인 {warn} · 문제 {error}",
        "no_placeholders": "이 템플릿에서 {{자리표시자}}를 찾지 못했습니다. F1을 눌러 넣는 방법을 확인하세요.",
        "saved": "저장됨: {path}",
        "select_hint": "행을 선택하면 자세한 내용이 나옵니다.",
        "confirm_title": "채우지 못한 항목이 있습니다",
        "confirm_body": "{count}개 항목에 문제가 있어 보고서에 {{{{자리표시자}}}} 그대로 남습니다.\n\n그래도 만들까요?",
        "error_title": "SlideFill",
        "missing_inputs": "엑셀을 먼저 고르세요.",
        "help_title": "SlideFill 사용법",
        "help_body": """엑셀만 고르면 됩니다. SlideFill이 파일을 읽고 내용을 골라, 보고서를 만들기 전에 모든 값을 보여 줍니다. 수치를 잘못 골랐다면 Ctrl+R을 눌러 바꾸세요.

이 아래는 회사 템플릿을 쓸 때의 설명입니다. "내 PowerPoint 템플릿 사용"을 켜고, 템플릿에 무엇을 어디에 넣을지 표시합니다.

1. PowerPoint에서 템플릿에 자리표시자를 입력합니다.

텍스트 (텍스트 상자, 표 칸, 노트, 차트 제목, 슬라이드 마스터 어디든)
    {{매출}}              엑셀의 이름 정의 (수식 > 이름 정의)
    {{요약!B3}}           시트와 셀 주소
    {{'매출 요약'!B3}}    시트 이름에 공백이 있어도 따옴표는 없어도 됩니다

표와 차트: 표나 차트를 우클릭 > 대체 텍스트 보기에서 입력합니다
    {{table:요약!A1:D10}}   또는 {{표:매출표}}
    {{chart:요약!A1:D5}}    또는 {{차트:추이범위}}

이름에는 이름 정의나, 범위라면 엑셀 표 이름을 쓸 수 있습니다.

표는 왼쪽 위 칸부터 채우므로 범위에 머리글 행을 포함하세요. 행은 데이터에 맞춰 늘거나 줄어듭니다. 열 수는 같아야 합니다.

차트는 첫 행을 계열 이름으로, 첫 열을 항목으로 읽습니다.

매일 쌓이는 데이터: 날짜 열을 맨 앞에 적으면 날짜로 가장 최근 행을 찾습니다. 소계 행과, 값이 아직 없는 미래 날짜는 건너뜁니다.
    {{최신:일별!A,C}}             A열 날짜가 가장 최근인 행의 C열 값
    {{최신:일별!A}}               가장 최근 날짜
    {{표:최신7:일별!A,C,D}}       최근 7행, 오래된 날짜부터
    {{차트:최신7:일별!A,C,D}}     같은 행을 차트로
표는 템플릿에 이미 있는 머리글 행을 그대로 둡니다. 머리글이 비어 있으면, 그리고 차트 계열 이름은 시트에서 가져옵니다.

SlideFill이 날짜 열과 값 열 하나로 직접 계산합니다:
    {{전주대비:일별!A,C}}     전주 같은 요일 대비 변화
    {{전일대비:일별!A,C}}     직전 값 대비 변화 ({{직전대비:...}}도 같음)
    {{평균대비7:일별!A,C}}    직전 7개 평균 대비 변화
    {{최고일30:일별!A,C}}     최근 30개 중 값이 가장 높은 날 (가장 낮은 날은 {{최저일30:...}})
    {{해설:일별!A,C}}         이 수치를 담은 보고 문장
    {{읽는법:일별!A,C}}       정해진 규칙으로 수치를 읽는 법: 평균에 섞인 주말, 평소와 크게 다른 날, 평소 범위를 벗어난 값, 전년 같은 달 대비 등
월별, 분기별, 연도별 데이터는 날짜를 보고 알아냅니다. 연도 열에서는 2021처럼 연도만 적어도 날짜로 씁니다.

이 표기를 직접 쓰지 않아도 됩니다. Ctrl+R을 누르면 SlideFill이 엑셀을 읽어 결과와 함께 보여 주고, 복사하거나 초안 템플릿으로 저장해 줍니다.

데이터 모양에 맞춰 위 표기 어디에나 쓸 수 있습니다:
    최신:일별!A,합계(D:E)        날짜별로 행을 더합니다. 합계라고 적을 때만 더합니다
    최신:일별!A,합계(D, B=동부)  B열에 동부가 들어간 행만
    최신:일별!A,건수(C)          날짜별 행 수
    최신:일별!A,합계(D*E)        수량 × 단가
    최신:일별!A,평균(D)          날짜별 행의 평균
    최신:일별!A,비율(B=Yes)      날짜별 행 중 B열에 Yes가 들어간 비율
    최신:월별!A+B,C              A열에 연도, B열에 월
    최신:월별!5,7                5행에 기간, 7행에 값 (가로로 펼친 표)
    최신:월별!C2:N2,C3:N3        범위 지정. 한 시트의 표 두 개를 구분할 때
    최신:연도*!A,C               연도로 시작하는 모든 시트를 이어서

날짜 없는 표 (첫 열이 항목 이름):
    {{읽는법:표!A,B}}       가장 큰 항목과 구성비. A,B,C로 적으면 두 기간 사이 증감
    {{순위1:표!A,B}}        가장 큰 항목의 이름 ({{순위값1:...}}은 값, {{구성비1:...}}은 구성비)
    {{표:순위5:표!A,B}}     상위 5개 ({{표:구성비5:...}}는 구성비 포함, {{차트:순위5:...}}도 가능)
비교할 값이 없으면 자리표시자를 그대로 두고 문제로 표시합니다.

이 PC의 모델이 쓴 후보에서 직접 고르는 문장:
    {{통찰:일별!A,C}}       Ctrl+I를 눌러 후보를 읽고 고릅니다
고르기 전에는 계산 해설 문장이 들어가고, 수치가 바뀌면 다시 골라야 합니다. Ollama와 모델이 필요하며, 없어도 나머지 기능은 그대로 됩니다.

2. 여기서 템플릿과 엑셀을 고릅니다. 만들기 전에 목록에서 모든 값을 확인할 수 있습니다.

3. 보고서 만들기를 누릅니다. 템플릿과 엑셀은 바뀌지 않고, 보고서는 항상 새 파일로 저장됩니다.

알아 두세요
- 값은 엑셀에 마지막으로 저장된 값입니다. 엑셀을 먼저 저장하세요.
- 숫자는 엑셀에 보이는 대로 들어갑니다: 85.3%, 1,234,567, 2026-11-03
- 채우지 못한 것은 {{자리표시자}} 그대로 남아 눈에 보입니다. 빈칸이라면 엑셀 셀이 비어 있는 것입니다.
- 암호나 문서 보안(DRM)이 걸린 파일은 읽지 못합니다.

단축키
    F5  미리보기          Ctrl+G  보고서 만들기
    Ctrl+T  템플릿        Ctrl+E  엑셀
    Ctrl+L  언어          F1  도움말""",
        "err_file_missing": "파일을 찾을 수 없습니다:\n{path}",
        "err_file_locked": "Windows가 이 파일을 열어 주지 않습니다. 다른 프로그램에서 닫고 다시 시도하세요.\n{path}",
        "err_file_unreadable": "이 파일을 읽지 못했습니다. 손상되었을 수 있습니다.\n{path}\n\n{detail}",
        "err_file_protected": "암호나 문서 보안(DRM)이 걸려 있거나 구형 포맷인 파일입니다. 보호를 풀거나 일반 Office 파일로 저장한 뒤 다시 시도하세요.\n{path}",
        "err_file_not_office": "일반 Office 파일이 아닙니다. 문서 보안(DRM)이 걸려 있을 수 있습니다.\n{path}",
        "err_wrong_type_excel": "읽을 수 없는 종류의 파일입니다. .xlsx .xlsm .xls .xlsb .ods .csv .tsv 파일을 읽습니다.\n{path}",
        "err_wrong_type_template": "템플릿은 .pptx 파일이어야 합니다. PowerPoint에서 다른 이름으로 저장해 변환하세요.\n{path}",
        "err_cannot_save": "보고서를 저장하지 못했습니다. 폴더를 확인하세요.\n{path}\n\n{detail}",
        "err_unexpected": "예상하지 못한 문제가 생겼습니다.\n\n{detail}",
        "issue_missing_name": "\"{name}\"이라는 이름 정의나 엑셀 표가 없습니다.",
        "issue_missing_sheet": "\"{sheet}\" 시트가 없습니다.",
        "issue_bad_ref": "\"{ref}\"는 올바른 셀 주소가 아닙니다.",
        "issue_range_for_text": "{source}는 여러 셀입니다. 텍스트에는 한 셀만 쓸 수 있습니다. 표라면 대체 텍스트를 쓰세요.",
        "issue_name_not_cells": "이름 \"{name}\"은 셀이 아니라 수식이나 상수입니다. 셀을 가리키도록 바꾸세요.",
        "issue_multi_area": "이름 \"{name}\"이 여러 곳을 가리킵니다. 한 범위만 쓰세요.",
        "issue_formula_no_cache": "{cell}의 수식에 저장된 결과가 없습니다. Excel에서 열어 저장하세요.",
        "issue_error_value": "{cell}이 Excel에서 {error} 오류를 표시합니다.",
        "issue_table_error_value": "{cell}이 Excel에서 {error} 오류를 표시합니다. 그대로 옮겼습니다.",
        "issue_empty_cell": "{cell}이 비어 있어 빈칸이 되었습니다.",
        "issue_format_fallback": "{cell}의 표시 형식({format})을 지원하지 않아 원래 값을 넣었습니다.",
        "issue_marker_in_text": "표·차트 표시는 텍스트가 아니라 표나 차트의 대체 텍스트에 넣으세요.",
        "issue_marker_wrong_shape": "{kind} 표시가 {kind}가 아닌 \"{shape}\"에 있습니다.",
        "issue_column_mismatch": "범위는 {range}열인데 표는 {table}열입니다.",
        "issue_merged_resize": "본문에 세로로 병합된 칸이 있거나 머리글 아래에 본문 행이 없어 행을 늘리거나 줄일 수 없습니다. 행 수를 맞추세요.",
        "issue_figures_not_read": "숫자가 들어 있지만 수치로 읽지 못한 칸 {count}개가 빠졌습니다(예: \"{text}\").",
        "issue_table_covers": "표가 {rows}행이 되면 아래의 \"{name}\"까지 닿아 덮을 가능성이 큽니다. 템플릿에서 표 아래 공간을 넓히세요.",
        "issue_latest_partial_rows": "표시할 행 중 {count}개는 일부 열이 비어 있습니다(예: {first}). 아직 값이 없는 기간이면 빈칸으로 나옵니다.",
        "issue_cross_backwards": "\"{second}\"가 \"{first}\"보다 이전 기간으로 보이는데, 증감은 첫 열에서 둘째 열로 계산합니다. 늘어난 값이 줄어든 것으로 읽힐 수 있으니 참조의 두 열 순서를 바꾸세요.",
        "issue_merged_value_lost": "\"{value}\"가 병합된 칸에 해당해 표시되지 않았습니다.",
        "issue_table_overflow": "{rows}행과 글자 줄바꿈을 합치면 표가 슬라이드 아래로 넘칠 가능성이 큽니다.",
        "issue_chart_too_small": "차트 범위는 2행 2열 이상이어야 합니다.",
        "issue_chart_unsupported": "거품형 차트는 지원하지 않습니다.",
        "issue_chart_not_number": "\"{value}\"는 숫자가 아니어서 차트에서 비워 둡니다.",
        "issue_chart_header_number": "첫 행의 \"{value}\"가 계열 이름이 되어 그 행은 차트에 그려지지 않습니다. 데이터라면 머리글 행까지 범위에 넣으세요.",
        "issue_unsupported_location": "슬라이드 {slide}에서 SlideFill이 읽지 못하는 곳(특이한 모양으로 저장된 SmartArt 등)에 있는 자리표시자는 그대로 남습니다. 일반 텍스트 상자로 옮기세요.",
        "issue_latest_bad": "\"{key}\"는 올바른 표기가 아닙니다. 최신:일별!A,C, 최신7:일별!A,C,D, 최신:월별!5,7(행)처럼 날짜를 맨 앞에 쓰고, 값에는 합계(D:E), 건수(A)도 쓸 수 있습니다.",
        "issue_latest_no_dates": "\"{sheet}\" 시트 {column}열에서 값이 있는 날짜 행을 찾지 못했습니다.",
        "issue_latest_ambiguous": "{date}에 값이 있는 행이 {count}개라 하나를 고를 수 없습니다. 날짜별 합계는 합계(C), 한 종류만은 합계(C, B=이름), 한 시트의 표 둘은 A10:A400,C10:C400처럼 적으세요.",
        "issue_latest_fewer": "{wanted}행을 요청했지만 날짜와 값이 있는 행은 {found}개뿐입니다.",
        "issue_latest_older": "{newest} 값이 아직 없어 {used} 값을 썼습니다. 보고서의 다른 값은 더 최근 날짜일 수 있습니다.",
        "issue_latest_duplicates": "{date}까지 날짜 {count}개가 여러 행에 있어 비교할 수 없습니다. 한 시트의 표 둘이면 A10:A400,C10:C400처럼 범위를 정하고, 날짜별 합계라면 합계(C)로 적으세요.",
        "issue_latest_sheets_overlap": "{date}이 \"{first}\"와 \"{second}\"에 모두 있어 더하면 두 번 세게 됩니다. 겹치지 않게 시트 이름을 적으세요.",
        "issue_latest_sheets_differ": "{line}의 머리글이 \"{first}\"와 \"{second}\"에서 다릅니다. 시트마다 다른 항목일 수 있습니다.",
        "issue_latest_zero": "최신 값인 {date}이 0입니다. 빈 기간의 수식이 미리 채운 0이라면 이 값은 최신 수치가 아닙니다.",
        "issue_latest_text_newer": "더 최근 {count}개 기간에 수치 대신 \"{text}\"이 있어 {used} 값을 썼습니다.",
        "issue_cross_nested": "아래 행들의 합인 행이 있어, 나란히 비교할 수 있는 {count}개만 비교합니다.",
        "issue_cross_outline": "항목 이름에 계층 번호가 있어 최상위 {count}개만 비교하고 구성비는 내지 않습니다.",
        "issue_cross_several_totals": "재무제표의 수입과 지출처럼 서로 다른 것의 합계 줄이 {count}개({names}) 있어, 줄끼리 순위를 매기지 않고 합계 줄만 적습니다.",
        "issue_latest_group_size": "{date}의 행이 {found}개입니다. 날짜마다 보통 {usual}개이므로 합계가 덜 쌓였거나 다른 행이 섞였을 수 있습니다.",
        "issue_latest_partial_day": "{date}의 마지막 기록이 {last}입니다. 보통 {usual}까지 있으므로 그날이 다 쌓이지 않았을 수 있습니다.",
        "issue_number_from_text": "글자로 적힌 숫자가 {count}칸 있습니다. 예: {cell}의 \"{text}\". 앞의 숫자를 읽었습니다.",
        "issue_cross_no_items": "\"{sheet}\" 시트 {column}열에 날짜가 없고, 수치가 있는 이름 행도 2개 미만입니다.",
        "issue_cross_total_differs": "시트에 적힌 합계 {stated}와 항목의 합 {added}이 다릅니다. 행이 겹칠 수 있어 구성비는 내지 않습니다.",
        "issue_cross_total_found": "\"{label}\"이 나머지 행의 합과 같아 항목이 아닌 합계로 봤습니다.",
        "issue_cross_no_total": "시트에 합계가 없어 항목을 더한 {added}을 기준으로 구성비를 냈습니다. 지수나 점수처럼 더할 수 없는 값이면 구성비는 뜻이 없습니다.",
        "issue_cross_unlabeled": "수치는 있지만 이름이 없는 행 {count}개를 뺐습니다.",
        "issue_cross_not_additive": "\"{header}\"는 비율, 지수, 평균이거나 나눌 전체가 없어 구성비를 낼 수 없습니다.",
        "issue_latest_grouped": "{column}열의 날짜가 묶음의 첫 행에만 있는 곳이 {count}곳입니다. 그 아래 행은 읽지 않았습니다.",
        "issue_latest_partial_year": "최신 값은 {date}까지의 {months}개월치로 보입니다. 1년치 값과 그대로 비교하면 변화가 크게 나옵니다.",
        "issue_latest_unfinished": "{date}이 속한 기간은 데이터를 마지막으로 저장한 {saved}에 아직 끝나지 않았습니다. 값이 다 쌓이지 않았을 수 있습니다.",
        "issue_calc_columns": "\"{key}\"에는 날짜 열과 값 열, 두 열만 적어야 합니다. 예: 전주대비:일별!A,C",
        "issue_calc_no_base": "{date}와 비교할 값이 없습니다. 그 날짜가 시트에 없거나, 행이 모자라거나, 기준값이 0입니다.",
        "issue_insight_not_chosen": "아직 고른 문장이 없어 계산 해설 문장을 넣었습니다. Ctrl+I를 누르면 후보를 볼 수 있습니다.",
        "issue_insight_stale": "문장을 고른 뒤 수치가 바뀌어 계산 해설 문장을 넣었습니다. Ctrl+I를 눌러 다시 고르세요.",
        "helper_button": "참조 추천 (Ctrl+R)",
        "helper_no_excel": "먼저 엑셀을 고르세요. 도우미가 파일을 읽고 템플릿에 넣을 것을 추천합니다.",
        "helper_title": "참조 추천",
        "helper_intro": "SlideFill이 이 엑셀을 읽고 아래 항목을 실제로 계산해 봤습니다. 보이는 값이 보고서에 들어가는 값입니다. 하나를 복사해 템플릿에 넣거나, 이미 다 넣어 둔 초안 템플릿을 저장하세요.",
        "helper_sheet": "시트",
        "helper_base": "무엇별로",
        "helper_value": "수치",
        "helper_how": "계산",
        "helper_group1": "이 열이",
        "helper_word1": "이 값인 행만",
        "helper_group2": "그리고 이 열이",
        "helper_word2": "이 값인 행만",
        "helper_none": "(모든 행)",
        "helper_no_name": "(머리글 없음)",
        "helper_every_row": "모든 행 합계",
        "helper_counted": "행 수",
        "helper_how_plain": "그대로",
        "helper_how_sum": "더하기",
        "helper_how_mean": "평균",
        "helper_how_count": "개수 세기",
        "helper_how_rate": "조건에 맞는 비율",
        "helper_col_what": "무엇",
        "helper_col_result": "보고서에 들어갈 값",
        "helper_col_place": "넣는 곳",
        "helper_place_text": "텍스트",
        "helper_place_alt": "표·차트의 대체 텍스트",
        "helper_where_text": "템플릿의 아무 텍스트에나 붙여 넣으세요.",
        "helper_where_alt": "템플릿의 표나 차트를 우클릭해 대체 텍스트 보기를 고른 뒤 붙여 넣으세요.",
        "helper_copy": "복사 (Ctrl+C)",
        "helper_copied": "복사했습니다: {reference}",
        "helper_starter": "초안 템플릿 저장",
        "helper_starter_saved": "저장하고 템플릿으로 지정했습니다: {path}",
        "helper_working": "엑셀을 읽는 중입니다. 큰 파일은 몇 분 걸릴 수 있습니다.",
        "helper_pick_row": "행을 고르면 넣는 곳을 알려 줍니다.",
        "helper_nothing": "이 시트에서 날짜나 이름 붙은 수치를 읽지 못했습니다. 다른 시트를 고르거나, F1을 눌러 직접 적으세요.",
        "helper_rate_needs_condition": "비율에는 조건이 필요합니다. \"이 열이\"에서 어떤 행을 셀지 고르세요.",
        "helper_shape_dates": "\"{sheet}\": {spec}에 날짜, {first} ~ {last}, {count}개 기간.",
        "helper_shape_names": "\"{sheet}\": {spec}에 항목 이름 {count}개, 날짜 없음.",
        "helper_shape_joined": "{pattern} 시트를 이어서 읽습니다.",
        "helper_note_first_table": "이 시트에 표가 {count}개 있어 첫 표만 씁니다.",
        "helper_note_last_form": "날짜 표기가 중간에 바뀌어(연도 다음 분기 등) 마지막 부분만 씁니다.",
        "helper_note_repeat": "{date}이 두 번 있어 그 뒤 {count}개 기간만 씁니다. 엑셀에서 그 날짜를 확인하세요.",
        "err_helper_empty_sheet": "\"{sheet}\"에서 날짜나 이름 붙은 수치를 읽지 못했습니다.",
        "suggest_latest": "최신 값",
        "suggest_note": "보고 문장",
        "suggest_reading": "읽는 법",
        "suggest_change": "두 기간 사이 증감",
        "suggest_prev": "직전 대비",
        "suggest_week": "전주 대비",
        "suggest_avg": "최근 평균 대비",
        "suggest_peak": "최근 30개 중 가장 높은 때",
        "suggest_insight": "고른 통찰 문장",
        "suggest_cross": "파일을 함께 본 문장",
        "suggest_doc_note": "문서와 수치",
        "suggest_doc_quote": "문서 내용",
        "suggest_doc_reading": "읽을 때 주의할 점",
        "suggest_table_doc": "문서 기록과 전후 수치 표",
        "suggest_table": "최근 6개 표",
        "suggest_chart": "최근 12개 차트",
        "suggest_top": "가장 큰 항목 이름",
        "suggest_topvalue": "가장 큰 항목 값",
        "suggest_share": "가장 큰 항목 구성비",
        "suggest_table_share": "상위 5개 표 (구성비 포함)",
        "suggest_table_top": "상위 5개 표",
        "suggest_chart_top": "상위 5개 차트",
        "suggest_by_item": "최신 기간 항목별",
        "suggest_table_by_item": "최신 기간 항목별 표",
        "insight_button": "통찰 문장 제안 (Ctrl+I)",
        "insight_title": "통찰 문장 제안",
        "insight_warning": "통찰 문장은 숫자를 옮겨 적는 데서 그치지 않고 그 숫자가 무슨 뜻인지 짚어 주는 한 문장입니다. 보고받는 사람이 표를 직접 해석하지 않아도 요점을 바로 알 수 있습니다.\n하나를 고르면 보고서에 들어가고, 고르지 않으면 지금 문장이 그대로 나갑니다. [계산]은 프로그램이 계산 결과로 쓴 문장입니다. [모델]은 이 PC의 모델이 쓴 문장이고 데이터는 PC 밖으로 나가지 않습니다. 시험에서 검사를 통과하고도 틀린 문장이 있었으니 근거와 대조한 뒤 쓰세요.",
        "insight_candidates": "제안된 문장",
        "insight_mark_calculated": "[계산] ",
        "insight_mark_model": "[모델·확인 필요] ",
        "insight_add": "문장에 추가",
        "insight_again": "더 만들기",
        "insight_chosen": "보고서에 넣을 문장 (고칠 수 있습니다. 비워 두면 계산 해설 문장을 씁니다)",
        "insight_apply": "보고서에 사용",
        "insight_close": "닫기",
        "insight_basis": "이 문장이 근거로 든 계산 결과:",
        "insight_hypotheses": "직접 확인할 생각거리 (보고서에 들어가지 않습니다):",
        "insight_working": "문장을 만드는 중입니다. {runs}회 중 {run}회째, 지금까지 {count}개.",
        "insight_done": "문장 {count}개. 이번 회차에 {rejected}개는 검사에서 뺐습니다.",
        "insight_none": "검사를 통과한 문장이 없습니다. {rejected}개를 제외했습니다. 더 만들기를 누르거나 계산 해설 문장을 쓰세요.",
        "insight_err_ollama_missing": "이 PC의 Ollama에 연결하지 못했습니다. Ollama를 설치하거나 실행한 뒤 더 만들기를 누르세요. 보고서는 계산 해설 문장으로 그대로 만들 수 있습니다.",
        "insight_err_model_missing": "모델이 설치되어 있지 않습니다. 터미널에서 실행하세요: ollama pull {model}",
        "insight_err_timeout": "모델 응답이 너무 오래 걸렸습니다. 더 만들기를 눌러 다시 시도하세요.",
        "insight_err_other": "모델이 응답하지 못했습니다: {detail}",
    },
}


def text(lang, key, /, **params):  # positional only: some messages have a {key} of their own
    table = STRINGS.get(lang, STRINGS["en"])
    template = table.get(key, STRINGS["en"].get(key, key))
    return template.format(**params) if params else template
