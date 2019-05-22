# Generator_ZZPB05

## Program

Project created using:

* Microsoft Excel with VBA: Office 365 *(Compatible with MS Excel 2010)*

## Source data

Data starts from 7th row. First six rows are irrelevant
**Column names** and **example row** below.

|NO.|CASE_NUMBER|SOURCE_STATUS|OUTPUT_STATUS|DATE_OF_CHANGE|LOGIN|
|---|---|---|---|---|---|
|987|ZA190502012345 |ZGŁOSZENIE|PROPOZYCJA|2019-05-03   20:55:37|jkowalski|

| COLUMN | *TYPE* |
|---|---|
| NO.| *LONG* |
| CASE_NUMBER | *STRING (15 CHARS)* |
| SOURCE_STATUS | *STRING* |
| OUTPUT_STATUS | *STRING* |
| DATE_OF_CHANGE | *DATE (yyyy-MM-dd HH:mm:ss)* |
| LOGIN | *STRING* |

## Output data

Output data are in the new worksheet with the name based on the date when the macro was started (yyyyMMdd_HHmmss).

**Column names** and **example row**

|LOGIN|CASE_NUMBER_1|CASE_NUMBER_2|CASE_NUMBER_3|
|---|---|---|---|
|jkowalski|ZA190502012345|ZA190504001527|ZA190427002137

## Assumptions

Requirements for the output data:

1. Cases cannot have the same date (when source_status = "PROPOZYCJA" and output_status = "ZREALIZOWANA") in the source data (ignore time).
2. Number of cases is equal to number of different dates (when source_status = "PROPOZYCJA" and output_status = "ZREALIZOWANA"). The maximum number is 3.
3. If exists, at least one case for each login must have row with this case and source_status = "ZGŁOSZENIE" and output status = "ZGŁOSZENIE".
4. Dates (when source_status = "PROPOZYCJA" and output_status = "ZREALIZOWANA") in the source data (ignore time) must be random and divided into three, disjoint sectors (begin days, middle days, last days).

## Instructions for use

To start macro you have to put both files into module in your file. In the source sheet you have to run macro "generator_zzpb05". It might take several minutes. After finishing the macro operation you have to run "preparing-report" macro while being in the output sheet. It should not take more than 10-15 seconds. After these actions you have two sheets: source and output.

## Mode of action

### For all macros

1. Before the macro starts, screen updating is turn off and calculation is set to manual calculation
2. All variables are declared at the start of the macro. Where it is not known size of the variable, the chosen variable was the bigger one (e.g. *Long* instead of *Integer*)
3. After the macro finishes, screen updating is turn on and caluclation is set to automatic calculation.

### Macro "generator_zzpb05"

1. Sets name of using sheets basing on current date and time.
2. Checks the number of rows in the source sheet.
3. Filters right data in source sheet basing on assumptions and copies them. Due to the impossibility of using 3 conditions, filter conatins only two of them.
    - SOURCE_STATUS = "ZGŁOSZENIE"
    - OUTPUT_STATUS = "PROPOZYCJA"
    - LOGIN <> "PaK Zdrowie"
    - LOGIN <> "portal świadczeniodawcy"
4. Creates new sheet and pastes only filtered rows (without column "NO.").
5. Deletes unnecessary columns ("SOURCE_STATUS" and "OUTPUT_STATUS").
6. Changes date format in the proper column ("DATE_OF_CHANGE"). Date format after change: yyyy-MM-dd.
7. Checks the number of rows in this new sheet.
8. Filters data using third condition.
    - LOGIN = "ass-system"
9. Deletes filtered rows (with header row).
10. Inserts new row and names new headers.
11. Removes duplicates rows.
12. Updates variable with number of rows in this sheet.
13. Filters right data in source sheet basing on assumptions and copies them. Due to the impossibility of using 3 conditions, filter conatins only two of them.
    - SOURCE_STATUS = "PROPOZYCJA"
    - OUTPUT_STATUS = "ZREALIZOWANA"
    - LOGIN <> "PaK Zdrowie"
    - LOGIN <> "portal świadczeniodawcy"
14. Creates new sheet and pastes only filtered rows (without column "NO.").
15. Deletes unnecessary columns ("SOURCE_STATUS" and "OUTPUT_STATUS").
16. Changes date format in the proper column ("DATE_OF_CHANGE"). Date format after change: yyyy-MM-dd.
17. Checks the number of rows in this new sheet.
18. Filters data using third condition.
    - LOGIN = "ass-system"
19. Deletes filtered rows (with header row).
20. Inserts new row and names new headers.
21. Removes duplicates rows.
22. Updates variable with number of rows in this sheet.
23. Fills new column ("RANDOM_NUMBER") with random numbers using *Rnd* function.
24. Sorts data basing on the new column ("RANDOM NUMBER").
25. Adds columns in two created sheets (one column in each sheet). Fills them with uniqe values composed of "CASE_NUMBER" and "LOGIN"
26. Checks in one sheet if for the case number exists the same login in the other sheet.
    - if exists then values in this column replaces by "YES"
    - else then values in this column replaces by "NO"
27. In the same sheet copy only two columns and pastes to the first empty column. Removes duplicates rows.
28. Checks the number of rows in this columns.
29. Sorts this data by login, then date.
30. Adds new sheet. It is the output sheet.
31. Copies logins and pastes them to the output sheet.
32. Removes duplicates in the output sheet.
33. Names new columns in the output sheet ("CASE_NUMBER_1", "CASE_NUMBER_2", "CASE_NUMBER_3", "ZGL-PRO-ZRE", "CASES_TOTAL_NUMBER", "DIFFERENT_DATES").
34. Inits variables. They will be used in loops.
35. Using loops fills columns "CASES_TOTAL_NUMBER", "DIFFERENT_DATES" and all dates.
36. Checks the number of rows in this columns.
37. Uses loops for all logins to fill cases columns.
38. If "DIFFERENT_DATES" for the login is greater or equal 3 then:
    1. Splits dates into three sectors:
        - First sector is from the first date to the number of different dates divided by 3 (integer division).
        - Second sector is from the number of different dates divided by 3 (integer division) plus one to the number of different dates divided by 3 (integer division) multiplied by 2.
        - In the third sector there are other dates. 
    2. Draws numbers for each sector.
    3. For each case finds the first value in the previous sheet (name ended by "pro-zre") where date is equal to the number from the drawn number (in the relevant sector).
    4. If, for founded case, the column from point 26 is not "YES", then interior color index of founded case changes into red (3).
    5. If all cases are red and the number of different dates is greater than 0, then finds case where the column from point 26 is equal to "YES", then fills this case to the right sector.
39. If "DIFFERENT_DATES" for the login is equal to 2:
    1. For both cases finds the first value in the previous sheet (name ended by "pro-zre") where date is equal to the first date for this login.
    2. If, for founded case, the column from point 26 is not "YES", then interior color index of founded case changes into red (3).
    3. If both cases are red and the number of different dates is greater than 0, then finds case where the column from point 26 is equal to "YES", then fills this case to the right column.
40. If "DIFFERENT_DATES" for the login is equal to 1:
    1. If the number of different dates is greater than 0, then finds case where the column from point 26 is equal to "YES".
    2. In other case, finds any case for this login and changes interior color index of founded case changes into red (3).
41. Removes unnecessary sheets.
42. Turns off autofilter in source sheet.

### Macro "preparing-report"

1. Checks the number of rows in the sheet.
2. Sets range of the output data.
3. Changes all interior color index to blank.
4. Makes borders.
5. Fits columns.
6. Deletes unnecessary columns.
7. Hides gridlines.
8. Select A1.

## Potential issues

* Other date system settings.
* Other versions of MS Excel.