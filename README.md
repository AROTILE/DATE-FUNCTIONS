# DATE-FUNCTIONS
DATE FUNCTIONS

Overview

This document provides an overview of common date functions used in Microsoft Excel, including calculations for holidays and leave periods. These functions enable you to work with dates and times in your spreadsheets, performing tasks such as formatting, calculating, and manipulating dates.

Common Date Functions

1. Date Formatting
- TODAY(): Returns the current date.
- NOW(): Returns the current date and time.
- TEXT(A1, "format"): Formats a date in cell A1 according to a specified format string.

2. Date Calculations
- DATEDIF(A1, B1, "unit"): Calculates the difference between two dates in a specified unit (days, months, years).
- EDATE(A1, months): Returns a date that is a specified number of months before or after a date.
- EOMONTH(A1, months): Returns the last day of the month that is a specified number of months before or after a date.
- NETWORKDAYS(A1, B1, holidays): Calculates the number of workdays between two dates, excluding weekends and holidays.
- WORKDAY(A1, days, holidays): Returns a date that is a specified number of workdays before or after a date, excluding weekends and holidays.

3. Date Extraction
- YEAR(A1): Returns the year of a date in cell A1.
- MONTH(A1): Returns the month of a date in cell A1.
- DAY(A1): Returns the day of a date in cell A1.

4. Date Creation
- DATE(year, month, day): Creates a date from separate year, month, and day values.

5. Holidays and Leave Periods
- To calculate holidays and leave periods, you can use the NETWORKDAYS and WORKDAY functions, specifying a range of holiday dates.
- Example: =NETWORKDAYS(A1, B1, holidays!A1:A10), where holidays!A1:A10 is a range of holiday dates.

Examples

- Getting the current date: =TODAY()
- Calculating the number of days between two dates: =DATEDIF(A1, B1, "D")
- Calculating the number of workdays between two dates, excluding weekends and holidays: =NETWORKDAYS(A1, B1, holidays!A1:A10)
- Calculating a leave period: =WORKDAY(A1, 10, holidays!A1:A10), where A1 is the start date and 10 is the number of workdays.

Tips and Best Practices

- Use the TODAY() function to insert the current date in a cell.
- Use the DATEDIF() function to calculate age or duration between two dates.
- Use the NETWORKDAYS() and WORKDAY() functions to calculate workdays and leave periods, excluding weekends and holidays.
- Keep a separate list of holiday dates to reference in your calculations.

Conclusion

Excel's date functions make it easy to work with dates and times in your spreadsheets, including calculations for holidays and leave periods. By understanding and utilizing these functions, you can efficiently perform date-related tasks and ensure accurate calculations and formatting.

Author
Arotile Oluwaseun Bamitale
