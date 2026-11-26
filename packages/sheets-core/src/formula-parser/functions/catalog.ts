// Metadata catalog for the built-in formula functions.
//
// The function registry stores only implementations; this catalog adds the
// human-facing metadata — signature and description — that powers formula
// autocomplete and parameter help. Every registered function has an entry
// (enforced by catalog.test.ts).

export type FunctionCategory =
  | 'math'
  | 'logical'
  | 'text'
  | 'lookup'
  | 'stats'
  | 'date'
  | 'information'
  | 'financial'
  | 'reference'
  | 'array';

export interface FunctionArg {
  name: string;
  optional?: boolean;
  /** Repeats — at most one per function, and it must be last. */
  variadic?: boolean;
}

export interface FunctionDoc {
  name: string; // UPPERCASE — matches the registry key
  category: FunctionCategory;
  args: FunctionArg[];
  description: string;
}

const req = (name: string): FunctionArg => ({ name });
const opt = (name: string): FunctionArg => ({ name, optional: true });
const rest = (name: string): FunctionArg => ({ name, optional: true, variadic: true });

const mathDocs: FunctionDoc[] = [
  { name: 'SUM', category: 'math', args: [req('number1'), rest('number2')], description: 'Adds all the numbers given as arguments.' },
  { name: 'AVERAGE', category: 'math', args: [req('number1'), rest('number2')], description: 'Returns the arithmetic mean of the arguments.' },
  { name: 'COUNT', category: 'math', args: [req('value1'), rest('value2')], description: 'Counts how many of the arguments are numbers.' },
  { name: 'COUNTA', category: 'math', args: [req('value1'), rest('value2')], description: 'Counts how many of the arguments are non-empty.' },
  { name: 'COUNTBLANK', category: 'math', args: [req('range')], description: 'Counts the empty cells in a range.' },
  { name: 'MAX', category: 'math', args: [req('number1'), rest('number2')], description: 'Returns the largest of the arguments.' },
  { name: 'MIN', category: 'math', args: [req('number1'), rest('number2')], description: 'Returns the smallest of the arguments.' },
  { name: 'PRODUCT', category: 'math', args: [req('number1'), rest('number2')], description: 'Multiplies all the numbers given as arguments.' },
  { name: 'SUMPRODUCT', category: 'math', args: [req('array1'), rest('array2')], description: 'Sums the products of corresponding array entries.' },
  { name: 'SUMIF', category: 'math', args: [req('range'), req('criterion'), opt('sum_range')], description: 'Adds the cells that meet a criterion.' },
  { name: 'SUMIFS', category: 'math', args: [req('sum_range'), req('criteria_range1'), req('criterion1'), rest('criteria_range, criterion')], description: 'Adds the cells that meet multiple criteria.' },
  { name: 'COUNTIF', category: 'math', args: [req('range'), req('criterion')], description: 'Counts the cells that meet a criterion.' },
  { name: 'COUNTIFS', category: 'math', args: [req('criteria_range1'), req('criterion1'), rest('criteria_range, criterion')], description: 'Counts the cells that meet multiple criteria.' },
  { name: 'AVERAGEIF', category: 'math', args: [req('range'), req('criterion'), opt('average_range')], description: 'Averages the cells that meet a criterion.' },
  { name: 'AVERAGEIFS', category: 'math', args: [req('average_range'), req('criteria_range1'), req('criterion1'), rest('criteria_range, criterion')], description: 'Averages the cells that meet multiple criteria.' },
  { name: 'SUBTOTAL', category: 'math', args: [req('function_num'), req('ref1'), rest('ref2')], description: 'Returns a subtotal (sum, average, count, …) chosen by a function number.' },
  { name: 'ABS', category: 'math', args: [req('number')], description: 'Returns the absolute value of a number.' },
  { name: 'SIGN', category: 'math', args: [req('number')], description: 'Returns the sign of a number: -1, 0 or 1.' },
  { name: 'INT', category: 'math', args: [req('number')], description: 'Rounds a number down to the nearest integer.' },
  { name: 'TRUNC', category: 'math', args: [req('number'), opt('digits')], description: 'Truncates a number to a number of digits.' },
  { name: 'ROUND', category: 'math', args: [req('number'), opt('digits')], description: 'Rounds a number to a number of digits.' },
  { name: 'ROUNDUP', category: 'math', args: [req('number'), opt('digits')], description: 'Rounds a number away from zero.' },
  { name: 'ROUNDDOWN', category: 'math', args: [req('number'), opt('digits')], description: 'Rounds a number toward zero.' },
  { name: 'MROUND', category: 'math', args: [req('number'), req('multiple')], description: 'Rounds a number to the nearest multiple.' },
  { name: 'CEILING', category: 'math', args: [req('number'), opt('significance')], description: 'Rounds a number up to the nearest multiple of significance.' },
  { name: 'FLOOR', category: 'math', args: [req('number'), opt('significance')], description: 'Rounds a number down to the nearest multiple of significance.' },
  { name: 'MOD', category: 'math', args: [req('number'), req('divisor')], description: 'Returns the remainder after division.' },
  { name: 'POWER', category: 'math', args: [req('base'), req('exponent')], description: 'Raises a number to a power.' },
  { name: 'SQRT', category: 'math', args: [req('number')], description: 'Returns the positive square root of a number.' },
  { name: 'EXP', category: 'math', args: [req('exponent')], description: 'Returns e raised to the given power.' },
  { name: 'LN', category: 'math', args: [req('number')], description: 'Returns the natural logarithm of a number.' },
  { name: 'LOG', category: 'math', args: [req('number'), opt('base')], description: 'Returns the logarithm of a number to a given base.' },
  { name: 'LOG10', category: 'math', args: [req('number')], description: 'Returns the base-10 logarithm of a number.' },
  { name: 'GCD', category: 'math', args: [req('number1'), rest('number2')], description: 'Returns the greatest common divisor.' },
  { name: 'LCM', category: 'math', args: [req('number1'), rest('number2')], description: 'Returns the least common multiple.' },
  { name: 'RAND', category: 'math', args: [], description: 'Returns a random number between 0 and 1.' },
  { name: 'RANDBETWEEN', category: 'math', args: [req('low'), req('high')], description: 'Returns a random integer between two values.' },
  { name: 'SIN', category: 'math', args: [req('number')], description: 'Returns the sine of an angle in radians.' },
  { name: 'COS', category: 'math', args: [req('number')], description: 'Returns the cosine of an angle in radians.' },
  { name: 'TAN', category: 'math', args: [req('number')], description: 'Returns the tangent of an angle in radians.' },
  { name: 'ASIN', category: 'math', args: [req('number')], description: 'Returns the arcsine of a number, in radians (-1 to 1).' },
  { name: 'ACOS', category: 'math', args: [req('number')], description: 'Returns the arccosine of a number, in radians (-1 to 1).' },
  { name: 'ATAN', category: 'math', args: [req('number')], description: 'Returns the arctangent of a number, in radians.' },
  { name: 'ATAN2', category: 'math', args: [req('x_num'), req('y_num')], description: 'Returns the arctangent of the point (x, y), in radians.' },
  { name: 'SINH', category: 'math', args: [req('number')], description: 'Returns the hyperbolic sine of a number.' },
  { name: 'COSH', category: 'math', args: [req('number')], description: 'Returns the hyperbolic cosine of a number.' },
  { name: 'TANH', category: 'math', args: [req('number')], description: 'Returns the hyperbolic tangent of a number.' },
  { name: 'ASINH', category: 'math', args: [req('number')], description: 'Returns the inverse hyperbolic sine of a number.' },
  { name: 'ACOSH', category: 'math', args: [req('number')], description: 'Returns the inverse hyperbolic cosine of a number (at least 1).' },
  { name: 'ATANH', category: 'math', args: [req('number')], description: 'Returns the inverse hyperbolic tangent of a number (between -1 and 1).' },
  { name: 'PI', category: 'math', args: [], description: 'Returns the value of pi.' },
  { name: 'RADIANS', category: 'math', args: [req('angle')], description: 'Converts degrees to radians.' },
  { name: 'DEGREES', category: 'math', args: [req('angle')], description: 'Converts radians to degrees.' },
  { name: 'FACT', category: 'math', args: [req('number')], description: 'Returns the factorial of a number.' },
  { name: 'FACTDOUBLE', category: 'math', args: [req('number')], description: 'Returns the double factorial of a number.' },
  { name: 'COMBIN', category: 'math', args: [req('number'), req('number_chosen')], description: 'Returns the number of combinations for a number of items.' },
  { name: 'PERMUT', category: 'math', args: [req('number'), req('number_chosen')], description: 'Returns the number of permutations for a number of items.' },
  { name: 'SUMSQ', category: 'math', args: [req('number1'), rest('number2')], description: 'Returns the sum of the squares of the arguments.' },
  { name: 'EVEN', category: 'math', args: [req('number')], description: 'Rounds a number away from zero to the nearest even integer.' },
  { name: 'ODD', category: 'math', args: [req('number')], description: 'Rounds a number away from zero to the nearest odd integer.' },
  { name: 'ISEVEN', category: 'math', args: [req('number')], description: 'Returns TRUE when a number is even.' },
  { name: 'ISODD', category: 'math', args: [req('number')], description: 'Returns TRUE when a number is odd.' },
  { name: 'CEILING.MATH', category: 'math', args: [req('number'), opt('significance'), opt('mode')], description: 'Rounds a number up to the nearest multiple of significance.' },
  { name: 'FLOOR.MATH', category: 'math', args: [req('number'), opt('significance'), opt('mode')], description: 'Rounds a number down to the nearest multiple of significance.' },
  { name: 'MAXIFS', category: 'math', args: [req('max_range'), req('criteria_range1'), req('criterion1'), rest('criteria_range, criterion')], description: 'Returns the largest value among the cells that meet multiple criteria.' },
  { name: 'MINIFS', category: 'math', args: [req('min_range'), req('criteria_range1'), req('criterion1'), rest('criteria_range, criterion')], description: 'Returns the smallest value among the cells that meet multiple criteria.' },
  { name: 'AGGREGATE', category: 'math', args: [req('function_num'), req('options'), req('ref1'), rest('ref2 or k')], description: 'Applies one of 19 functions to a list, optionally ignoring errors.' },
];

const logicalDocs: FunctionDoc[] = [
  { name: 'AND', category: 'logical', args: [req('logical1'), rest('logical2')], description: 'Returns TRUE if all arguments are TRUE.' },
  { name: 'OR', category: 'logical', args: [req('logical1'), rest('logical2')], description: 'Returns TRUE if any argument is TRUE.' },
  { name: 'XOR', category: 'logical', args: [req('logical1'), rest('logical2')], description: 'Returns TRUE if an odd number of arguments are TRUE.' },
  { name: 'NOT', category: 'logical', args: [req('logical')], description: 'Reverses the logic of its argument.' },
  { name: 'TRUE', category: 'logical', args: [], description: 'Returns the logical value TRUE.' },
  { name: 'FALSE', category: 'logical', args: [], description: 'Returns the logical value FALSE.' },
  { name: 'IF', category: 'logical', args: [req('condition'), req('value_if_true'), opt('value_if_false')], description: 'Returns one value if a condition is true and another if it is false.' },
  { name: 'IFERROR', category: 'logical', args: [req('value'), req('value_if_error')], description: 'Returns a fallback value if the first argument is an error.' },
  { name: 'IFNA', category: 'logical', args: [req('value'), req('value_if_na')], description: 'Returns a fallback value if the first argument is #N/A.' },
  { name: 'IFS', category: 'logical', args: [req('condition1'), req('value1'), rest('condition, value')], description: 'Returns the value for the first condition that is true.' },
  { name: 'SWITCH', category: 'logical', args: [req('expression'), req('case1'), req('value1'), rest('case, value / default')], description: 'Compares an expression against cases and returns the first match.' },
];

const textDocs: FunctionDoc[] = [
  { name: 'CONCAT', category: 'text', args: [req('text1'), rest('text2')], description: 'Joins several text values into one.' },
  { name: 'CONCATENATE', category: 'text', args: [req('text1'), rest('text2')], description: 'Joins several text values into one.' },
  { name: 'TEXTJOIN', category: 'text', args: [req('delimiter'), req('ignore_empty'), req('text1'), rest('text2')], description: 'Joins text values together with a delimiter.' },
  { name: 'LEFT', category: 'text', args: [req('text'), opt('num_chars')], description: 'Returns the first characters of a text value.' },
  { name: 'RIGHT', category: 'text', args: [req('text'), opt('num_chars')], description: 'Returns the last characters of a text value.' },
  { name: 'MID', category: 'text', args: [req('text'), req('start'), req('num_chars')], description: 'Returns characters from the middle of a text value.' },
  { name: 'LEN', category: 'text', args: [req('text')], description: 'Returns the number of characters in a text value.' },
  { name: 'FIND', category: 'text', args: [req('find_text'), req('within_text'), opt('start')], description: 'Finds one text value within another (case-sensitive).' },
  { name: 'SEARCH', category: 'text', args: [req('find_text'), req('within_text'), opt('start')], description: 'Finds one text value within another (case-insensitive, wildcards).' },
  { name: 'SUBSTITUTE', category: 'text', args: [req('text'), req('old_text'), req('new_text'), opt('instance')], description: 'Replaces occurrences of old text with new text.' },
  { name: 'REPLACE', category: 'text', args: [req('old_text'), req('start'), req('num_chars'), req('new_text')], description: 'Replaces part of a text value by position.' },
  { name: 'UPPER', category: 'text', args: [req('text')], description: 'Converts text to uppercase.' },
  { name: 'LOWER', category: 'text', args: [req('text')], description: 'Converts text to lowercase.' },
  { name: 'PROPER', category: 'text', args: [req('text')], description: 'Capitalizes the first letter of each word.' },
  { name: 'TRIM', category: 'text', args: [req('text')], description: 'Removes extra spaces from text.' },
  { name: 'TEXT', category: 'text', args: [req('value'), req('format')], description: 'Formats a number as text using a format string.' },
  { name: 'VALUE', category: 'text', args: [req('text')], description: 'Converts a text value to a number.' },
  { name: 'REPT', category: 'text', args: [req('text'), req('count')], description: 'Repeats text a given number of times.' },
  { name: 'CHAR', category: 'text', args: [req('number')], description: 'Returns the character for a character code.' },
  { name: 'CODE', category: 'text', args: [req('text')], description: 'Returns the character code of the first character.' },
  { name: 'EXACT', category: 'text', args: [req('text1'), req('text2')], description: 'Returns TRUE if two text values are exactly equal.' },
];

const lookupDocs: FunctionDoc[] = [
  { name: 'VLOOKUP', category: 'lookup', args: [req('lookup_value'), req('table'), req('col_index'), opt('is_sorted')], description: "Searches a range's first column and returns a value from a column." },
  { name: 'HLOOKUP', category: 'lookup', args: [req('lookup_value'), req('table'), req('row_index'), opt('is_sorted')], description: "Searches a range's first row and returns a value from a row." },
  { name: 'XLOOKUP', category: 'lookup', args: [req('lookup_value'), req('lookup_array'), req('return_array'), opt('if_not_found'), opt('match_mode'), opt('search_mode')], description: 'Searches a range and returns the matching item from another range.' },
  { name: 'MATCH', category: 'lookup', args: [req('lookup_value'), req('lookup_array'), opt('match_type')], description: 'Returns the position of a value within a range.' },
  { name: 'XMATCH', category: 'lookup', args: [req('lookup_value'), req('lookup_array'), opt('match_mode'), opt('search_mode')], description: 'Returns the position of a value within a range.' },
  { name: 'INDEX', category: 'lookup', args: [req('array'), opt('row'), opt('column')], description: 'Returns the value at a position in a range.' },
  { name: 'LOOKUP', category: 'lookup', args: [req('lookup_value'), req('lookup_vector'), opt('result_vector')], description: 'Looks up a value in a vector and returns a corresponding result.' },
  { name: 'CHOOSE', category: 'lookup', args: [req('index'), req('value1'), rest('value2')], description: 'Returns one of a list of values, chosen by index.' },
  { name: 'ROWS', category: 'lookup', args: [req('range')], description: 'Returns the number of rows in a range.' },
  { name: 'COLUMNS', category: 'lookup', args: [req('range')], description: 'Returns the number of columns in a range.' },
  { name: 'ADDRESS', category: 'lookup', args: [req('row'), req('column'), opt('abs_num'), opt('a1'), opt('sheet')], description: 'Builds a cell-reference string from row and column numbers.' },
  { name: 'HYPERLINK', category: 'lookup', args: [req('url'), opt('label')], description: 'Creates a hyperlink with optional display text.' },
];

const statsDocs: FunctionDoc[] = [
  { name: 'MEDIAN', category: 'stats', args: [req('number1'), rest('number2')], description: 'Returns the median of the arguments.' },
  { name: 'MODE', category: 'stats', args: [req('number1'), rest('number2')], description: 'Returns the most frequently occurring value.' },
  { name: 'VAR', category: 'stats', args: [req('number1'), rest('number2')], description: 'Estimates variance from a sample.' },
  { name: 'VARP', category: 'stats', args: [req('number1'), rest('number2')], description: 'Calculates variance of an entire population.' },
  { name: 'STDEV', category: 'stats', args: [req('number1'), rest('number2')], description: 'Estimates standard deviation from a sample.' },
  { name: 'STDEVP', category: 'stats', args: [req('number1'), rest('number2')], description: 'Calculates standard deviation of an entire population.' },
  { name: 'PERCENTILE', category: 'stats', args: [req('array'), req('k')], description: 'Returns the k-th percentile of values (k between 0 and 1).' },
  { name: 'QUARTILE', category: 'stats', args: [req('array'), req('quart')], description: 'Returns a quartile (0–4) of a data set.' },
  { name: 'RANK', category: 'stats', args: [req('number'), req('ref'), opt('order')], description: 'Returns the rank of a number within a list.' },
  { name: 'LARGE', category: 'stats', args: [req('array'), req('k')], description: 'Returns the k-th largest value in a data set.' },
  { name: 'SMALL', category: 'stats', args: [req('array'), req('k')], description: 'Returns the k-th smallest value in a data set.' },
  { name: 'CORREL', category: 'stats', args: [req('array1'), req('array2')], description: 'Returns the correlation coefficient of two ranges.' },
  { name: 'COVAR', category: 'stats', args: [req('array1'), req('array2')], description: 'Returns the population covariance of two ranges.' },  { name: 'STDEV.S', category: 'stats', args: [req('number1'), rest('number2')], description: 'Estimates standard deviation from a sample.' },
  { name: 'STDEV.P', category: 'stats', args: [req('number1'), rest('number2')], description: 'Calculates standard deviation of an entire population.' },
  { name: 'VAR.S', category: 'stats', args: [req('number1'), rest('number2')], description: 'Estimates variance from a sample.' },
  { name: 'VAR.P', category: 'stats', args: [req('number1'), rest('number2')], description: 'Calculates variance of an entire population.' },
  { name: 'PERCENTILE.INC', category: 'stats', args: [req('array'), req('k')], description: 'Returns the k-th percentile of values (k from 0 to 1, inclusive).' },
  { name: 'PERCENTILE.EXC', category: 'stats', args: [req('array'), req('k')], description: 'Returns the k-th percentile of values (k strictly between 0 and 1).' },
  { name: 'QUARTILE.INC', category: 'stats', args: [req('array'), req('quart')], description: 'Returns a quartile (0-4) of a data set, inclusive.' },
  { name: 'QUARTILE.EXC', category: 'stats', args: [req('array'), req('quart')], description: 'Returns a quartile (1-3) of a data set, exclusive.' },
  { name: 'RANK.EQ', category: 'stats', args: [req('number'), req('ref'), opt('order')], description: 'Returns the rank of a number within a list; ties share the top rank.' },
  { name: 'MODE.SNGL', category: 'stats', args: [req('number1'), rest('number2')], description: 'Returns the most frequently occurring value.' },
  { name: 'SLOPE', category: 'stats', args: [req('known_ys'), req('known_xs')], description: 'Returns the slope of the linear regression line.' },
  { name: 'INTERCEPT', category: 'stats', args: [req('known_ys'), req('known_xs')], description: 'Returns the intercept of the linear regression line.' },
  { name: 'RSQ', category: 'stats', args: [req('known_ys'), req('known_xs')], description: 'Returns the square of the Pearson correlation coefficient.' },
  { name: 'FORECAST', category: 'stats', args: [req('x'), req('known_ys'), req('known_xs')], description: 'Predicts a value along a linear trend.' },
  { name: 'FORECAST.LINEAR', category: 'stats', args: [req('x'), req('known_ys'), req('known_xs')], description: 'Predicts a value along a linear trend.' },
  { name: 'GEOMEAN', category: 'stats', args: [req('number1'), rest('number2')], description: 'Returns the geometric mean of positive numbers.' },
  { name: 'FREQUENCY', category: 'stats', args: [req('data_array'), req('bins_array')], description: 'Counts values per interval, as a column (spills).' },
];

const dateDocs: FunctionDoc[] = [
  { name: 'TODAY', category: 'date', args: [], description: 'Returns the current date.' },
  { name: 'NOW', category: 'date', args: [], description: 'Returns the current date and time.' },
  { name: 'DATE', category: 'date', args: [req('year'), req('month'), req('day')], description: 'Builds a date from year, month and day.' },
  { name: 'TIME', category: 'date', args: [req('hour'), req('minute'), req('second')], description: 'Builds a time from hour, minute and second.' },
  { name: 'YEAR', category: 'date', args: [req('date')], description: 'Returns the year of a date.' },
  { name: 'MONTH', category: 'date', args: [req('date')], description: 'Returns the month of a date.' },
  { name: 'DAY', category: 'date', args: [req('date')], description: 'Returns the day of the month of a date.' },
  { name: 'HOUR', category: 'date', args: [req('time')], description: 'Returns the hour of a time.' },
  { name: 'MINUTE', category: 'date', args: [req('time')], description: 'Returns the minute of a time.' },
  { name: 'SECOND', category: 'date', args: [req('time')], description: 'Returns the second of a time.' },
  { name: 'WEEKDAY', category: 'date', args: [req('date'), opt('type')], description: 'Returns the day of the week of a date.' },
  { name: 'WEEKNUM', category: 'date', args: [req('date'), opt('type')], description: 'Returns the week number of a date.' },
  { name: 'EDATE', category: 'date', args: [req('start_date'), req('months')], description: 'Returns the date a number of months from a start date.' },
  { name: 'EOMONTH', category: 'date', args: [req('start_date'), req('months')], description: 'Returns the last day of the month a number of months away.' },
  { name: 'DATEDIF', category: 'date', args: [req('start_date'), req('end_date'), req('unit')], description: 'Returns the difference between two dates in a chosen unit.' },
  { name: 'NETWORKDAYS', category: 'date', args: [req('start_date'), req('end_date'), opt('holidays')], description: 'Counts the working days between two dates.' },
  { name: 'WORKDAY', category: 'date', args: [req('start_date'), req('days'), opt('holidays')], description: 'Returns a date a number of working days away.' },
  { name: 'DATEVALUE', category: 'date', args: [req('date_text')], description: 'Converts a date written as text to a date.' },
  { name: 'TIMEVALUE', category: 'date', args: [req('time_text')], description: 'Converts a time written as text to a time.' },
];

const informationDocs: FunctionDoc[] = [
  { name: 'ISBLANK', category: 'information', args: [req('value')], description: 'Returns TRUE if the value is empty.' },
  { name: 'ISNUMBER', category: 'information', args: [req('value')], description: 'Returns TRUE if the value is a number.' },
  { name: 'ISTEXT', category: 'information', args: [req('value')], description: 'Returns TRUE if the value is text.' },
  { name: 'ISLOGICAL', category: 'information', args: [req('value')], description: 'Returns TRUE if the value is a boolean.' },
  { name: 'ISERROR', category: 'information', args: [req('value')], description: 'Returns TRUE if the value is any error.' },
  { name: 'ISERR', category: 'information', args: [req('value')], description: 'Returns TRUE if the value is an error other than #N/A.' },
  { name: 'ISNA', category: 'information', args: [req('value')], description: 'Returns TRUE if the value is the #N/A error.' },
  { name: 'TYPE', category: 'information', args: [req('value')], description: 'Returns a number indicating the type of a value.' },
  { name: 'ERROR.TYPE', category: 'information', args: [req('value')], description: 'Returns a number identifying an error value.' },
  { name: 'NA', category: 'information', args: [], description: 'Returns the #N/A error value.' },
  { name: 'N', category: 'information', args: [req('value')], description: 'Converts a value to a number.' },
];

const financialDocs: FunctionDoc[] = [
  { name: 'PMT', category: 'financial', args: [req('rate'), req('nper'), req('pv'), opt('fv'), opt('type')], description: 'Returns the periodic payment for a loan or annuity.' },
  { name: 'FV', category: 'financial', args: [req('rate'), req('nper'), req('pmt'), opt('pv'), opt('type')], description: 'Returns the future value of an investment.' },
  { name: 'PV', category: 'financial', args: [req('rate'), req('nper'), req('pmt'), opt('fv'), opt('type')], description: 'Returns the present value of an investment.' },
  { name: 'NPER', category: 'financial', args: [req('rate'), req('pmt'), req('pv'), opt('fv'), opt('type')], description: 'Returns the number of periods for an investment.' },
  { name: 'RATE', category: 'financial', args: [req('nper'), req('pmt'), req('pv'), opt('fv'), opt('type'), opt('guess')], description: 'Returns the interest rate per period of an annuity.' },
  { name: 'NPV', category: 'financial', args: [req('rate'), req('value1'), rest('value2')], description: 'Returns the net present value of a series of cash flows.' },
  { name: 'IRR', category: 'financial', args: [req('values'), opt('guess')], description: 'Returns the internal rate of return for a series of cash flows.' },
  { name: 'IPMT', category: 'financial', args: [req('rate'), req('per'), req('nper'), req('pv'), opt('fv'), opt('type')], description: 'Returns the interest portion of a given payment.' },
  { name: 'PPMT', category: 'financial', args: [req('rate'), req('per'), req('nper'), req('pv'), opt('fv'), opt('type')], description: 'Returns the principal portion of a given payment.' },
];

const referenceDocs: FunctionDoc[] = [
  { name: 'ROW', category: 'reference', args: [opt('reference')], description: 'Returns the row number of a reference.' },
  { name: 'COLUMN', category: 'reference', args: [opt('reference')], description: 'Returns the column number of a reference.' },
  { name: 'OFFSET', category: 'reference', args: [req('reference'), req('rows'), req('cols'), opt('height'), opt('width')], description: 'Returns a reference shifted from a starting reference.' },
  { name: 'INDIRECT', category: 'reference', args: [req('ref_text'), opt('a1')], description: 'Returns the reference specified by a text string.' },
  { name: 'ISREF', category: 'reference', args: [req('value')], description: 'Returns TRUE if the value is a reference.' },
  { name: 'ISFORMULA', category: 'reference', args: [req('reference')], description: 'Returns TRUE if the referenced cell contains a formula.' },
  { name: 'CELL', category: 'reference', args: [req('info_type'), opt('reference')], description: 'Returns information about a cell.' },
];

const arrayDocs: FunctionDoc[] = [
  { name: 'SEQUENCE', category: 'array', args: [req('rows'), opt('columns'), opt('start'), opt('step')], description: 'Generates a sequence of numbers in an array.' },
  { name: 'UNIQUE', category: 'array', args: [req('array'), opt('by_column'), opt('exactly_once')], description: 'Returns the unique rows of a range.' },
  { name: 'SORT', category: 'array', args: [req('array'), opt('sort_index'), opt('sort_order'), opt('by_column')], description: 'Sorts the rows of a range.' },
  { name: 'SORTBY', category: 'array', args: [req('array'), req('by_array1'), opt('sort_order1'), rest('by_array, sort_order')], description: 'Sorts a range by the values in one or more other ranges.' },
  { name: 'FILTER', category: 'array', args: [req('array'), req('include'), opt('if_empty')], description: 'Returns the rows of a range that meet a condition.' },
  { name: 'SPLIT', category: 'array', args: [req('text'), req('delimiter'), opt('split_by_each'), opt('remove_empty')], description: 'Splits text into an array around a delimiter.' },
];

const allDocs: FunctionDoc[] = [
  ...mathDocs,
  ...logicalDocs,
  ...textDocs,
  ...lookupDocs,
  ...statsDocs,
  ...dateDocs,
  ...informationDocs,
  ...financialDocs,
  ...referenceDocs,
  ...arrayDocs,
];

/** Metadata for every built-in function, keyed by uppercase name. */
export const functionCatalog: Record<string, FunctionDoc> = Object.fromEntries(
  allDocs.map((doc) => [doc.name, doc])
);

/** The catalog entry for a function name (case-insensitive), or undefined. */
export function getFunctionDoc(name: string): FunctionDoc | undefined {
  return functionCatalog[name.toUpperCase()];
}

/** Every catalog entry, sorted by name. */
export function getAllFunctionDocs(): FunctionDoc[] {
  return [...allDocs].sort((a, b) => a.name.localeCompare(b.name));
}

/** Catalog entries whose name starts with `prefix` (case-insensitive), sorted. */
export function searchFunctions(prefix: string): FunctionDoc[] {
  const needle = prefix.toUpperCase();
  return getAllFunctionDocs().filter((doc) => doc.name.startsWith(needle));
}
