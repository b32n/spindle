// Functions that keep an empty argument in place (`F(a,,b)`) so a later argument stays in its position.
//
// The parser drops empty arguments for every other function. Excel gives an omitted argument its default, and
// for these functions a default in the middle of the list is common (`INDEX(a,,2)`,
// `NUMBERVALUE(a,,",")`), so the placeholder evaluates to `undefined`, as an argument left off the end does.
// A leading `_xlfn.` (how Excel writes newer functions into files) is ignored when matching.
// Kept apart from the registry so the parser can import it without loading every function.

export const omittableArgFunctions: ReadonlySet<string> = new Set([
  'INDEX',
  'NUMBERVALUE',
  'FIXED',
  'WORKDAY.INTL',
  'NETWORKDAYS.INTL',
  'XIRR',
]);
