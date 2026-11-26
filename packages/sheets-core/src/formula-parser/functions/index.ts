// Built-in formula function registry.

import type { EagerFn, LazyFn, RefFn } from './helpers';
import { mathFunctions } from './math';
import { mathExtFunctions } from './math-ext';
import { logicalEagerFunctions, logicalLazyFunctions } from './logical';
import { textFunctions } from './text';
import { lookupFunctions } from './lookup';
import { statsFunctions } from './stats';
import { statsExtFunctions } from './stats-ext';
import { aggregateFunctions } from './aggregate';
import { dateFunctions } from './date';
import { dateExtFunctions } from './date-ext';
import { textExtFunctions } from './text-ext';
import { informationEagerFunctions, informationLazyFunctions } from './information';
import { financialFunctions } from './financial';
import { financialExtFunctions } from './financial-ext';
import { referenceFunctions } from './reference';
import { arrayFunctions } from './array';
import { arrayExtFunctions } from './array-ext';

export type { EagerFn, LazyFn, RefFn } from './helpers';
export { omittableArgFunctions } from './omittable';

/** Functions whose arguments are fully evaluated before the function runs. */
export const eagerFunctions: Record<string, EagerFn> = {
  ...mathFunctions,
  ...mathExtFunctions,
  ...aggregateFunctions,
  ...logicalEagerFunctions,
  ...textFunctions,
  ...textExtFunctions,
  ...lookupFunctions,
  ...statsFunctions,
  ...statsExtFunctions,
  ...dateFunctions,
  ...dateExtFunctions,
  ...informationEagerFunctions,
  ...financialFunctions,
  ...financialExtFunctions,
  ...arrayFunctions,
  ...arrayExtFunctions,
};

/** Functions that receive argument thunks so they can short-circuit. */
export const lazyFunctions: Record<string, LazyFn> = {
  ...logicalLazyFunctions,
  ...informationLazyFunctions,
};

/** Functions that work with an argument's reference rather than its value. */
export const refFunctions: Record<string, RefFn> = {
  ...referenceFunctions,
};

/**
 * Functions whose result can change without any input changing. A formula
 * containing one of these is recomputed on every recalculation pass.
 */
export const volatileFunctions: ReadonlySet<string> = new Set([
  'RAND',
  'RANDBETWEEN',
  'NOW',
  'TODAY',
  'OFFSET',
  'INDIRECT',
]);

/** Every registered built-in function name (eager, lazy and reference). */
export function getRegisteredFunctionNames(): Set<string> {
  return new Set([
    ...Object.keys(eagerFunctions),
    ...Object.keys(lazyFunctions),
    ...Object.keys(refFunctions),
  ]);
}
