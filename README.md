# HyperFormula Utilities

Utility functions for
[HyperFormula](https://handsontable.github.io/hyperformula/).

## Features

- Parse formulas and get all sheet names
- Load sheet and all sheets its formulas depend on
- On-demand loading of sheets for fire-and-forget formulas

### Limitations

- Sheet name must be alphanumeric (case-insensitive)

## Install

```
npm i @art-of-coding/hyperformula-utils
```

## Example

The example below calculates a fire-and-forget formula which dynamically loads
the required dependencies on-demand. Dependent sheets are only loaded if no
sheet with the name exists.

See the source code for more details.

```ts
import { HyperFormula, Sheet } from "hyperformula";
import { calculateFormula } from "@art-of-coding/hyperformula-utils";

const hfInstance = HyperFormula.buildEmpty({
  licenseKey: "gpl-v3",
});

// Our custom resolve function
async function resolve(name: string): Promise<Sheet> {
  if (name === "Sheet1") {
    return [["=Sheet2!B1*Sheet2!C1"]];
  } else if (name === "Sheet2") {
    return [["1", "2", "=A1*B1"]];
  }
  throw new Error(`Sheet "${name}" not found`);
}

// Outputs '22'
console.log(await calculateFormula(hfInstance, "=Sheet1!A1*5.5", resolve));
```

## API Documentation

### `findDependencies(sheet: Sheet): string[]`

Find all dependencies for the given sheet. Will check all formulas for
references to other sheets and return their dependencies as well.

```ts
import { findDependencies } from "@art-of-coding/hyperformula-utils";

const sheet = [["1", "2", "=AnotherSheet!:A1*A1"]];
const dependencies = findDependencies(sheet);

/*
[
  'AnotherSheet'
]
*/
```

### `extractFormulas(sheet: Sheet): string[]`

Extract all formulas from the sheet. Formulas are any string that starts with
`=`.

```ts
import { extractFormulas } from "@art-of-coding/hyperformula-utils";

const sheet = [["1", "2", "=AnotherSheet!:A1*A1"]];
const formulas = extractFormulas(sheet);

/*
[
  '=AnotherSheet!:A1*A1'
]
*/
```

### `extractSheetNames(formula: string): string[]`

Extract all sheet names from a formula.

Sheet names must be alphanumeric.

```ts
import { extractSheetNames } from "@art-of-coding/hyperformula-utils";

const sheetNames = extractSheetNames("=AnotherSheet!:A1*A1");

/*
[
  '=AnotherSheet'
]
*/
```

### `simpleCellAddressFromString(cellAddress: string): { row: number, col: number }`

Parse string into a simple cell address. Does not include the `sheet` property.

The cell address should not contain absolute addresses (e.g. references to other
sheets).

```ts
import { extractSheetNames } from "@art-of-coding/hyperformula-utils";

const address = simpleCellAddressFromString("A1");

/*
{
  row: 0,
  col: 0
}
*/
```
