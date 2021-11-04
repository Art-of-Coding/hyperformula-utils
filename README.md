# HyperFormula Utilities

Utility functions for
[HyperFormula](https://handsontable.github.io/hyperformula/).

## Features

- Parse formulas and get all sheet names
- Load sheet and all sheets its formulas depend on
- On-demand loading of sheets for fire-and-forget formulas

### Limitations

- Sheet name must be alphanumberic (case-insansitive)

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
