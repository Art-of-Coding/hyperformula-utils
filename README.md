# HyperFormula Utilities

Utility functions for
[HyperFormula](https://handsontable.github.io/hyperformula/).

## Install

```
npm i @art-of-coding/hyperformula-utils
```

## Example

See the source code for more details.

```ts
import { HyperFormula, Sheet } from "hyperformula";
import { calculateFormula } from "@art-of-coding/hyperformula-utils";

const hfInstance = HyperFormula.buildEmpty({
  licenseKey: "gpl-v3",
});

async function resolve(name: string): Promise<Sheet> {
  if (name === "Sheet1") {
    return [["=Sheet2!B1*Sheet2!C1"]];
  } else if (name === "Sheet2") {
    return [["1", "2", "=A1*B1"]];
  }
  throw new Error(`Sheet "${name}" not found`);
}

console.log(await calculateFormula(hfInstance, "=Sheet1!A1*5.5", resolve));
```
