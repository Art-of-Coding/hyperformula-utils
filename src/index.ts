import type { Sheet, Sheets } from 'hyperformula'
import type { HyperFormula } from 'hyperformula'
import { generateSheetName, findDependencies, extractSheetNames } from './functions'

/** Load a sheet and all its formula dependencies. */
export async function addSheet(
  hfInstance: HyperFormula,
  name: string,
  resolve: (name: string) => Promise<Sheet>,
) {
  // load all required sheets into memory before suspending evaluation
  // and adding all sheets at once
  const sheets: Sheets = {}
  const hasSheet = (sheetName: string) => !!sheets[sheetName] || hfInstance.getSheetId(sheetName) !== undefined
  const loadSheet = async (sheetName: string) => {
    if (!hasSheet(sheetName)) {
      const sheet = await resolve(sheetName)
      const missingDependencies = findDependencies(sheet).filter(dep => !hasSheet(dep))
      if (missingDependencies.length) {
        for (const missingDependency of missingDependencies) {
          await loadSheet(missingDependency)
        }
      }

      sheets[sheetName] = sheet
    }
  }

  await loadSheet(name)
  return addSheets(hfInstance, sheets)
}

/** Add multiple sheets and resuming evaluation. */
export function addSheets(hfInstance: HyperFormula, sheets: Sheets) {
  hfInstance.suspendEvaluation()
  for (const name of Object.keys(sheets)) {
    let sheetId = hfInstance.getSheetId(name)
    if (sheetId) {
      continue
    }
    const sheetName = hfInstance.addSheet(name)
    sheetId = hfInstance.getSheetId(sheetName)
    hfInstance.setSheetContent(sheetId!, sheets[name])
  }

  return hfInstance.resumeEvaluation()
}

/** Calculate a fire-and-forget formula. */
export async function calculateFormula(
  hfInstance: HyperFormula,
  formula: string,
  resolve: (name: string) => Promise<Sheet>
) {
  const dependencies = extractSheetNames(formula)
  if (dependencies.length) {
    hfInstance.suspendEvaluation()
    for (const dependency of dependencies) {
      await addSheet(hfInstance, dependency, resolve)
    }
    hfInstance.resumeEvaluation()
  }

  // if adding and removing the sheet induces performance issues,
  // create a single reusable sheet to act as formula context
  const sheetId = hfInstance.getSheetId(hfInstance.addSheet(generateSheetName()))
  const result = hfInstance.calculateFormula(formula, sheetId!)
  hfInstance.removeSheet(sheetId!)
  return result
}
