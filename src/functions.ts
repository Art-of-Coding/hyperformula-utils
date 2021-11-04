import type { Sheet } from 'hyperformula'
import { customAlphabet } from 'nanoid'

const alphanum = ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z', '0', '1', '2', '3', '4', '5', '6', '7', '8', '9']
export const generateSheetName = customAlphabet(alphanum.join(''), 8)

/** Find dependencies for a sheet. */
export function findDependencies(sheet: Sheet) {
  const formulas = extractFormulas(sheet)
  return deduplicate(...formulas.map(formula => extractSheetNames(formula)))
}

/** Extract formulas from the sheet. */
export function extractFormulas(sheet: Sheet): string[] {
  const formulas: string[] = []

  for (let row = 0; row < sheet.length; row++) {
    for (let col = 0; col < sheet[row].length; col++) {
      const rawValue = sheet[row][col]
      if (typeof rawValue === 'string' && rawValue.startsWith('=')) {
        formulas.push(rawValue)
      }
    }
  }

  return formulas
}

/** Extract sheet names from the formula. */
export function extractSheetNames(formula: string): string[] {
  if (!formula.includes('!')) {
    return []
  }

  // NOTE: Replace with regex?
  const sheetNames: string[] = []
  let sheetName = ''
  for (let i = 0; i < formula.length; i++) {
    if (alphanum.includes(formula[i].toLowerCase())) {
      sheetName += formula[i]
    } else if (formula[i] === '!') {
      sheetNames.push(sheetName)
      sheetName = ''
    } else {
      sheetName = ''
    }
  }
  return sheetNames
}

// Deduplicate arrays and create a new array with unique elements
export function deduplicate<Type>(...arrays: Type[][]) {
  const result: Type[] = []
  for (const array of arrays) {
    for (const value of array) {
      if (!result.includes(value)) {
        result.push(value)
      }
    }
  }
  return result
}