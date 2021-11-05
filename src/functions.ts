import type { Sheet } from 'hyperformula'
import { customAlphabet } from 'nanoid'

const alphanum = ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z', '0', '1', '2', '3', '4', '5', '6', '7', '8', '9']
export const generateSheetName = customAlphabet(alphanum.join(''), 8)

/**
 * Find all dependencies for the given sheet.
 * @param sheet The sheet
 * @returns A list of dependencies
 */
export function findDependencies(sheet: Sheet) {
  const formulas = extractFormulas(sheet)
  return deduplicate(...formulas.map(formula => extractSheetNames(formula)))
}

/**
 * Extract all formulas from the given sheet.
 * @param sheet The sheet
 * @returns A list of formulas
 */
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

/**
 * Extract all sheet names from the given formula.
 * @param formula The formula
 * @returns A list of sheet names
 */
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

const alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'

/**
 * Parse a simple cell address from a string in A1 notation.
 * The address may not contain references to other sheets.
 * @param cellAddress The cell address in A1 notation
 * @returns The row and column numbers
 */
export function simpleCellAddressFromString(cellAddress: string): { row: number, col: number } {
  let rowString = ''
  let colString = ''
  for (let i = 0; i < cellAddress.length; i++) {
    if (alphabet.includes(cellAddress[i].toUpperCase())) {
      rowString += cellAddress[i]
    } else {
      colString += cellAddress[i]
    }
  }

  const offset = ((rowString.length - 1) * alphabet.length)
  const row = offset + alphabet.indexOf(rowString.substr(offset))
  return { row, col: parseInt(colString) - 1}
}

function deduplicate<Type>(...arrays: Type[][]) {
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