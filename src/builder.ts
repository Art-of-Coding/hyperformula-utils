import type { Sheet } from 'hyperformula'
import { simpleCellAddressFromString } from './functions'

export class Builder {
  #map: Map<{ row: number, col: number }, string> = new Map()

  /**
   * Set the contents of a cell.
   * @param cellAddress The call address in A1 notation
   * @param value The cell value
   * @returns The Builder instance
   */
  setCellContent(cellAddress: string, value: string) {
    this.#map.set(simpleCellAddressFromString(cellAddress), value)
    return this
  }

  /**
   * Remove the contents of a cell.
   * @param cellAddress The call address in A1 notation
   * @returns The Builder instance
   */
  removeCellContent(cellAddress: string) {
    const address = simpleCellAddressFromString(cellAddress)
    for (const mapped of this.#map.keys()) {
      if (mapped.row === address.row && mapped.col === address.col) {
        this.#map.delete(mapped)
        break
      }
    }
    return this
  }

  /**
   * Build the sheet based on the cells set.
   * @returns The sheet
   */
  build(): Sheet {
    const sheet: Sheet = []
    for (const address of this.#map.keys()) {
      if (!sheet[address.row]) {
        const remaining = sheet.length - address.row + 1
        if (remaining > 0) {
          for (let i = 0; i < remaining; i++) {
            sheet.push([''])
          }
        }
      }
      if (!sheet[address.row][address.col]) {
        const remaining = sheet[address.row].length - address.col + 1
        if (remaining > 0) {
          for (let i = 0; i < remaining; i++) {
            sheet.push([''])
          }
        }
      }
      sheet[address.col][address.row] = this.#map.get(address)
    }
    return sheet
  }
}