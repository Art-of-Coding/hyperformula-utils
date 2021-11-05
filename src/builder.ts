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
      if (sheet.length < address.row) {
        for (let i = 0; i < sheet.length; i++) {
          if (!sheet[i]) {
            sheet[i] = []
          }
        }
      }

      const col = sheet[address.col] ?? []
      if (col.length < address.col) {
        for (let i = 0; i < col.length; i++) {
          if (!col[i]) {
            col[i] = ''
          }
        }
      }
      sheet[address.col] = col
      sheet[address.row][address.col] = this.#map.get(address)
    }
    return sheet
  }
}