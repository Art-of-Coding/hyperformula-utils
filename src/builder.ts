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
    let highestRow = 0
    let highestCol = 0
    for (const address of this.#map.keys()) {
      if (address.col > highestCol) {
        highestCol = address.col
      }
      if (address.row > highestRow) {
        highestRow = address.row
      }
    }

    for (let col = 0; col <= highestCol; col++) {
      sheet.push(new Array(highestRow).fill(null))
      const r = sheet[col]
      for (let row = 0; row <= highestRow; row++) {
        r.push(null)
      }
    }

    for (const address of this.#map.keys()) {
      sheet[address.col][address.row] = this.#map.get(address)
    }
    return sheet
  }
}