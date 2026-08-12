// Thanks to Nathan Wall
// https://stackoverflow.com/questions/12504042/what-is-a-method-that-can-be-used-to-increment-letters#12504061

export default class CellIdHelper {
  private _chars: string;
  private _nextId: number[];

  constructor(chars = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ') {
    this._chars = chars;
    this._nextId = [0];
  }

  start(index: number): this {
    this._nextId = [index];
    return this;
  }

  next(): string {
    const r = [];
    for (const char of this._nextId) {
      r.unshift(this._chars[char]);
    }
    this._increment();
    return r.join('');
  }

  _increment(): void {
    for (let i = 0; i < this._nextId.length; i++) {
      const val = ++this._nextId[i];
      if (val >= this._chars.length) {
        this._nextId[i] = 0;
      } else {
        return;
      }
    }
    this._nextId.push(0);
  }

   
  *[Symbol.iterator]() {
    while (true) {
      yield this.next();
    }
  }

  static increment(letterNumber: number): string {
    const Generator = new this('ABCDEFGHIJKLMNOPQRSTUVWXYZ');
    return Generator.start(letterNumber).next();
  }

  static setRange(range: string, colId: number, length?: number): string {
    const info = range.split('!');
    const spans = info[1].split(':');
    const start = spans[0].split('$');
    const startRow = Number(spans[0].split('$')[2]);

    const colLetter = CellIdHelper.increment(colId);

    let endCell = '';
    if (length !== undefined) {
      const endRow = String(startRow + length - 1);
      endCell = `:$${colLetter}$${endRow}`;
    }

    const newRange = `${info[0]}!$${colLetter}$${start[2]}${endCell}`;
    return newRange;
  }

  static getSpanString(
    startColNumber: number,
    startRowNumber: number,
    cols: number,
    rows: number,
  ): string {
    const startColLetter = CellIdHelper.increment(startColNumber);
    const endColLetter = CellIdHelper.increment(startColNumber + cols);
    const endRowNumber = startRowNumber + rows;
    return `${startColLetter}${startRowNumber}:${endColLetter}${endRowNumber}`;
  }

  static getCellAddressString(c: number, r: number): string {
    const colLetter = CellIdHelper.increment(c);
    return `${colLetter}${r + 1}`;
  }

  /**
   * Extracts the zero-based column index from a cell address.
   * e.g. "A1" => 0, "B12" => 1, "AA3" => 26
   * Returns null if there are no leading letters to parse.
   */
  static getColumnIndex(address: string): number | null {
    const letters = address ? address.match(/^[A-Za-z]+/) : null;
    if (!letters) {
      return null;
    }

    const chars = letters[0].toUpperCase();
    let index = 0;
    for (let i = 0; i < chars.length; i++) {
      index = index * 26 + (chars.charCodeAt(i) - 64);
    }

    return index - 1;
  }

  /**
   * Extracts the row number from a cell address.
   * e.g. "A1" => 1, "B12" => 12
   * Returns null if there is no trailing number to parse.
   */
  static getRowNumber(address: string): number | null {
    const digits = address ? address.match(/\d+$/) : null;
    return digits ? Number(digits[0]) : null;
  }
}
