// Thanks to Nathan Wall
// https://stackoverflow.com/questions/12504042/what-is-a-method-that-can-be-used-to-increment-letters#12504061

export default class StringIdGenerator {
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

  // eslint-disable-next-line
  *[Symbol.iterator]() {
    while (true) {
      yield this.next();
    }
  }
}
