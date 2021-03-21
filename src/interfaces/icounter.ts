export interface ICounter {
  set(): void | PromiseLike<void>;

  get(): number;

  name: string;
  count: number;

  _increment(): number;
}
