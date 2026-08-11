declare module 'regexp.escape' {
  /** Ponyfill of the `RegExp.escape` proposal. */
  const escape: (value: string) => string;
  export default escape;
}
