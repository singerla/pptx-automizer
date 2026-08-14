/**
 * Performance patch for @xmldom/xmldom's SAX parser (verified on 0.9.10).
 *
 * `lib/sax.js` compiles a fresh RegExp for every closing tag it parses
 * (`g.reg('^', g.QName_group, g.S_OPT, '$')`) and for every `getMatch`
 * probe - measured at ~19% of total runtime when appending large slides,
 * since every part read from an archive goes through this parser.
 *
 * `grammar.reg` builds its RegExp purely from its arguments and returns it
 * without the `g` flag (no `lastIndex` state), and the runtime call sites
 * combine a small fixed set of grammar fragments - so caching the compiled
 * RegExp by source is safe and cache growth is bounded by the grammar.
 *
 * Applied via the exports object, which is the `g` the SAX parser calls
 * through; `grammar.js`-internal calls at module load are unaffected.
 * Remove once fixed upstream in @xmldom/xmldom.
 */

let applied = false;

export function patchXmldomSaxRegExpCache(): void {
  if (applied) {
    return;
  }
  applied = true;

  try {
    // Deep import on purpose: the patch targets the module object that
    // sax.js holds. Guarded, in case a future xmldom version restricts
    // deep imports via an `exports` map - the patch is an optimization
    // only, parsing works identically without it.
    // eslint-disable-next-line @typescript-eslint/no-require-imports
    const grammar = require('@xmldom/xmldom/lib/grammar');
    if (typeof grammar?.reg !== 'function') {
      return;
    }

    const original = grammar.reg;
    const cache = new Map<string, RegExp>();

    grammar.reg = function (...args: (string | RegExp)[]): RegExp {
      // Key = the exact source string `reg` compiles (same mapping and
      // join), so two calls share a cache entry iff they build the same
      // RegExp.
      const key = args
        .map((part) => (typeof part === 'string' ? part : part.source))
        .join('');
      let compiled = cache.get(key);
      if (!compiled) {
        compiled = original.apply(this, args);
        cache.set(key, compiled);
      }
      return compiled;
    };
  } catch {
    // Patch is best-effort; never break parsing over it.
  }
}
