# Tier-3 renderer (ROADMAP Phase 5): pptx-thumbnailer as a service container,
# so visual-regression baselines depend only on Docker — no LibreOffice, no
# poppler, no Node globals on contributor machines.
#
# EVERYTHING in this file pins the renderer. The visual baselines in
# __tests__/visual-baselines/ mean "what THIS image rendered when a human last
# blessed them" — bumping the base digest, the thumbnailer version or the font
# set can shift glyph metrics/antialiasing and requires regenerating baselines
# in the same PR (UPDATE_BASELINES=1 yarn test:visual).
FROM node:22-bookworm-slim@sha256:d649c27dae7ba0137b3cef5dd75baa422c08dc3d9e3fc0c23dfb172dc3cc6436

# libreoffice-impress + poppler-utils form the render pipeline
# (pptx -> pdf -> png); the fonts are metric-compatible replacements for the
# Office defaults (Liberation = Arial/Times/Courier, Carlito/Caladea =
# Calibri/Cambria) — golden decks must stick to these families.
# LibreOffice itself comes from Debian bookworm (stable: security/point
# updates only), which keeps the renderer version quasi-frozen.
RUN apt-get update && apt-get install -y --no-install-recommends \
      libreoffice-impress \
      poppler-utils \
      fonts-liberation \
      fonts-crosextra-carlito \
      fonts-crosextra-caladea \
      fonts-dejavu-core \
    && rm -rf /var/lib/apt/lists/*

# Exact version — the npm package ships server, client and CLI
RUN npm install -g pptx-thumbnailer@0.1.0

ENV NODE_ENV=production \
    PORT=3000 \
    HOME=/tmp/home

RUN mkdir -p /tmp/home && chown node:node /tmp/home
USER node

EXPOSE 3000
CMD ["pptx-thumbnailer", "serve"]
