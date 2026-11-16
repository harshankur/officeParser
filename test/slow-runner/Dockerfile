# The slow runner: the test suites on Linux, at about a GitHub-hosted runner's speed (scripts/test-slow.js
# builds it for x86 and sets its CPUs and memory). Only the dependencies are built into the image, so it is rebuilt
# when package.json, package-lock.json or this file changes; the repository's files are copied in for
# each run.
FROM node:24.22.3-bookworm

# The HTML engine's PDF tests need a browser. Puppeteer's own Chrome has no Linux ARM build, so the
# image has Debian's Chromium (built for both, so --platform linux/arm64 works too), which puppeteer launches through PUPPETEER_EXECUTABLE_PATH. As root in a
# container Chromium cannot start its sandbox, so a wrapper starts it without one: the container is the
# boundary here, and the library's own launch options are unchanged.
RUN apt-get update \
    && apt-get install -y --no-install-recommends chromium fonts-liberation fonts-dejavu-core \
    && rm -rf /var/lib/apt/lists/* \
    && printf '#!/bin/sh\nexec /usr/bin/chromium --no-sandbox "$@"\n' > /usr/local/bin/chromium-no-sandbox \
    && chmod +x /usr/local/bin/chromium-no-sandbox

# Puppeteer's Chrome is not downloaded, and husky's hooks are not installed in a copy with no git repository.
ENV PUPPETEER_SKIP_DOWNLOAD=true \
    PUPPETEER_EXECUTABLE_PATH=/usr/local/bin/chromium-no-sandbox \
    HUSKY=0 \
    CI=true

WORKDIR /work
COPY package.json package-lock.json ./
RUN npm ci --no-audit --no-fund
