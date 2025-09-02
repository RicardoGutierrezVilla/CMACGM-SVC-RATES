FROM apify/actor-node:20

LABEL maintainer="julian@primefreight.com" description="Docker Image for CMA contracts"

# Check preinstalled packages
RUN npm ls crawlee apify puppeteer playwright

# Globally disable the update-notifier.
RUN npm config --global set update-notifier false

COPY package*.json ./ 



# Install default dependencies, print versions of everything
RUN npm --quiet set progress=false \
    && npm config --global set update-notifier false \
    && npm install --omit=dev --omit=optional --no-package-lock --prefer-online \
    && echo "Installed NPM packages:" \
    && (npm list --omit=dev --omit=optional || true) \
    && echo "Node.js version:" \
    && node --version \
    && echo "NPM version:" \
    && npm --version

# Copy all source code (respects .dockerignore)
COPY . ./

# Create and run as a non-root user.
RUN adduser -h /home/apify -D apify && \
    chown -R apify:apify ./
USER apify

CMD ["npm", "start", "--silent"]