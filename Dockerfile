FROM node:20-slim
WORKDIR /app
COPY package*.json ./
RUN npm ci --omit=dev
COPY . .
RUN mkdir -p data/Input/AX365 data/Input/"Branch Inquiry" data/Input/Checkout data/Input/"Detailed Reports" data/Output
ENV NODE_ENV=production
ENV DATA_ROOT=/app/data
EXPOSE 3000
CMD ["node", "server.js"]
