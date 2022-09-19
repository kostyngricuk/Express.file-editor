FROM node:18-alpine

EXPOSE 8080
WORKDIR /app
COPY web .
RUN npm install
CMD ["npm", "run", "start"]
