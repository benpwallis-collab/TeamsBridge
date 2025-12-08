FROM denoland/deno:alpine-1.44.4

WORKDIR /app
COPY . .

# run the server
CMD ["run", "--allow-net", "--allow-env", "teams-bridge.ts"]
