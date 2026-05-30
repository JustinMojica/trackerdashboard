if (process.env.PORT && !process.env.TRACKER_SERVER_PORT) {
  process.env.TRACKER_SERVER_PORT = process.env.PORT;
}

await import("./secureAccessServer.mjs");
