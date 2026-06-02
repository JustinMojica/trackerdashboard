if (process.env.PORT && !process.env.TRACKER_SERVER_PORT) {
  process.env.TRACKER_SERVER_PORT = process.env.PORT;
}

import("./server/secureAccessServer.mjs").catch((error) => {
  console.error(error);
  process.exit(1);
});
