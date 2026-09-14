import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import tailwindcss from "@tailwindcss/vite";

function debugTaxFormEndpoint() {
  return {
    name: "debug-tax-form-endpoint",

    configureServer(server) {
      server.middlewares.use("/api/debug/tax-records", (req, res, next) => {
        if (req.method !== "POST") {
          next();
          return;
        }

        const chunks = [];

        req.on("data", (chunk) => {
          chunks.push(chunk);
        });

        req.on("end", () => {
          const body = Buffer.concat(chunks);

          console.log(
            `[Tax Form Debug] Received ${body.length} bytes of FormData`,
          );

          res.statusCode = 200;

          res.setHeader("Content-Type", "application/json");

          res.end(
            JSON.stringify({
              ok: true,
              message: "FormData received successfully.",
              receivedBytes: body.length,
            }),
          );
        });

        req.on("error", (error) => {
          console.error("[Tax Form Debug]", error);

          res.statusCode = 500;

          res.setHeader("Content-Type", "application/json");

          res.end(
            JSON.stringify({
              ok: false,
              message: "Unable to read FormData.",
            }),
          );
        });
      });
    },
  };
}

export default defineConfig({
  plugins: [react(), tailwindcss(), debugTaxFormEndpoint()],
});
