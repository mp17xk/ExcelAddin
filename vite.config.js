import { defineConfig } from "vite";

export default defineConfig(async ({ command }) => {
  const config = {
    base: command === 'build' ? '/ExcelAddin/' : '/',
    server: {
      host: true,
      port: 3000,
      strictPort: true,
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
    },
  };

  // Carica i certificati SSL solo in sviluppo locale (non in CI/build)
  if (command === 'serve') {
    const { getHttpsServerOptions } = await import("office-addin-dev-certs");
    config.server.https = await getHttpsServerOptions();
  }

  return config;
});
