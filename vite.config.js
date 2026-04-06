import { defineConfig } from "vite";
import { getHttpsServerOptions } from "office-addin-dev-certs";

export default defineConfig(async ({ command }) => {
  // Richiedi i certificati SSL di sviluppo generati da Microsoft
  const httpsOptions = await getHttpsServerOptions();

  return {
    base: command === 'build' ? '/ExcelAddin/' : '/',
    server: {
      host: true, // Permette l'accesso da altri PC nella rete locale
      port: 3000,
      https: httpsOptions,
      strictPort: true,
      headers: {
        "Access-Control-Allow-Origin": "*", // Evita errori CORS quando è hostato tramite iframe in Excel
      },
    },
  };
});
