import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import { MsalProvider } from "@azure/msal-react";
import { registerSW } from "virtual:pwa-register";
import App from "./App";
import { warmApi } from "./api";
import { initializeAuth, msal } from "./auth";
import "./styles.css";

registerSW({ immediate: true });

async function bootstrap() {
  warmApi();
  await initializeAuth();

  createRoot(document.getElementById("root")!).render(
    <StrictMode>
      <MsalProvider instance={msal}>
        <App />
      </MsalProvider>
    </StrictMode>
  );
}

bootstrap().catch((error) => {
  console.error(error);
  document.body.innerHTML =
    '<main class="fatal-error"><h1>Unable to start</h1><p>Check the application configuration and reload.</p></main>';
});
