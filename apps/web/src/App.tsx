import { useCallback, useEffect, useState } from "react";
import { useIsAuthenticated } from "@azure/msal-react";
import {
  api,
  ApiError,
  type AliasSet,
  type DomainInfo,
  type OnboardingInfo,
  type Session,
  type WorkspaceSnapshot
} from "./api";
import { getActiveAccount, signIn, signOut } from "./auth";

const DOMAIN_KEY = "alias-manager.domain";

type Notice = { tone: "success" | "error" | "info"; text: string } | null;

function Icon({ name }: { name: "copy" | "trash" | "plus" | "mail" | "shield" | "refresh" | "logout" }) {
  const paths: Record<string, React.ReactNode> = {
    copy: <><rect x="9" y="9" width="10" height="10" rx="2"/><path d="M6 15H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h8a2 2 0 0 1 2 2v1"/></>,
    trash: <><path d="M3 6h18"/><path d="M8 6V4h8v2"/><path d="M19 6l-1 14H6L5 6"/><path d="M10 11v5M14 11v5"/></>,
    plus: <><path d="M12 5v14M5 12h14"/></>,
    mail: <><rect x="3" y="5" width="18" height="14" rx="2"/><path d="m4 7 8 6 8-6"/></>,
    shield: <><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10"/><path d="m9 12 2 2 4-4"/></>,
    refresh: <><path d="M20 11a8 8 0 1 0-2.3 5.7"/><path d="M20 4v7h-7"/></>,
    logout: <><path d="M10 17l5-5-5-5"/><path d="M15 12H3"/><path d="M14 3h5a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2h-5"/></>
  };

  return (
    <svg
      className="icon"
      viewBox="0 0 24 24"
      aria-hidden="true"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.8"
      strokeLinecap="round"
      strokeLinejoin="round"
    >
      {paths[name]}
    </svg>
  );
}

function App() {
  const authenticated = useIsAuthenticated();
  const account = getActiveAccount();
  const [session, setSession] = useState<Session | null>(null);
  const [organization, setOrganization] = useState("");
  const [domain, setDomain] = useState("");
  const [domains, setDomains] = useState<DomainInfo[]>([]);
  const [aliasSet, setAliasSet] = useState<AliasSet | null>(null);
  const [onboarding, setOnboarding] = useState<OnboardingInfo | null>(null);
  const [busy, setBusy] = useState("");
  const [notice, setNotice] = useState<Notice>(null);
  const [organizationReady, setOrganizationReady] = useState(false);
  const [pendingDelete, setPendingDelete] = useState("");

  const clientConfigured = Boolean(
    import.meta.env.VITE_ENTRA_CLIENT_ID &&
      import.meta.env.VITE_API_SCOPE &&
      import.meta.env.VITE_API_BASE_URL
  );

  const handleError = useCallback((error: unknown) => {
    const message = error instanceof Error ? error.message : "The request failed.";
    setNotice({ tone: "error", text: message });

    if (error instanceof ApiError && error.code === "EXCHANGE_NOT_READY") {
      setOrganizationReady(false);
    }
  }, []);

  const applyWorkspace = useCallback((snapshot: WorkspaceSnapshot) => {
    setOrganization(snapshot.organization);
    setDomains(snapshot.domains);
    setAliasSet(snapshot.aliasSet);
    setOnboarding(null);
    setOrganizationReady(true);
    setPendingDelete("");

    setDomain((current) => {
      const remembered = localStorage.getItem(DOMAIN_KEY) ?? "";
      const currentIsValid = snapshot.domains.some((item) => item.domain === current);
      const rememberedIsValid = snapshot.domains.some((item) => item.domain === remembered);
      const preferredCustomDomain = snapshot.domains.find(
        (item) => !item.domain.endsWith(".onmicrosoft.com")
      )?.domain;
      const nextDomain =
        (currentIsValid ? current : "") ||
        (rememberedIsValid && !remembered.endsWith(".onmicrosoft.com") ? remembered : "") ||
        preferredCustomDomain ||
        snapshot.domains.find((item) => item.isDefault)?.domain ||
        snapshot.domains[0]?.domain ||
        "";

      if (nextDomain) localStorage.setItem(DOMAIN_KEY, nextDomain);
      return nextDomain;
    });
  }, []);

  const loadWorkspace = useCallback(async (mode: "bootstrap" | "refresh" = "bootstrap") => {
    setBusy(mode);
    if (mode === "bootstrap") setNotice(null);

    try {
      const snapshot = await api.bootstrap();
      applyWorkspace(snapshot);
    } catch (error) {
      setAliasSet(null);
      if (error instanceof ApiError && error.code === "EXCHANGE_NOT_READY") {
        try {
          setOnboarding(await api.onboarding());
        } catch {
          setOnboarding(null);
        }
      }
      handleError(error);
    } finally {
      setBusy("");
    }
  }, [applyWorkspace, handleError]);

  useEffect(() => {
    if (!authenticated) {
      setSession(null);
      setOrganization("");
      setDomains([]);
      setAliasSet(null);
      setOnboarding(null);
      setOrganizationReady(false);
      setPendingDelete("");
      setBusy("");
      return;
    }

    void api.session().then(setSession).catch(handleError);
    void loadWorkspace("bootstrap");
  }, [authenticated, handleError, loadWorkspace]);

  async function createAlias() {
    if (!organization || !domain) return;

    setBusy("create");
    setNotice(null);
    try {
      const data = await api.createAlias(organization, domain);
      setAliasSet(data);
      setNotice({
        tone: "success",
        text: `${data.aliases[0]?.address ?? "The new alias"} is ready. Mail goes to your inbox.`
      });
    } catch (error) {
      handleError(error);
    } finally {
      setBusy("");
    }
  }

  async function deleteAlias(address: string) {
    if (!organization) return;

    setBusy(address);
    setNotice(null);
    try {
      const data = await api.deleteAlias(organization, address);
      setAliasSet(data);
      setPendingDelete("");
      setNotice({ tone: "info", text: `${address} is deleted.` });
    } catch (error) {
      handleError(error);
    } finally {
      setBusy("");
    }
  }

  async function copy(text: string, label = "Copied.") {
    try {
      await navigator.clipboard.writeText(text);
      setNotice({ tone: "success", text: label });
    } catch {
      setNotice({ tone: "error", text: "The browser could not copy the text." });
    }
  }

  if (!clientConfigured) {
    return (
      <main className="fatal-error" id="main-content">
        <h1>Configuration required</h1>
        <p>Set the client ID, API scope, and API URL build variables.</p>
      </main>
    );
  }

  if (!authenticated) {
    return (
      <>
        <a className="skip-link" href="#main-content">Skip to main content</a>
        <main className="signin-shell" id="main-content">
          <section className="signin-panel" aria-labelledby="welcome-title">
            <div className="brand brand-large">
              <span className="brand-mark"><Icon name="mail" /></span>
              <span>Mail Alias Manager</span>
            </div>
            <p className="eyebrow">Microsoft 365 aliases</p>
            <h1 id="welcome-title">Use email aliases with your current inbox.</h1>
            <p className="hero-copy">
              Create short aliases for sign-ups and forms. Mail arrives in your current Microsoft 365 inbox.
            </p>
            <button className="primary-button sign-in-button" type="button" onClick={() => void signIn()}>
              Sign in with Microsoft
            </button>
            <div className="trust-row" aria-label="Key features">
              <span><Icon name="shield" /> Uses your existing mailbox</span>
              <span>Keeps up to 30 aliases</span>
            </div>
          </section>

          <aside className="signin-aside" aria-label="How aliases work">
            <div className="route-demo">
              <span className="route-label">ALIASES</span>
              <div className="route-addresses">
                <span>jpeterson@company.com</span>
                <span>mcarter@company.com</span>
                <span>abrooks@company.com</span>
              </div>
              <div className="route-arrow" aria-hidden="true">↓</div>
              <span className="route-label">YOUR INBOX</span>
              <strong>you@company.com</strong>
              <p className="route-note">Create an alias. Use it. Mail arrives in your inbox.</p>
            </div>
          </aside>
        </main>
      </>
    );
  }

  const inboxAddress = aliasSet?.primaryAddress || session?.username || account?.username || "Your Microsoft 365 inbox";

  return (
    <div className="app-shell">
      <a className="skip-link" href="#main-content">Skip to main content</a>

      <header className="topbar">
        <div className="brand">
          <span className="brand-mark"><Icon name="mail" /></span>
          <span>Mail Alias Manager</span>
        </div>
        <div className="account-menu">
          <div>
            <strong>{session?.name ?? account?.name ?? "Microsoft 365 user"}</strong>
            <span>{session?.username ?? account?.username}</span>
          </div>
          <button
            className="icon-button"
            type="button"
            onClick={() => void signOut()}
            aria-label="Sign out"
            title="Sign out"
          >
            <Icon name="logout" />
          </button>
        </div>
      </header>

      <main className="workspace" id="main-content">
        <section className="page-heading" aria-labelledby="page-title">
          <div>
            <p className="eyebrow">Mail alias manager</p>
            <h1 id="page-title">Email aliases</h1>
            <p>Create an alias and use it where you need email. Mail goes to your current inbox.</p>
          </div>
          {organizationReady && aliasSet ? (
            <div className="capacity" aria-label={`${aliasSet.count} of ${aliasSet.limit} aliases active`}>
              <strong>{aliasSet.count}</strong>
              <span>of {aliasSet.limit} active</span>
            </div>
          ) : null}
        </section>

        {notice ? (
          <div className={`notice notice-${notice.tone}`} role={notice.tone === "error" ? "alert" : "status"}>
            {notice.text}
          </div>
        ) : null}

        {busy === "bootstrap" && !organizationReady ? (
          <section className="connection-panel loading-panel" aria-live="polite">
            <span className="loading-dot" aria-hidden="true" />
            <div>
              <h2>Connecting to Microsoft 365</h2>
              <p>Loading your inbox, domains, and aliases.</p>
            </div>
          </section>
        ) : null}

        {organizationReady ? (
          <section className="connection-panel" aria-labelledby="inbox-heading">
            <div className="section-heading">
              <p className="eyebrow">Connected</p>
              <h2 id="inbox-heading">Your inbox</h2>
              <p>Mail to your aliases arrives here.</p>
            </div>
            <dl className="connection-details">
              <div>
                <dt>Inbox</dt>
                <dd>{inboxAddress}</dd>
              </div>
              <div>
                <dt>Organization</dt>
                <dd>{organization}</dd>
              </div>
            </dl>
          </section>
        ) : null}

        {!organizationReady && onboarding ? (
          <section className="onboarding-panel" aria-labelledby="admin-heading">
            <div className="section-heading">
              <p className="eyebrow">Admin setup</p>
              <h2 id="admin-heading">Authorize this app once</h2>
              <p>A Microsoft 365 admin must authorize the app. Then the admin must run the Exchange setup script.</p>
            </div>
            <div className="onboarding-actions">
              <a className="primary-button" href={onboarding.adminConsentUrl} target="_blank" rel="noreferrer">
                Grant admin consent
              </a>
              <button
                className="secondary-button"
                type="button"
                onClick={() => void copy(onboarding.setupScript, "Setup script copied.")}
              >
                <Icon name="copy" /> Copy setup script
              </button>
            </div>
            <details>
              <summary>Review setup script</summary>
              <pre tabIndex={0}><code>{onboarding.setupScript}</code></pre>
            </details>
          </section>
        ) : null}

        {organizationReady ? (
          <>
            <section className="create-panel" aria-labelledby="create-heading">
              <div className="section-heading">
                <p className="eyebrow">New alias</p>
                <h2 id="create-heading">Create an alias</h2>
                <p>Choose a domain. The app creates a short name.</p>
              </div>

              <form
                className="create-form"
                onSubmit={(event) => {
                  event.preventDefault();
                  void createAlias();
                }}
              >
                <div className="field">
                  <label htmlFor="domain">Domain</label>
                  <select
                    id="domain"
                    name="domain"
                    value={domain}
                    aria-describedby="domain-help"
                    onChange={(event) => {
                      setDomain(event.target.value);
                      localStorage.setItem(DOMAIN_KEY, event.target.value);
                    }}
                  >
                    {domains.map((item) => (
                      <option value={item.domain} key={item.domain}>
                        {item.domain}{item.isDefault ? " — default" : ""}
                      </option>
                    ))}
                  </select>
                  <span className="field-help" id="domain-help">
                    Example: jpeterson@{domain || "your-domain.com"}
                  </span>
                </div>

                <button
                  className="primary-button create-button"
                  type="submit"
                  disabled={busy === "create" || !domain}
                  aria-busy={busy === "create"}
                >
                  <Icon name="plus" />
                  {busy === "create" ? "Creating…" : "Create alias"}
                </button>
              </form>
            </section>

            <section className="aliases-section" aria-labelledby="aliases-heading" aria-busy={busy === "refresh"}>
              <div className="aliases-toolbar">
                <div>
                  <p className="eyebrow">Aliases</p>
                  <h2 id="aliases-heading">Your aliases</h2>
                  <p>Newest aliases appear first.</p>
                </div>
                <button
                  className="secondary-button compact-button"
                  type="button"
                  onClick={() => void loadWorkspace("refresh")}
                  disabled={busy === "refresh"}
                  aria-busy={busy === "refresh"}
                >
                  <Icon name="refresh" />
                  {busy === "refresh" ? "Refreshing…" : "Refresh"}
                </button>
              </div>

              {!aliasSet ? (
                <div className="loading-state" role="status">
                  <span className="loading-dot" aria-hidden="true" />
                  <div>
                    <h3>Loading aliases</h3>
                    <p>Reading the current mailbox state.</p>
                  </div>
                </div>
              ) : aliasSet.aliases.length ? (
                <ol className="alias-list">
                  {aliasSet.aliases.map((alias, index) => (
                    <li key={alias.address} className={`alias-row${index === 0 ? " is-newest" : ""}`}>
                      <span className="alias-icon" aria-hidden="true"><Icon name="mail" /></span>
                      <div className="alias-address">
                        <strong>{alias.address}</strong>
                        {index === 0 ? <span className="newest-badge">Newest</span> : null}
                      </div>

                      <div className="row-actions">
                        {pendingDelete === alias.address ? (
                          <>
                            <span className="delete-question">Delete this alias?</span>
                            <button
                              className="secondary-button compact-button"
                              type="button"
                              autoFocus
                              onClick={() => setPendingDelete("")}
                              disabled={busy === alias.address}
                            >
                              Cancel
                            </button>
                            <button
                              className="danger-button compact-button"
                              type="button"
                              onClick={() => void deleteAlias(alias.address)}
                              disabled={busy === alias.address}
                              aria-busy={busy === alias.address}
                            >
                              {busy === alias.address ? "Deleting…" : "Delete"}
                            </button>
                          </>
                        ) : (
                          <>
                            <button
                              className="secondary-button compact-button"
                              type="button"
                              onClick={() => void copy(alias.address, "Alias copied.")}
                            >
                              <Icon name="copy" /> Copy
                            </button>
                            <button
                              className="icon-button danger"
                              type="button"
                              onClick={() => {
                                setPendingDelete(alias.address);
                                setNotice(null);
                              }}
                              aria-label={`Delete ${alias.address}`}
                              title="Delete alias"
                            >
                              <Icon name="trash" />
                            </button>
                          </>
                        )}
                      </div>
                    </li>
                  ))}
                </ol>
              ) : (
                <div className="empty-state">
                  <span className="empty-icon"><Icon name="mail" /></span>
                  <h3>No aliases yet</h3>
                  <p>Create an alias above. Mail to the alias arrives in {inboxAddress}.</p>
                </div>
              )}

              <footer className="ledger-footer">
                <strong>Alias limit</strong>
                <p>You can keep up to 30 aliases. When you create alias 31, the app deletes the oldest alias that it created.</p>
              </footer>
            </section>
          </>
        ) : null}
      </main>
    </div>
  );
}

export default App;
