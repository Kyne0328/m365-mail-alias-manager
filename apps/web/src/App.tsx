import { useCallback, useEffect, useMemo, useState } from "react";
import { useIsAuthenticated } from "@azure/msal-react";
import { api, ApiError, type AliasSet, type DomainInfo, type OnboardingInfo, type Session } from "./api";
import { getActiveAccount, signIn, signOut } from "./auth";

const ORG_KEY = "alias-manager.organization";
const DOMAIN_KEY = "alias-manager.domain";
const PREFIX_KEY = "alias-manager.prefix";
const DEFAULT_PREFIX = "temp";

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

  return <svg className="icon" viewBox="0 0 24 24" aria-hidden="true" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">{paths[name]}</svg>;
}

function App() {
  const authenticated = useIsAuthenticated();
  const account = getActiveAccount();
  const [session, setSession] = useState<Session | null>(null);
  const [organization, setOrganization] = useState(() => localStorage.getItem(ORG_KEY) ?? "");
  const [domain, setDomain] = useState(() => localStorage.getItem(DOMAIN_KEY) ?? "");
  const [prefix, setPrefix] = useState(() => localStorage.getItem(PREFIX_KEY) ?? DEFAULT_PREFIX);
  const [domains, setDomains] = useState<DomainInfo[]>([]);
  const [aliasSet, setAliasSet] = useState<AliasSet | null>(null);
  const [onboarding, setOnboarding] = useState<OnboardingInfo | null>(null);
  const [busy, setBusy] = useState("");
  const [notice, setNotice] = useState<Notice>(null);
  const [organizationReady, setOrganizationReady] = useState(false);

  const clientConfigured = Boolean(
    import.meta.env.VITE_ENTRA_CLIENT_ID &&
      import.meta.env.VITE_API_SCOPE &&
      import.meta.env.VITE_API_BASE_URL
  );

  const normalizedPrefix = useMemo(
    () => prefix.trim().toLowerCase().replace(/\s+/g, "-"),
    [prefix]
  );

  const handleError = useCallback((error: unknown) => {
    const message = error instanceof Error ? error.message : "Something went wrong.";
    setNotice({ tone: "error", text: message });

    if (error instanceof ApiError && error.code === "EXCHANGE_NOT_READY") {
      setOrganizationReady(false);
    }
  }, []);

  useEffect(() => {
    if (!authenticated) {
      setSession(null);
      return;
    }

    api.session().then(setSession).catch(handleError);
  }, [authenticated, handleError]);

  const loadDomains = useCallback(async (org = organization) => {
    const clean = org.trim().toLowerCase();
    if (!clean) return;

    setBusy("organization");
    setNotice(null);

    try {
      const [domainResult, onboardingResult] = await Promise.all([
        api.domains(clean),
        api.onboarding(clean)
      ]);

      setDomains(domainResult.domains);
      setOnboarding(onboardingResult);
      setOrganizationReady(true);
      localStorage.setItem(ORG_KEY, clean);
      setOrganization(clean);

      const stillValid = domainResult.domains.some((item) => item.domain === domain);
      const nextDomain =
        stillValid
          ? domain
          : domainResult.domains.find((item) => item.isDefault)?.domain ??
            domainResult.domains[0]?.domain ??
            "";

      setDomain(nextDomain);
      if (nextDomain) localStorage.setItem(DOMAIN_KEY, nextDomain);
    } catch (error) {
      try {
        setOnboarding(await api.onboarding(clean));
      } catch {
        setOnboarding(null);
      }
      handleError(error);
    } finally {
      setBusy("");
    }
  }, [domain, handleError, organization]);

  useEffect(() => {
    if (authenticated && organization) {
      void loadDomains(organization);
    }
  }, [authenticated]);

  const loadAliases = useCallback(async () => {
    if (!organizationReady) return;

    setBusy("aliases");
    try {
      const data = await api.aliases(organization);
      setAliasSet(data);
    } catch (error) {
      handleError(error);
    } finally {
      setBusy("");
    }
  }, [handleError, organization, organizationReady]);

  useEffect(() => {
    if (organizationReady) {
      void loadAliases();
    }
  }, [organizationReady]);

  async function generateAlias() {
    if (!domain || !normalizedPrefix) return;

    setBusy("generate");
    setNotice(null);
    try {
      const data = await api.createAlias(organization, domain, normalizedPrefix);
      setAliasSet(data);
      localStorage.setItem(PREFIX_KEY, normalizedPrefix);
      setPrefix(normalizedPrefix);
      setNotice({
        tone: "success",
        text: `${data.aliases[0]?.address ?? "The new alias"} is ready to receive mail.`
      });
    } catch (error) {
      handleError(error);
    } finally {
      setBusy("");
    }
  }

  async function deleteAlias(address: string) {
    setBusy(address);
    setNotice(null);
    try {
      const data = await api.deleteAlias(organization, address);
      setAliasSet(data);
      setNotice({ tone: "info", text: `${address} was removed.` });
    } catch (error) {
      handleError(error);
    } finally {
      setBusy("");
    }
  }

  async function copy(text: string, label = "Copied to clipboard.") {
    await navigator.clipboard.writeText(text);
    setNotice({ tone: "success", text: label });
  }

  if (!clientConfigured) {
    return (
      <main className="fatal-error">
        <h1>Configuration required</h1>
        <p>Set the VITE_ENTRA_CLIENT_ID, VITE_API_SCOPE, and VITE_API_BASE_URL build variables.</p>
      </main>
    );
  }

  if (!authenticated) {
    return (
      <main className="signin-shell">
        <section className="signin-panel" aria-labelledby="welcome-title">
          <div className="brand brand-large">
            <span className="brand-mark"><Icon name="mail" /></span>
            <span>Mail Alias Manager</span>
          </div>
          <p className="eyebrow">Microsoft 365 alias rotation</p>
          <h1 id="welcome-title">Fresh inbox aliases, without fresh mailboxes.</h1>
          <p className="hero-copy">
            Sign in with a work or school account. Your organization chooses whether to authorize the app; the service never hard-codes or assumes your tenant, domain, or mailbox.
          </p>
          <button className="primary-button sign-in-button" type="button" onClick={() => void signIn()}>
            Continue with Microsoft
            <span aria-hidden="true">→</span>
          </button>
          <div className="trust-row">
            <span><Icon name="shield" /> Tenant-admin controlled</span>
            <span>30-alias FIFO set</span>
          </div>
        </section>
        <aside className="signin-aside" aria-label="How it works">
          <div className="route-demo">
            <span className="route-label">YOUR MAILBOX</span>
            <strong>you@company.com</strong>
            <div className="route-line" />
            <div className="route-addresses">
              <span>m365am-temp-000031-a1b2c3d4e5f6@company.com</span>
              <span>m365am-temp-000030-14e98a20bc31@company.com</span>
              <span>m365am-temp-000029-7c40f61b882a@company.com</span>
            </div>
            <div className="route-note">31st created → oldest removed</div>
          </div>
        </aside>
      </main>
    );
  }

  return (
    <div className="app-shell">
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
          <button className="icon-button" type="button" onClick={() => void signOut()} aria-label="Sign out">
            <Icon name="logout" />
          </button>
        </div>
      </header>

      <main className="workspace">
        <section className="page-heading">
          <div>
            <p className="eyebrow">Alias workspace</p>
            <h1>Disposable addresses. One real inbox.</h1>
            <p>Generated aliases receive mail in your existing Microsoft 365 mailbox. No extra mailbox license is created.</p>
          </div>
          {organizationReady && aliasSet ? (
            <div className="capacity" aria-label={`${aliasSet.count} of ${aliasSet.limit} managed aliases active`}>
              <span>{aliasSet.count}</span>
              <small>/ {aliasSet.limit} active</small>
            </div>
          ) : null}
        </section>

        {notice ? (
          <div className={`notice notice-${notice.tone}`} role={notice.tone === "error" ? "alert" : "status"}>
            {notice.text}
          </div>
        ) : null}

        <section className="configuration-panel" aria-labelledby="organization-heading">
          <div className="section-number">01</div>
          <div className="configuration-copy">
            <h2 id="organization-heading">Connect your organization</h2>
            <p>Enter the tenant’s primary <code>.onmicrosoft.com</code> domain. The API verifies it belongs to the tenant in your sign-in token.</p>
          </div>
          <form
            className="organization-form"
            onSubmit={(event) => {
              event.preventDefault();
              void loadDomains();
            }}
          >
            <label htmlFor="organization">Organization domain</label>
            <div className="input-action">
              <input
                id="organization"
                name="organization"
                value={organization}
                onChange={(event) => setOrganization(event.target.value)}
                placeholder="contoso.onmicrosoft.com"
                inputMode="url"
                autoCapitalize="none"
                spellCheck={false}
                required
                pattern="[A-Za-z0-9.-]+\.onmicrosoft\.com"
                aria-describedby="organization-help"
              />
              <button className="secondary-button" type="submit" disabled={busy === "organization"}>
                {busy === "organization" ? "Checking…" : "Connect"}
              </button>
            </div>
            <span className="field-help" id="organization-help">Example: contoso.onmicrosoft.com</span>
          </form>
        </section>

        {!organizationReady && onboarding ? (
          <section className="onboarding-panel" aria-labelledby="admin-heading">
            <div>
              <p className="eyebrow">Tenant administrator</p>
              <h2 id="admin-heading">One-time authorization is required</h2>
              <p>
                An Exchange administrator must grant the application permission and run the least-privilege setup script once for this tenant.
              </p>
            </div>
            <div className="onboarding-actions">
              <a className="primary-button" href={onboarding.adminConsentUrl} target="_blank" rel="noreferrer">
                Grant admin consent
              </a>
              <button className="secondary-button" type="button" onClick={() => void copy(onboarding.setupScript, "Setup script copied.")}>
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
            <section className="control-strip" aria-labelledby="alias-settings-heading">
              <div className="section-number">02</div>
              <div className="control-copy">
                <h2 id="alias-settings-heading">Alias pattern</h2>
                <p>Choose an accepted domain and a recognizable prefix. The counter is calculated from Exchange every time.</p>
              </div>
              <div className="control-fields">
                <div className="field">
                  <label htmlFor="domain">Domain</label>
                  <select id="domain" name="domain" value={domain} onChange={(event) => setDomain(event.target.value)}>
                    {domains.map((item) => (
                      <option value={item.domain} key={item.domain}>
                        {item.domain}{item.isDefault ? " — default" : ""}
                      </option>
                    ))}
                  </select>
                </div>
                <div className="field">
                  <label htmlFor="prefix">Prefix</label>
                  <input
                    id="prefix"
                    name="prefix"
                    value={prefix}
                    onChange={(event) => setPrefix(event.target.value)}
                    maxLength={24}
                    pattern="[A-Za-z0-9][A-Za-z0-9._-]{0,23}"
                    required
                    autoCapitalize="none"
                    spellCheck={false}
                    aria-describedby="prefix-help"
                  />
                  <span className="field-help" id="prefix-help">Creates m365am-{normalizedPrefix || "temp"}-000001-xxxxxxxxxxxx@{domain || "…"}</span>
                </div>
              </div>
            </section>

            <section className="aliases-section" aria-labelledby="aliases-heading">
              <div className="aliases-toolbar">
                <div>
                  <p className="eyebrow">03 · Active aliases</p>
                  <h2 id="aliases-heading">Routing ledger</h2>
                </div>
                <div className="toolbar-actions">
                  <button className="icon-button" type="button" onClick={() => void loadAliases()} aria-label="Refresh aliases" disabled={busy === "aliases"}>
                    <Icon name="refresh" />
                  </button>
                  <button className="primary-button generate-button" type="button" onClick={() => void generateAlias()} disabled={busy === "generate" || !normalizedPrefix}>
                    <Icon name="plus" />
                    {busy === "generate" ? "Creating…" : "Generate email"}
                  </button>
                </div>
              </div>

              {aliasSet?.aliases.length ? (
                <ol className="alias-list">
                  {aliasSet.aliases.map((alias, index) => (
                    <li key={alias.address} className="alias-row">
                      <div className="sequence-marker" aria-hidden="true">
                        <span>{String(index + 1).padStart(2, "0")}</span>
                      </div>
                      <div className="alias-address">
                        <strong>{alias.address.split("@")[0]}</strong>
                        <span>@{alias.address.split("@")[1]}</span>
                      </div>
                      <div className="alias-meta">
                        <span>#{String(alias.sequence).padStart(6, "0")}</span>
                        {index === 0 ? <span className="newest-badge">Newest</span> : null}
                      </div>
                      <div className="row-actions">
                        <button className="icon-button" type="button" onClick={() => void copy(alias.address)} aria-label={`Copy ${alias.address}`}>
                          <Icon name="copy" />
                        </button>
                        <button className="icon-button danger" type="button" onClick={() => void deleteAlias(alias.address)} disabled={busy === alias.address} aria-label={`Delete ${alias.address}`}>
                          <Icon name="trash" />
                        </button>
                      </div>
                    </li>
                  ))}
                </ol>
              ) : (
                <div className="empty-state">
                  <span className="empty-icon"><Icon name="mail" /></span>
                  <h3>No managed aliases yet</h3>
                  <p>Generate your first address. It will route into {aliasSet?.mailbox ?? session?.username ?? "your mailbox"}.</p>
                  <button className="primary-button" type="button" onClick={() => void generateAlias()}>
                    <Icon name="plus" /> Generate first email
                  </button>
                </div>
              )}

              <footer className="ledger-footer">
                <span>FIFO policy</span>
                <p>The app keeps at most 30 managed aliases on this mailbox. Creating the next one removes the oldest app-managed alias first, even if you changed prefix or domain.</p>
              </footer>
            </section>
          </>
        ) : null}
      </main>
    </div>
  );
}

export default App;
