import { app } from "@microsoft/teams-js";

export interface HostInfo {
  inTeams: boolean;
  context?: app.Context;
}

let hostPromise: Promise<HostInfo> | null = null;

/** Initializes the Teams JS SDK once; resolves inTeams=false when opened in a plain browser. */
export function initHost(): Promise<HostInfo> {
  hostPromise ??= (async () => {
    try {
      await withTimeout(app.initialize(), 3000);
      const context = await app.getContext();
      applyTheme(context.app.theme);
      app.registerOnThemeChangeHandler(applyTheme);
      return { inTeams: true, context };
    } catch {
      return { inTeams: false };
    }
  })();
  return hostPromise;
}

/** Tells Teams the page finished loading (hides the host's loading indicator). */
export function notifyLoaded(): void {
  try {
    app.notifySuccess();
  } catch {
    /* not in Teams */
  }
}

export function applyTheme(theme: string | undefined): void {
  const t = theme === "default" ? "light" : theme === "contrast" ? "contrast" : theme === "dark" ? "dark" : "";
  if (t) document.documentElement.dataset.theme = t;
  else delete document.documentElement.dataset.theme;
}

function withTimeout<T>(p: Promise<T>, ms: number): Promise<T> {
  return new Promise((resolve, reject) => {
    const id = setTimeout(() => reject(new Error("timeout")), ms);
    p.then(
      (v) => (clearTimeout(id), resolve(v)),
      (e) => (clearTimeout(id), reject(e)),
    );
  });
}
