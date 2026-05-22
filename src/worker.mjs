const SITE_ROUTES = [
  {
    path: "/",
    title: "RISE Performance Platform",
    description: "Main landing page for the performance platform.",
  },
  {
    path: "/docs/",
    title: "Docs Home",
    description: "Primary documentation landing page and site index.",
  },
  {
    path: "/docs/portfolio-operations-dashboard/",
    title: "ATLAS RISE Ops Dashboard",
    description: "Operational dashboard workspace for the portfolio team.",
  },
  {
    path: "/docs/portfolio-operations-dashboard/financial-accountability.html",
    title: "Financial Accountability",
    description: "Companion dashboard for financial performance review and drill-down.",
  },
];

const BUILD_INFO = {
  name: "rise-performance-platform-site",
  compatibilityDate: "2026-05-22",
  assetDirectory: "./docs",
  workerEntrypoint: "./src/worker.mjs",
  runtime: "cloudflare-workers",
};

function json(data, init = {}) {
  const headers = new Headers(init.headers);
  if (!headers.has("content-type")) {
    headers.set("content-type", "application/json; charset=utf-8");
  }
  headers.set("cache-control", "no-store");

  return new Response(JSON.stringify(data, null, 2), {
    status: init.status ?? 200,
    headers,
  });
}

function noContent() {
  return new Response(null, { status: 204 });
}

function withCommonHeaders(response) {
  const headers = new Headers(response.headers);
  headers.set("x-content-type-options", "nosniff");
  headers.set("referrer-policy", "strict-origin-when-cross-origin");
  headers.set("x-frame-options", "DENY");
  headers.set("permissions-policy", "camera=(), microphone=(), geolocation=()");
  return new Response(response.body, {
    status: response.status,
    statusText: response.statusText,
    headers,
  });
}

function buildCandidatePaths(pathname) {
  const candidates = [pathname];

  if (pathname === "/") {
    return candidates;
  }

  const endsWithSlash = pathname.endsWith("/");
  const hasExtension = pathname.split("/").pop().includes(".");

  if (!endsWithSlash && !hasExtension) {
    candidates.push(`${pathname}/index.html`);
    candidates.push(`${pathname}.html`);
    candidates.push(`${pathname}/`);
  }

  if (endsWithSlash) {
    candidates.push(`${pathname}index.html`);
  }

  return [...new Set(candidates)];
}

async function fetchAsset(request, env) {
  const url = new URL(request.url);
  const candidates = buildCandidatePaths(url.pathname);

  for (const pathname of candidates) {
    const candidateUrl = new URL(url);
    candidateUrl.pathname = pathname;
    const candidateRequest = new Request(candidateUrl, request);
    const response = await env.ASSETS.fetch(candidateRequest);

    if (response.status !== 404 || pathname === candidates[candidates.length - 1]) {
      return response;
    }
  }

  return null;
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    if (request.method === "OPTIONS") {
      return noContent();
    }

    if (request.method !== "GET" && request.method !== "HEAD") {
      return new Response("Method Not Allowed", {
        status: 405,
        headers: {
          allow: "GET, HEAD, OPTIONS",
        },
      });
    }

    if (url.pathname === "/__health" || url.pathname === "/api/status") {
      return json({
        ok: true,
        service: "rise-performance-platform-site",
        time: new Date().toISOString(),
        routes: SITE_ROUTES.length,
      });
    }

    if (url.pathname === "/api/site-map" || url.pathname === "/api/routes") {
      return json({
        service: "rise-performance-platform-site",
        routes: SITE_ROUTES,
      });
    }

    if (url.pathname === "/api/build-info") {
      return json({
        service: "rise-performance-platform-site",
        build: BUILD_INFO,
      });
    }

    const assetResponse = await fetchAsset(request, env);
    if (assetResponse) {
      return withCommonHeaders(assetResponse);
    }

    return new Response("Not Found", {
      status: 404,
      headers: {
        "content-type": "text/plain; charset=utf-8",
        "cache-control": "no-store",
      },
    });
  },
};
