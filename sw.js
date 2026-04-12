const scriptUrl = new URL(self.location.href);
const query = scriptUrl.searchParams;

const SW_CONFIG = {
  cloudBase: normalizeBase(query.get("cloudBase") || ""),
  appBasePath: normalizeAppBasePath(query.get("appBasePath") || "/"),
};

function normalizeBase(value) {
  return String(value || "")
    .trim()
    .replace(/\/+$/, "");
}

function normalizeAppBasePath(value) {
  const text = String(value || "/").trim();
  if (!text) {
    return "/";
  }
  return text.endsWith("/") ? text : text + "/";
}

function toAbsoluteUrl(value) {
  try {
    return new URL(String(value || ""), self.location.origin).toString();
  } catch (error) {
    return "";
  }
}

function resolveNotificationClickUrl(raw) {
  const text = String(raw || "").trim();
  if (text) {
    const absolute = toAbsoluteUrl(text);
    if (absolute) {
      return absolute;
    }
  }
  return toAbsoluteUrl(SW_CONFIG.appBasePath) || self.registration.scope || self.location.origin;
}

async function fetchPendingPushEvents() {
  if (!SW_CONFIG.cloudBase) {
    return [];
  }
  const subscription = await self.registration.pushManager.getSubscription();
  if (!subscription || !subscription.endpoint) {
    return [];
  }
  const url =
    SW_CONFIG.cloudBase +
    "/v2/push/pending?endpoint=" +
    encodeURIComponent(String(subscription.endpoint || "").trim());
  try {
    const response = await fetch(url, {
      method: "GET",
      cache: "no-store",
    });
    if (!response.ok) {
      return [];
    }
    const parsed = await response.json();
    const events = parsed && Array.isArray(parsed.events) ? parsed.events : [];
    return events;
  } catch (error) {
    return [];
  }
}

async function showPendingPushEvents() {
  const events = await fetchPendingPushEvents();
  if (!events.length) {
    return;
  }
  const iconUrl = toAbsoluteUrl(SW_CONFIG.appBasePath + "assets/icons/icon-192.png");
  for (const item of events) {
    const title = String(item && item.title ? item.title : "").trim() || "工作交接通知";
    const body = String(item && item.body ? item.body : "").trim();
    const clickUrl = resolveNotificationClickUrl(item && item.clickUrl ? item.clickUrl : "");
    const tag = "handover-push-" + String(item && item.id ? item.id : Date.now());
    await self.registration.showNotification(title, {
      body: body,
      icon: iconUrl || undefined,
      badge: iconUrl || undefined,
      tag: tag,
      renotify: true,
      data: {
        clickUrl: clickUrl,
      },
    });
  }
}

self.addEventListener("install", (event) => {
  event.waitUntil(self.skipWaiting());
});

self.addEventListener("activate", (event) => {
  event.waitUntil(self.clients.claim());
});

self.addEventListener("message", (event) => {
  const data = event && event.data && typeof event.data === "object" ? event.data : null;
  if (!data || data.type !== "HANDOVER_SW_CONFIG") {
    return;
  }
  if (typeof data.cloudBase === "string" && data.cloudBase.trim()) {
    SW_CONFIG.cloudBase = normalizeBase(data.cloudBase);
  }
  if (typeof data.appBasePath === "string" && data.appBasePath.trim()) {
    SW_CONFIG.appBasePath = normalizeAppBasePath(data.appBasePath);
  }
});

self.addEventListener("push", (event) => {
  event.waitUntil(showPendingPushEvents());
});

self.addEventListener("notificationclick", (event) => {
  event.notification.close();
  const targetUrl = resolveNotificationClickUrl(event.notification && event.notification.data && event.notification.data.clickUrl);
  event.waitUntil(
    (async () => {
      const windowClients = await self.clients.matchAll({
        type: "window",
        includeUncontrolled: true,
      });
      for (const client of windowClients) {
        if (client && "focus" in client) {
          try {
            if (client.url === targetUrl) {
              await client.focus();
              return;
            }
          } catch (error) {
            // ignore
          }
        }
      }
      for (const client of windowClients) {
        if (client && "navigate" in client && "focus" in client) {
          try {
            await client.navigate(targetUrl);
            await client.focus();
            return;
          } catch (error) {
            // ignore and try openWindow fallback
          }
        }
      }
      if (self.clients.openWindow) {
        await self.clients.openWindow(targetUrl);
      }
    })(),
  );
});
