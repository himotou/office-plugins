const RESOURCE_URL_STORAGE_KEY = "link-bind.resource-url";

function getAddinBaseUrl(currentUrl: string): string {
  return new URL("./", currentUrl).toString();
}

function looksLikeLegacyDefaultUrl(resourceUrl: string, currentUrl: string): boolean {
  try {
    const parsedUrl = new URL(resourceUrl);
    const currentOrigin = new URL(currentUrl).origin;
    return parsedUrl.origin === currentOrigin && parsedUrl.pathname === "/resource-picker.html";
  } catch (error) {
    return false;
  }
}

export function getDefaultResourcePageUrl(currentUrl: string): string {
  return new URL("resource-picker.html", getAddinBaseUrl(currentUrl)).toString();
}

export function getStoredResourcePageUrl(currentUrl: string): string {
  try {
    const savedValue = window.localStorage.getItem(RESOURCE_URL_STORAGE_KEY)?.trim();
    if (!savedValue || looksLikeLegacyDefaultUrl(savedValue, currentUrl)) {
      const defaultUrl = getDefaultResourcePageUrl(currentUrl);
      window.localStorage.setItem(RESOURCE_URL_STORAGE_KEY, defaultUrl);
      return defaultUrl;
    }

    return savedValue;
  } catch (error) {
    return getDefaultResourcePageUrl(currentUrl);
  }
}

export function saveResourcePageUrl(resourceUrl: string): void {
  window.localStorage.setItem(RESOURCE_URL_STORAGE_KEY, resourceUrl.trim());
}

export function resetResourcePageUrl(currentUrl: string): string {
  const defaultUrl = getDefaultResourcePageUrl(currentUrl);
  window.localStorage.setItem(RESOURCE_URL_STORAGE_KEY, defaultUrl);
  return defaultUrl;
}

export function normalizeAbsoluteUrl(rawUrl: string): string {
  return new URL(rawUrl.trim()).toString();
}
