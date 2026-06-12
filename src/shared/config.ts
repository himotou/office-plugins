const RESOURCE_URL_STORAGE_KEY = "link-bind.resource-url";

export function getDefaultResourcePageUrl(origin: string): string {
  return new URL("/resource-picker.html", origin).toString();
}

export function getStoredResourcePageUrl(origin: string): string {
  try {
    const savedValue = window.localStorage.getItem(RESOURCE_URL_STORAGE_KEY)?.trim();
    return savedValue || getDefaultResourcePageUrl(origin);
  } catch (error) {
    return getDefaultResourcePageUrl(origin);
  }
}

export function saveResourcePageUrl(resourceUrl: string): void {
  window.localStorage.setItem(RESOURCE_URL_STORAGE_KEY, resourceUrl.trim());
}

export function resetResourcePageUrl(origin: string): string {
  const defaultUrl = getDefaultResourcePageUrl(origin);
  window.localStorage.setItem(RESOURCE_URL_STORAGE_KEY, defaultUrl);
  return defaultUrl;
}

export function normalizeAbsoluteUrl(rawUrl: string): string {
  return new URL(rawUrl.trim()).toString();
}
