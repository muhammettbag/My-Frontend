export const API_BASE_URL = "https://testmb.runasp.net/api";

export const apiUrl = (path) => {
  if (!path) return API_BASE_URL
  const normalizedPath = path.startsWith("/") ? path : `/${path}`
  return `${API_BASE_URL}${normalizedPath}`
}
