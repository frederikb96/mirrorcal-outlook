import { OUTLOOK_HOSTS } from './core/constants.ts';

/**
 * Public half of the key that fixes the extension id, so the id (and with it the stored settings)
 * does not depend on the folder an unpacked copy is loaded from.
 */
export const EXTENSION_KEY =
  'MIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAsLGQyIzC+KR5yF09A6L4/q2ht1xQe6SwzrRfWJ5/Z/v+jI+wKPkEIPfQt/+QITyDhVWxB1bG+Kyfc1KTz51xN09QJjEibQ5Y8EfY/kC4B11kKiixfvNIw9SIcurLJ+ryAWmAV7bw+TLuk7XjmU//52fWvJDXh4vLYqsGV210gZ9zxKXB/Bz+LMSgKTLGsOaxvoBJdp0XFpxpdDQyqMzp/Aqx5rWFif18U0oNVSnHu1aav1hsTi5NKmrWUFjS/Zt049W/xF4cn2USZmHrC900BG5KzISZoue1zew52/IwTkL9ZRDkuytDB0S06VwBZTLGbmwncO0zzBjqYd4a7Y/nLwIDAQAB';

export function buildManifest(version: string) {
  return {
    manifest_version: 3,
    key: EXTENSION_KEY,
    name: 'MirrorCal Outlook',
    version,
    description: 'Mirrors your Outlook web calendar one-way into a CalDAV calendar.',
    icons: { '16': 'icons/icon-16.png', '48': 'icons/icon-48.png', '128': 'icons/icon-128.png' },
    action: { default_title: 'MirrorCal Outlook settings' },
    options_ui: { page: 'options.html', open_in_tab: true },
    background: { service_worker: 'service-worker.js', type: 'module' },
    permissions: ['alarms', 'storage', 'webRequest'],
    host_permissions: OUTLOOK_HOSTS.map((host) => `https://${host}/*`),
    optional_host_permissions: ['https://*/*'],
  };
}
