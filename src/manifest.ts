import { OUTLOOK_HOSTS } from './core/constants.ts';

export function buildManifest(version: string) {
  return {
    manifest_version: 3,
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
