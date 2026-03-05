/* ============================================================
   TASKFLOW — Service Worker (sw.js)
   Handles push notifications and background sync
   ============================================================ */

const CACHE_NAME = 'taskflow-v1';

// Listen for push events (from server-side push)
self.addEventListener('push', function (event) {
    let data = { title: 'TaskFlow', body: 'You have a new notification.' };
    if (event.data) {
        try {
            data = event.data.json();
        } catch (e) {
            data.body = event.data.text();
        }
    }
    const options = {
        body: data.body || '',
        icon: '/favicon.ico',
        badge: '/favicon.ico',
        tag: data.tag || 'taskflow-notif',
        vibrate: [100, 50, 100],
        data: { url: data.url || '/' },
        actions: [
            { action: 'open', title: 'Open TaskFlow' },
            { action: 'dismiss', title: 'Dismiss' }
        ]
    };
    event.waitUntil(
        self.registration.showNotification(data.title || 'TaskFlow', options)
    );
});

// Handle notification click — open app
self.addEventListener('notificationclick', function (event) {
    event.notification.close();
    if (event.action === 'dismiss') return;
    event.waitUntil(
        clients.matchAll({ type: 'window', includeUncontrolled: true }).then(function (clientList) {
            // Focus existing window if available
            for (var i = 0; i < clientList.length; i++) {
                var client = clientList[i];
                if (client.url.includes(self.location.origin) && 'focus' in client) {
                    return client.focus();
                }
            }
            // Otherwise open new window
            if (clients.openWindow) {
                return clients.openWindow(event.notification.data?.url || '/');
            }
        })
    );
});

// Handle messages from main thread (show notification from app)
self.addEventListener('message', function (event) {
    if (event.data && event.data.type === 'SHOW_NOTIFICATION') {
        var options = {
            body: event.data.body || '',
            icon: '/favicon.ico',
            badge: '/favicon.ico',
            tag: event.data.tag || 'taskflow-notif',
            vibrate: [100, 50, 100],
            data: { url: '/' }
        };
        self.registration.showNotification(event.data.title || 'TaskFlow', options);
    }
});

// Activate immediately
self.addEventListener('activate', function (event) {
    event.waitUntil(self.clients.claim());
});

// Install — skip waiting for immediate activation
self.addEventListener('install', function (event) {
    self.skipWaiting();
});
