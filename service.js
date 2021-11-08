self.addEventListener('install', function (event) {
    self.skipWaiting();
});

self.addEventListener('activate', function (event) {
    event.waitUntil(self.clients.claim());
});

self.addEventListener('notificationclick', function (event) {
    if (event.action.split('.')[0] !== 'zoom') return
    const id = event.action.split('.')[1], pwd = event.action.split('.')[2], teacher = event.action.split('.')[3]
    if (!id || !pwd) return
    self.registration.showNotification(`${teacher} 선생님 줌에 접속했어요.`, {
        body: `암호는 ${pwd}에요.`
    })
    self.clients.openWindow(`zoommtg://zoom.us/join?confno=` + id)
});
