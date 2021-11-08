self.addEventListener('install', function (event) {
    self.skipWaiting();
});

self.addEventListener('activate', function (event) {
    event.waitUntil(self.clients.claim());
});

self.addEventListener('notificationclick', function (event) {
    const id = event.action.split('.')[1], pwd = event.action.split('.')[2], teacher = event.action.split('.')[3]
    self.registration.showNotification(`${teacher} 선생님 줌에 접속했어요.`, {
        body: `암호는 ${pwd}에요.`
    })
    self.clients.openWindow(`zoommtg://zoom.us/join?confno=` + id)
});
