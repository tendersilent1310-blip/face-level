// 설치(홈 화면 앱)용 최소 서비스워커. 캐시하지 않음: 항상 최신 버전을 받음
self.addEventListener('install',e=>self.skipWaiting());
self.addEventListener('activate',e=>e.waitUntil(self.clients.claim()));
self.addEventListener('fetch',()=>{});
