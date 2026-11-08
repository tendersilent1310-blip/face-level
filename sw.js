// 설치(홈 화면 앱)용 서비스워커. 화면(HTML)은 항상 네트워크에서 최신으로 받음(폰 캐시 무시). 인터넷이 없을 때만 저장본 사용
self.addEventListener('install',e=>self.skipWaiting());
self.addEventListener('activate',e=>e.waitUntil(self.clients.claim()));
self.addEventListener('fetch',e=>{
  const r=e.request;
  if(r.method==='GET'&&(r.mode==='navigate'||/\.html(\?|$)/.test(r.url))){
    e.respondWith(fetch(r,{cache:'no-store'}).catch(()=>fetch(r)));
  }
});
