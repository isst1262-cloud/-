// 최소한의 서비스 워커: PWA 설치(홈 화면 추가) 가능하게만 함. 오프라인 캐싱은 하지 않음
// (데이터가 항상 최신이어야 하는 스크리너 특성상 캐시로 인한 오래된 데이터 노출을 피함).
self.addEventListener("install", () => self.skipWaiting());
self.addEventListener("activate", (e) => e.waitUntil(self.clients.claim()));
self.addEventListener("fetch", () => {});
