// 등록상태 데이터 주소 — 구글 Apps Script 웹 앱 URL (…/exec). 설치 방법은 apps-script/Code.gs 맨 위 참고
// 비어 있으면 등록상태 · 어드민 페이지가 예시 데이터로 보임 (어드민 예시는 아무 비밀번호나 통과, 저장 안 됨)
window.REG_API = 'https://script.google.com/macros/s/AKfycbxAHaz-eRRoPu_zACqwY6VK3HpChs0oTL3Yu3zuQsUlM-00y5Q9DOC1NgX00QHrpWl8/exec';

// 등록 단계 (시트 '등록상태' 열 값과 같아야 함)
window.REG_STATUSES = ['접수', '가등록', '완등록'];

window.REG_DEMO = (function(){
  var names = ['김하늘','이서준','박지민','최예린','정도윤','강민서'];
  var churches = ['세상의 빛 동광교회','구로동교회','새숨교회','컴앤씨교회','주예수교회','복음과도시'];
  var schedules = ['1/21–22 (목, 금) 전일', '1/21 (목) 저녁 집회', '1/22 (금) 저녁 집회'];
  return names.map(function(n, i){
    return {
      row: i + 2,
      time: new Date(2026, 10, 1 + i, 10 + i).toISOString(),
      church: churches[i],
      name: n,
      schedule: schedules[i % 3],
      count: String(8 + i * 5),
      phone: '010-1234-' + String(1000 + i * 1379).slice(-4),
      status: window.REG_STATUSES[i % 3]
    };
  });
})();
