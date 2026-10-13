const steps=[
{title:'완료할 업무를 요청합니다',tab:'목표 전달',image:1,text:'공식 회의록 사이트와 수집할 연도, 저장할 위치를 전달합니다. “어떻게 하나요?”에서 “이 업무를 완료해줘”로 요청을 구체화합니다.',check:'수집 범위·기준일·저장 위치를 정했는가?',caption:'GPTWork_1 · 사이트와 로컬 저장을 요청한 화면'},
{title:'사이트를 확인하고 수집합니다',tab:'브라우저 수집',image:2,extra:[3],text:'브라우저에서 목록과 다운로드 방식을 확인합니다. 원본을 받을 수 없는 자료는 공개 본문 사본과 구분해 기록했습니다.',check:'의회 대수 변경과 누락·다운로드 제한을 확인했는가?',caption:'GPTWork_2 · 공개 회의록 목록 확인'},
{title:'자료를 로컬 폴더에 저장합니다',tab:'로컬 저장',image:4,extra:[5,3],text:'HWP 원본과 공개 본문 사본을 나눠 저장하고, 출처·저장 상태를 회의목록.csv에 남겼습니다. 파일이 저장됐는지 직접 확인합니다.',check:'18건의 원본과 7건의 사본을 구분했는가?',caption:'GPTWork_4 · 원본·사본·목록이 정리된 폴더'},
{title:'저장한 자료를 이어서 분석합니다',tab:'분석 요청',image:6,text:'로컬에 저장한 회의록을 바탕으로 의원별 지적사항을 정리하고 Excel 파일로 저장하도록 요청합니다. 날짜와 발언 근거, 반복 여부를 함께 확인합니다.',check:'의혹·제안·사실을 구분하고 원문 근거를 요구했는가?',caption:'GPTWork_6 · 의원별 분석을 요청한 실제 화면'},
{title:'Excel 결과와 근거를 확인합니다',tab:'결과 검토',image:9,extra:[7,8],text:'의원별 요약·지적사항 상세·자료범위와 분석기준을 확인합니다. 생성된 표는 초안으로 검토하고, 필요한 내용은 원문과 대조합니다.',check:'130개는 분석 항목 수입니다. 활동 평가·순위로 해석하지 않았는가?',caption:'GPTWork_9 · Excel에서 열린 분석 결과'}];
const prompts=[`https://ams.sdmcouncil.go.kr/assem/search/simple/council.do?searchCodeType=B

서대문구의회 사이트에서 2026년 본회의 회의록을 확인하고,
수집 시점까지 공개된 자료를 연결된 로컬 폴더에 저장해줘.
개회식 포함 여부를 구분하고 의회 대수가 바뀌는 구간도 확인해줘.
원본 다운로드가 가능하면 원본을 저장해줘.
제한되거나 첨부가 없으면 제한을 우회하지 말고,
접근 가능한 공개 본문을 사본으로 보존해줘.
회의일·회기·출처 URL·저장 상태·파일 경로를 회의목록.csv에 기록하고,
수집 기준일과 누락·제한 사항을 알려줘.`,
`저장한 회의록을 근거로 의원별 지적사항과 개선 요구를 분석해 XLSX로 저장해줘.
의원별 요약 / 지적사항 상세 / 자료범위·분석기준의 3개 시트로 구성해줘.
상세 시트에 회의 날짜·회기·발언자·주제·요약·개선 요구·집행부 답변·원문 발췌·출처를 넣어줘.
여러 날짜에서 같은 사안이 확인될 때만 반복 제기로 표시해줘.
의혹은 발언자의 주장으로 구분하고, 검토 약속을 조치 완료로 해석하지 마.
본문에 없는 서면질문은 확인 불가로 표시해줘.
항목 수를 의원 활동 평가나 순위로 사용하지 마.
분석 기간·읽지 못한 파일·제외 자료·원본과 임시자료 범위를 명시해줘.`];
const cases=[
['▦','여러 엑셀 취합·분류','부서별 교육용 실적표와 기존 분류 예시','표 형식 통합, 분류·누락·합계 확인','통합 XLSX · 요약 차트','교육용 부서별 실적 엑셀을 하나로 합치고 지난 분류 예시를 참고해 정리해줘. 애매한 분류와 누락은 확인 대상으로 표시하고 합계를 검증한 XLSX로 저장해줘.'],
['⌕','타 지자체 정책 사례 조사','공식 사이트와 비교 기준','대상·기간·내용·신청 경로·출처 정리','정책 비교 XLSX','지정한 지자체 공식 사이트에서 청년 지원 정책을 조사해 대상·기간·내용·신청 경로·출처 URL·확인일을 비교표로 정리해줘. 게시일과 신청 기간은 구분해줘.'],
['▤','회의자료로 발표자료 작성','교육용 회의 메모·실적표·기관 양식','핵심 내용과 할 일을 발표 순서로 구성','PPTX 초안','제공한 교육용 회의 메모와 실적표를 근거로 기관 PPT 양식에 맞춘 5장 발표자료를 만들어줘. 확인되지 않은 사실과 제안을 구분하고 출처를 남겨줘.'],
['✦','행사 홍보물 변형 제작','기존 이미지와 확정된 일정·문구','기존 분위기를 참고해 행사별 이미지 수정','홍보 이미지 초안','기존 행사 홍보 이미지의 분위기를 참고해 제공한 확정 일정과 문구로 새 홍보물 초안을 만들어줘. 날짜·장소·담당 연락처를 임의로 만들지 마.'],
['▱','인수인계 문서 정리','교육용 복사 폴더와 분류 기준','업무·연도별 분류, 이름과 목록 정리','정리한 사본 폴더 · 파일목록','교육용 복사 폴더를 업무·연도별로 분류하는 계획을 먼저 보여줘. 원본은 그대로 두고 사본을 정리해줘. 분류가 애매한 파일은 별도 목록으로 남겨줘.'],
['◷','공개자료 정기 조사','공식 사이트·조사 기준·주기','새 공고와 변경 사항을 주기적으로 확인','반복 조사 결과','검토한 공개 공고 조사 흐름을 매주 월요일 오전 8시 한국시간 기준으로 반복하도록 설정해줘. 변경 사항과 출처를 알려주고 사용할 수 없는 자료는 확인 불가로 표시해줘.']];
let activeStep=0,activePrompt=0;const $=s=>document.querySelector(s);const tabs=$('.demo-tabs');
steps.forEach((s,i)=>{const b=document.createElement('button');b.id=`tab-${i}`;b.setAttribute('role','tab');b.setAttribute('aria-controls','stepPanel');b.innerHTML=`<span>0${i+1}</span>${s.tab}`;b.addEventListener('click',()=>setStep(i));b.addEventListener('keydown',e=>{if(['ArrowRight','ArrowLeft','Home','End'].includes(e.key)){e.preventDefault();const next=e.key==='Home'?0:e.key==='End'?4:(i+(e.key==='ArrowRight'?1:4))%5;setStep(next);tabs.children[next].focus()}});tabs.append(b)});
function openImage(src,caption){$('#modalImage').src=src;$('#modalImage').alt=caption;$('#modalCaption').textContent=caption;$('#imageModal').showModal()}
function setStep(i){activeStep=(i+5)%5;const s=steps[activeStep];[...tabs.children].forEach((b,j)=>{b.setAttribute('aria-selected',j===activeStep);b.tabIndex=j===activeStep?0:-1});$('#stepPanel').setAttribute('aria-labelledby',`tab-${activeStep}`);$('#stepNumber').textContent=`STEP 0${activeStep+1}`;$('#stepTitle').textContent=s.title;$('#stepText').textContent=s.text;$('#stepCheck').textContent=s.check;$('#stepCount').textContent=`${activeStep+1} / 5`;$('#stepImage').src=`assets/step-${s.image}.webp`;$('#stepImage').alt=s.caption;$('#stepCaption').textContent=s.caption;$('#stepImageButton').onclick=()=>openImage(`assets/step-${s.image}.webp`,s.caption);$('#extraImages').replaceChildren();(s.extra||[]).forEach(n=>{const b=document.createElement('button');b.className='extra-image';b.textContent=({3:'수집 상태 기록',5:'HWP 원본 목록',7:'앱에서 결과 보기',8:'저장한 파일 확인'})[n];b.onclick=()=>openImage(`assets/step-${n}.webp`,b.textContent);$('#extraImages').append(b)});$('#stepPanel').classList.remove('enter');requestAnimationFrame(()=>$('#stepPanel').classList.add('enter'))}
$('#previous').onclick=()=>setStep(activeStep-1);$('#next').onclick=()=>setStep(activeStep+1);setStep(0);
document.querySelectorAll('[data-image]').forEach(b=>b.onclick=()=>openImage(b.dataset.image,b.dataset.caption));$('#closeModal').onclick=()=>$('#imageModal').close();$('#imageModal').addEventListener('click',e=>{if(e.target===$('#imageModal')){const r=e.target.getBoundingClientRect();if(e.clientX<r.left||e.clientX>r.right||e.clientY<r.top||e.clientY>r.bottom)e.target.close()}});
function setPrompt(i){activePrompt=i;$('#promptText').textContent=prompts[i];document.querySelectorAll('[data-prompt]').forEach(b=>b.setAttribute('aria-pressed',Number(b.dataset.prompt)===i));$('#copyStatus').textContent=''}document.querySelectorAll('[data-prompt]').forEach(b=>b.onclick=()=>setPrompt(Number(b.dataset.prompt)));setPrompt(0);
async function copy(text){try{await navigator.clipboard.writeText(text);return true}catch{const t=document.createElement('textarea');t.value=text;t.style.position='fixed';t.style.opacity='0';document.body.append(t);t.select();const ok=document.execCommand('copy');t.remove();return ok}}
$('#copyPrompt').onclick=async()=>{$('#copyStatus').textContent=(await copy(prompts[activePrompt]))?'요청문을 복사했습니다.':'복사할 수 없습니다. 요청문을 직접 선택해 주세요.'};
cases.forEach((c,i)=>{const a=document.createElement('article');a.className='case-card';a.innerHTML=`<div class="case-icon" aria-hidden="true">${c[0]}</div><h3>${c[1]}</h3><dl><dt>자료</dt><dd>${c[2]}</dd><dt>맡길 업무</dt><dd>${c[3]}</dd><dt>결과</dt><dd>${c[4]}</dd></dl><details><summary>요청 예시 보기</summary><p>${c[5]}</p>${i===5?'<p>예약·실행 환경과 필요한 도구를 확인합니다. 이 페이지는 예약을 생성하지 않습니다.</p>':''}<button>요청문 복사</button></details>`;a.querySelector('button').onclick=async()=>{const ok=await copy(c[5]);$('#toast').textContent=ok?'요청문을 복사했습니다.':'복사할 수 없습니다. 직접 선택해 주세요.';$('#toast').classList.add('show');setTimeout(()=>$('#toast').classList.remove('show'),2500)};$('#caseGrid').append(a)});
document.querySelectorAll('[data-answer]').forEach(b=>b.onclick=()=>{$('#quizResult').textContent=b.dataset.answer});
let paused=matchMedia('(prefers-reduced-motion: reduce)').matches;function motion(){document.body.classList.toggle('no-motion',paused);$('#motion').textContent=paused?'움직임 켜기':'움직임 끄기';$('#motion').setAttribute('aria-pressed',paused)}motion();$('#motion').onclick=()=>{paused=!paused;motion()};
document.querySelectorAll('.tilt').forEach(c=>{c.addEventListener('pointermove',e=>{if(paused||e.pointerType==='touch')return;const r=c.getBoundingClientRect();c.style.transform=`perspective(900px) rotateX(${-(e.clientY-r.top-r.height/2)/70}deg) rotateY(${(e.clientX-r.left-r.width/2)/70}deg)`});c.addEventListener('pointerleave',()=>c.style.transform='')});
function progress(){const h=document.documentElement.scrollHeight-innerHeight;$('.reading').style.width=`${h>0?scrollY/h*100:0}%`}addEventListener('scroll',progress,{passive:true});addEventListener('resize',progress);progress();
