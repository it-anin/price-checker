(function(){
 var modal=document.querySelector('#design-preview'),frame=document.querySelector('#design-frame'),current=1,timer,opener;
 function show(id){
  current=(id+19)%20+1;
  var d=DESIGNS[current-1],path='concepts/'+String(current).padStart(2,'0')+'.html';
  frame.src=path;document.querySelector('#full-view').href=path;
  document.querySelector('#preview-title').textContent=d.name;
  document.querySelector('#preview-number').textContent=String(current).padStart(2,'0')+' / '+d.thai;
  document.querySelector('#preview-description').textContent=d.description;
  document.querySelector('#preview-counter').textContent=String(current).padStart(2,'0')+' / 20';
  if(!modal.open){modal.showModal();document.body.style.overflow='hidden'}
 }
 document.addEventListener('click',function(e){var b=e.target.closest('[data-open]');if(b){opener=b;show(Number(b.dataset.open))}});
 function close(){modal.close()}
 document.querySelector('#close-preview').onclick=close;
 modal.addEventListener('click',function(e){if(e.target===modal){var r=modal.getBoundingClientRect();if(e.clientX<r.left||e.clientX>r.right||e.clientY<r.top||e.clientY>r.bottom)close()}});
 modal.addEventListener('close',function(){document.body.style.overflow='';frame.src='about:blank';if(opener)opener.focus()});
 document.querySelector('#previous-design').onclick=function(){show(current-1)};
 document.querySelector('#next-design').onclick=function(){show(current+1)};
 document.addEventListener('keydown',function(e){if(modal.open&&!/INPUT|TEXTAREA/.test(e.target.tagName)){if(e.key==='ArrowLeft'){e.preventDefault();show(current-1)}if(e.key==='ArrowRight'){e.preventDefault();show(current+1)}}});
 function view(mobile){document.querySelector('.preview-stage').classList.toggle('mobile',mobile);document.querySelector('#desktop-view').classList.toggle('selected',!mobile);document.querySelector('#mobile-view').classList.toggle('selected',mobile)}
 document.querySelector('#desktop-view').onclick=function(){view(false)};
 document.querySelector('#mobile-view').onclick=function(){view(true)};
 document.querySelector('#design-search').addEventListener('input',function(){
  var q=this.value.toLowerCase().trim(),count=0;
  document.querySelectorAll('.design-card').forEach(function(card){var d=DESIGNS[Number(card.dataset.id)-1],text=String(d.id).padStart(2,'0')+' '+d.name+' '+d.thai+' '+d.description+' '+d.kind;card.hidden=!!q&&!text.toLowerCase().includes(q);if(!card.hidden)count++});
  document.querySelector('#result-count').textContent=count+' แบบ';
 });
 document.querySelector('#choose-design').onclick=async function(){
  var d=DESIGNS[current-1],text='เลือกแบบ '+String(current).padStart(2,'0')+' — '+d.name+' / '+d.thai;
  try{await navigator.clipboard.writeText(text)}catch(e){}
  var toast=document.querySelector('#gallery-toast');toast.textContent=text;toast.classList.add('visible');clearTimeout(timer);timer=setTimeout(function(){toast.classList.remove('visible')},4000);
 };
 var query=Number(new URLSearchParams(location.search).get('design'));if(Number.isInteger(query)&&query>=1&&query<=20)show(query);
})();
