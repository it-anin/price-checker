(function(){
 var toast=document.querySelector('.toast'),timer;
 function notify(text){toast.textContent=text;toast.classList.add('visible');clearTimeout(timer);timer=setTimeout(function(){toast.classList.remove('visible')},2400)}
 document.addEventListener('click',function(event){
  var tab=event.target.closest('[data-tab]');
  if(tab){document.querySelectorAll('[data-tab]').forEach(function(b){b.setAttribute('aria-selected',String(b===tab))});document.querySelectorAll('[data-panel]').forEach(function(p){p.hidden=p.dataset.panel!==tab.dataset.tab});return}
  var branch=event.target.closest('[data-branch]');
  if(branch){document.querySelectorAll('[data-branch]').forEach(function(b){b.setAttribute('aria-selected',String(b===branch))});document.querySelector('#branch-name').textContent=branch.dataset.branch;return}
  var inspect=event.target.closest('[data-inspect]');
  if(inspect){document.querySelectorAll('[data-inspect]').forEach(function(b){b.classList.toggle('selected',b===inspect)});document.querySelector('#inspect-title').textContent=inspect.dataset.inspect;return}
  var module=event.target.closest('[data-module]');
  if(module){document.querySelectorAll('[data-module]').forEach(function(b){b.classList.toggle('selected',b===module)});document.querySelectorAll('[data-module-panel]').forEach(function(p){p.hidden=p.dataset.modulePanel!==module.dataset.module});return}
  var demo=event.target.closest('[data-demo]');
  if(demo){notify('ตัวอย่าง: '+demo.dataset.demo+' · ยังไม่เปลี่ยนข้อมูลจริง')}
 });
 var input=document.querySelector('.search input');
 if(input){input.addEventListener('input',function(){
  var query=input.value.trim().toLowerCase(),items=document.querySelectorAll('.group .action,.palette .action,.directory-row,.launch-tile,.inspect-item,.console-grid .action');
  items.forEach(function(item){item.hidden=!!query&&!item.textContent.toLowerCase().includes(query)})
 })}
 document.addEventListener('keydown',function(e){
  if(e.key==='/'&&!/INPUT|TEXTAREA/.test(document.activeElement.tagName)&&input){e.preventDefault();input.focus()}
  if(e.target.matches('[role=tab]')&&['ArrowLeft','ArrowRight','Home','End'].includes(e.key)){
   e.preventDefault();var items=Array.from(e.target.parentElement.querySelectorAll('[role=tab]')),i=items.indexOf(e.target),next=e.key==='Home'?0:e.key==='End'?items.length-1:(i+(e.key==='ArrowRight'?1:items.length-1))%items.length;items[next].click();items[next].focus()
  }
 });
})();
