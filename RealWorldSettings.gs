const DB23=Object.freeze({
  SHIFT_DEF:'DEFINICJE_ZMIAN',
  STAFF:'MACIERZ_OBSADY',
  FUNCTIONS:'FUNKCJE_DODATKOWE',
  CALENDAR:'KALENDARZ_MODYFIKACJI'
});

function dbInstallRealWorldSettings_(){
  const defs={
    [DB23.SHIFT_DEF]:['LOKALIZACJA_ID','GRUPA_DNI','DZIEŃ_TYGODNIA','ZMIANA_ID','NAZWA','START','KONIEC','KONIEC_DZIEŃ_PLUS','AKTYWNA'],
    [DB23.STAFF]:['LOKALIZACJA_ID','GRUPA_DNI','ZMIANA_ID','ROLA','FUNKCJA_WYMAGANA','MIN_OSÓB','OPTYMALNIE_OSÓB','MAX_OSÓB','AKTYWNA'],
    [DB23.FUNCTIONS]:['KOD','NAZWA','ROLA_WYMAGANA','LOKALIZACJA','TYP_PRZYDZIAŁU','AKTYWNA','OPIS'],
    [DB23.CALENDAR]:['ID','DATA','LOKALIZACJA_ID','TYP','ZMIANA_ID','ROLA','WARTOŚĆ','START','KONIEC','DZIEŃ_PLUS','NAZWA','UWAGI','AKTYWNY']
  };
  const ss=SpreadsheetApp.getActive();
  Object.keys(defs).forEach(name=>{
    let sh=ss.getSheetByName(name);if(!sh)sh=ss.insertSheet(name);
    const h=defs[name];if(sh.getMaxColumns()<h.length)sh.insertColumnsAfter(sh.getMaxColumns(),h.length-sh.getMaxColumns());
    sh.getRange(1,1,1,h.length).setValues([h]).setBackground('#172554').setFontColor('#fff').setFontWeight('bold').setWrap(true);
    sh.setFrozenRows(1);sh.setHiddenGridlines(true);sh.autoResizeColumns(1,h.length);
  });
  db23WriteIfEmpty_(DB23.SHIFT_DEF,db23ShiftDefinitions_());
  db23WriteIfEmpty_(DB23.STAFF,db23Staffing_());
  db23WriteIfEmpty_(DB23.FUNCTIONS,[
    ['HOST','Host','KELNER','KRUCZA','ZMIANA','TAK','Rola krążąca, liczona oddzielnie od kelnera'],
    ['ZAMKNIĘCIE_BARU','Zamyka bar','BARMAN','WSZYSTKIE','WYMÓG_ZMIANY','TAK','Wymagana na każdej wieczornej zmianie'],
    ['ZAMKNIĘCIE_SALI','Zamyka salę','KELNER','KRUCZA','WYMÓG_ZMIANY','TAK','Konfigurowalny wymóg wieczorny'],
    ['MENADŻER_ZESPOŁU','Menadżer zespołu','DOWOLNA','WG_MATRYCY','UPRAWNIENIE','TAK','Nadzoruje grafik swojej roli'],
    ['EVENT_ROTACYJNY','Rotacyjny event','BARMAN','KRUCZA,PAWILONY','SEGMENT','TAK','Praca w segmentach w obu lokalach']
  ]);
}

function dbLoadDemo(){
  dbInstall();
  const rows=[],first=['Anna','Marta','Julia','Zofia','Aleksandra','Natalia','Katarzyna','Iga','Lena','Maja','Monika','Karolina','Ewa','Alicja','Weronika','Joanna','Paulina','Magdalena','Dominika','Agata'];
  const last=['Nowak','Kowalska','Wiśniewska','Wójcik','Kamińska','Lewandowska','Zielińska','Szymańska','Woźniak','Dąbrowska','Kozłowska','Jankowska','Mazur','Krawczyk','Piotrowska','Grabowska','Pawłowska','Michalska','Król','Wieczorek'];
  const roles=[['KELNER',28],['BARMAN',19],['PIZZABAR',18],['PREP',7],['POMOC',4]];let i=0;
  roles.forEach(([role,count])=>{for(let n=0;n<count;n++,i++){
    const manager=n===0||(role==='BARMAN'&&n===1),base=role==='BARMAN'?(n===1||n%5===0?'PAWILONY':'KRUCZA'):(role==='PIZZABAR'&&n%4===0?'PAWILONY':'KRUCZA');
    const rotating=(role==='BARMAN'||role==='PIZZABAR')&&n>=Math.max(2,count-5),kr=base==='KRUCZA'||rotating,paw=(base==='PAWILONY'||rotating)&&(role==='BARMAN'||role==='PIZZABAR');
    rows.push([
      `P${String(i+1).padStart(3,'0')}`,'TAK',`${first[i%20]} ${last[(Math.floor(i/20)*5+i%5)%20]}`,`${role.toLowerCase()}.${String(n+1).padStart(2,'0')}@demo.pl`,`500${String(100000+i).slice(-6)}`,
      role,base,kr?'TAK':'NIE',paw?'TAK':'NIE','TAK',(role==='BARMAN'||role==='PIZZABAR')?'TAK':'NIE',
      role==='KELNER'&&n%4===0?'TAK':'NIE',role==='BARMAN'&&(manager||n%3===0)?'TAK':'NIE',role==='KELNER'&&(manager||n%6===0)?'TAK':'NIE',
      manager?'TAK':'NIE',manager?role:'',manager?(role==='BARMAN'?(n===0?'KRUCZA':'PAWILONY'):'WSZYSTKIE'):'',
      rotating?'TAK':'NIE',rotating?'TAK':'NIE',(rotating||n%4===0)?'TAK':'NIE',n%5===0?'TAK':'NIE',
      n<count-3?'UMOWA O PRACĘ':'CZĘŚĆ ETATU',n<count-3?1:.5,n<count-3?168:84,'2026-01-01','',
      role==='PREP'?'TAK':'NIE','NIE','NIE',6,48,11,manager?3:1,manager?'KIEROWNIK':'PRACOWNIK',
      `KAD-${String(i+1).padStart(4,'0')}`,'Dane demonstracyjne','GOTOWY'
    ]);
  }});
  const sh=SpreadsheetApp.getActive().getSheetByName(DB.EMP);
  const body=sh.getRange(2,1,Math.max(1,sh.getMaxRows()-1),sh.getMaxColumns());
  body.clearDataValidations();
  body.clearContent();
  if(sh.getMaxColumns()>EMP_HEADERS.length){
    sh.getRange(1,EMP_HEADERS.length+1,sh.getMaxRows(),sh.getMaxColumns()-EMP_HEADERS.length)
      .clearContent().clearDataValidations();
  }
  sh.getRange(2,1,rows.length,EMP_HEADERS.length).setValues(rows);
  dbApplyValidations_();db23ApplyCheckboxes_();dbValidate23_();dbTouchVersion();
  return {ok:true,employees:rows.length};
}

function db23ShiftDefinitions_(){
  const rows=[],add=(loc,group,days,id,name,start,end,plus)=>days.forEach(day=>rows.push([loc,group,day,id,name,start,end,plus,'TAK']));
  add('KRUCZA','PON-CZW',['PON','WT','ŚR','CZW'],'RANO','Poranna','10:00','17:00',0);
  add('KRUCZA','PON-CZW',['PON','WT','ŚR','CZW'],'WIECZÓR','Wieczorna','17:00','01:00',1);
  add('KRUCZA','PT-ND',['PT','SOB','ND'],'RANO','Poranna','10:00','17:00',0);add('KRUCZA','PT-ND',['PT','SOB','ND'],'ŚRODEK','Środkowa','15:00','23:00',0);add('KRUCZA','PT-ND',['PT','SOB','ND'],'WIECZÓR','Wieczorna','17:00','03:00',1);
  add('PAWILONY','ND-CZW',['ND','PON','WT','ŚR','CZW'],'RANO','Poranna','10:00','17:00',0);add('PAWILONY','ND-CZW',['ND','PON','WT','ŚR','CZW'],'WIECZÓR','Wieczorna','17:00','01:00',1);
  add('PAWILONY','PT-SOB',['PT','SOB'],'RANO','Poranna','12:00','19:00',0);add('PAWILONY','PT-SOB',['PT','SOB'],'WIECZÓR','Wieczorna','19:00','05:00',1);return rows;
}
function db23Staffing_(){
  const r=[],a=(l,g,s,role,n,f)=>r.push([l,g,s,role,f||'',n,n,n,'TAK']);
  [['PON-CZW','RANO',{KELNER:4,BARMAN:2,PIZZABAR:2,PREP:5,POMOC:1}],['PON-CZW','WIECZÓR',{KELNER:8,BARMAN:3,PIZZABAR:4,POMOC:1}],['PT-ND','RANO',{KELNER:6,BARMAN:3,PIZZABAR:3,PREP:5,POMOC:1}],['PT-ND','ŚRODEK',{KELNER:2,BARMAN:1}],['PT-ND','WIECZÓR',{KELNER:8,BARMAN:5,PIZZABAR:5,POMOC:2}]].forEach(x=>Object.keys(x[2]).forEach(role=>a('KRUCZA',x[0],x[1],role,x[2][role],x[1]==='WIECZÓR'&&role==='BARMAN'?'ZAMKNIĘCIE_BARU':x[1]==='WIECZÓR'&&role==='KELNER'?'ZAMKNIĘCIE_SALI':'')));
  ['PON-CZW','PT-ND'].forEach(g=>{a('KRUCZA',g,'RANO','KELNER',1,'HOST');a('KRUCZA',g,'WIECZÓR','KELNER',1,'HOST');});
  ['ND-CZW','PT-SOB'].forEach(g=>{a('PAWILONY',g,'RANO','BARMAN',1,'');a('PAWILONY',g,'RANO','PIZZABAR',1,'');a('PAWILONY',g,'WIECZÓR','BARMAN',2,'ZAMKNIĘCIE_BARU');a('PAWILONY',g,'WIECZÓR','PIZZABAR',2,'');});return r;
}
function db23WriteIfEmpty_(name,rows){const sh=SpreadsheetApp.getActive().getSheetByName(name);if(sh.getLastRow()===1&&rows.length)sh.getRange(2,1,rows.length,rows[0].length).setValues(rows);}
function db23ApplyCheckboxes_(){const sh=SpreadsheetApp.getActive().getSheetByName(DB.EMP),n=Math.max(1,sh.getMaxRows()-1);[2,8,9,10,11,12,13,14,15,18,19,20,21,27,28,29].forEach(c=>sh.getRange(2,c,n,1).insertCheckboxes('TAK','NIE'));}
function dbValidate23_(){
  const sh=SpreadsheetApp.getActive().getSheetByName(DB.EMP),data=sh.getDataRange().getValues(),h=data.shift(),idx=Object.fromEntries(h.map((x,i)=>[x,i])),errors=[],names=new Set(),ids=new Set(),managerCounts={};
  data.filter(r=>r.some(Boolean)).forEach((r,i)=>{
    const row=i+2,id=r[idx['PRACOWNIK_ID*']],name=r[idx['IMIĘ_I_NAZWISKO*']],role=r[idx['ROLA_GŁÓWNA*']];
    if(ids.has(id))errors.push(`Wiersz ${row}: powtórzone ID ${id}`);ids.add(id);if(names.has(name))errors.push(`Wiersz ${row}: powtórzone imię i nazwisko ${name}`);names.add(name);
    if(r[idx.HOST]==='TAK'&&role!=='KELNER')errors.push(`Wiersz ${row}: HOST tylko dla KELNER`);
    if(r[idx.ZAMKNIĘCIE_BARU]==='TAK'&&role!=='BARMAN')errors.push(`Wiersz ${row}: ZAMKNIĘCIE_BARU tylko dla BARMAN`);
    if(r[idx.ZAMKNIĘCIE_SALI]==='TAK'&&role!=='KELNER')errors.push(`Wiersz ${row}: ZAMKNIĘCIE_SALI tylko dla KELNER`);
    if(role==='PREP'&&(r[idx.PAWILONY_STANDARD]==='TAK'||r[idx.TYLKO_RANO]!=='TAK'))errors.push(`Wiersz ${row}: PREP tylko KRUCZA i rano`);
    if(r[idx.MENADŻER_ZESPOŁU]==='TAK'){const key=role==='BARMAN'?`${role}|${r[idx.ZARZĄDZA_LOKALIZACJĄ]}`:role;managerCounts[key]=(managerCounts[key]||0)+1;}
  });
  ['KELNER','PIZZABAR','PREP','POMOC','BARMAN|KRUCZA','BARMAN|PAWILONY'].forEach(k=>{if(managerCounts[k]!==1)errors.push(`${k}: wymagany dokładnie 1 menadżer, znaleziono ${managerCounts[k]||0}`);});
  SpreadsheetApp.getActive().toast(errors.length?`Błędy: ${errors.length}`:'Baza 2.3 jest poprawna.','Walidacja',7);return {ok:!errors.length,errors};
}
function dbValidate(){return dbValidate23_();}
