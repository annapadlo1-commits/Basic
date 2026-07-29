const DB = Object.freeze({
  VERSION:'2.2.1-DEMO-INSTALL-FIX',
  EMP:'BAZA_PRACOWNIKÓW',
  LOC:'LOKALIZACJE',
  SHIFT:'TYPY_ZMIAN',
  RULES:'REGUŁY_PLANOWANIA',
  SCENARIOS:'SCENARIUSZE',
  MODES:'TRYBY_OPTYMALIZACJI',
  LEVELS:'POZIOMY_OBSADY',
  SKILLS:'SŁOWNIK_KOMPETENCJI',
  GUIDE:'START'
});

const EMP_HEADERS = [
  'PRACOWNIK_ID*','AKTYWNY*','IMIĘ_I_NAZWISKO*','EMAIL*','TELEFON',
  'LOKALIZACJA_DOMYŚLNA*','DOZWOLONE_LOKALIZACJE*','PRZEŁOŻONY_EMAIL',
  'TYP_UMOWY*','ETAT*','GODZINY_MIESIĘCZNE*','DATA_ZATRUDNIENIA_OD*','DATA_ZATRUDNIENIA_DO',
  'TYLKO_RANO','TYLKO_POPOŁUDNIE','BEZ_WEEKENDÓW','DOSTĘPNY_STANDBY',
  'MAX_DNI_Z_RZĘDU','MAX_GODZIN_TYGODNIOWO','MIN_ODPOCZYNEK_H',
  'PREFEROWANE_DNI','NIEPREFEROWANE_DNI','PREFEROWANE_ZMIANY','PREFEROWANE_LOKALIZACJE',
  'KOMPETENCJE*','MOŻE_OTWIERAĆ','MOŻE_ZAMYKAĆ','MOŻE_BYĆ_LIDEREM',
  'PRIORYTET_PLANOWANIA','ROLA_APLIKACJI*','ID_KADROMIERZ','UWAGI','STATUS_REKORDU'
];

function onOpen(){
  SpreadsheetApp.getUi().createMenu('GRAFIK PRO — BAZA')
    .addItem('Instaluj strukturę','dbInstall')
    .addItem('Załaduj 60 pracowników DEMO','dbLoadDemo')
    .addSeparator()
    .addItem('Sprawdź bazę','dbValidate')
    .addItem('Oznacz zmiany do synchronizacji','dbTouchVersion')
    .addToUi();
}

function dbInstall(){
  const ss=SpreadsheetApp.getActive();
  const defs={
    [DB.EMP]:EMP_HEADERS,
    [DB.LOC]:['LOKALIZACJA_ID*','NAZWA*','ADRES','AKTYWNA*','KIEROWNIK_EMAIL','KOLOR','STREFA'],
    [DB.SHIFT]:['ZMIANA_ID*','NAZWA*','START*','KONIEC*','PŁATNE_GODZINY*','TYP*','KOLOR','WYMAGANE_KOMPETENCJE'],
    [DB.RULES]:['KLUCZ','WARTOŚĆ','OPIS','EDYTOWALNE'],
    [DB.SCENARIOS]:['SCENARIUSZ_ID*','NAZWA*','MNOŻNIK_ZAPOTRZEBOWANIA*','MNOŻNIK_BUDŻETU*','DOMYŚLNY_POZIOM_OBSADY','NADGODZINY','MAX_NADGODZIN_H','REDUKCJA_DOSTĘPNOŚCI_PROC','ZATRUDNIENIE_CZASOWE','AKTYWNY','OPIS'],
    [DB.MODES]:['TRYB_ID*','NAZWA*','WAGA_KOSZT_PROC*','WAGA_PREFERENCJE_PROC*','WAGA_SPRAWIEDLIWOŚĆ_PROC*','WAGA_POKRYCIE_PROC*','WAGA_CIĄGŁOŚĆ_PROC*','AKTYWNY','OPIS'],
    [DB.LEVELS]:['POZIOM_ID*','NAZWA*','ŹRÓDŁO_CELU*','MNOŻNIK*','LIMIT_BUDŻETU_PROC','AKTYWNY','OPIS'],
    [DB.SKILLS]:['KOD','NAZWA','OPIS','AKTYWNA']
  };
  Object.keys(defs).forEach(name=>{
    let sh=ss.getSheetByName(name);if(!sh)sh=ss.insertSheet(name);
    const h=defs[name];if(sh.getMaxColumns()<h.length)sh.insertColumnsAfter(sh.getMaxColumns(),h.length-sh.getMaxColumns());
    sh.getRange(1,1,1,h.length).setValues([h]).setBackground('#172554').setFontColor('#fff').setFontWeight('bold').setWrap(true);
    sh.setFrozenRows(1);sh.setFrozenColumns(name===DB.EMP?3:1);sh.setHiddenGridlines(true);
    sh.getRange(1,1,Math.max(2,sh.getMaxRows()),h.length).setVerticalAlignment('middle');
    sh.autoResizeColumns(1,h.length);
  });
  dbStart_();dbSeedReferences_();dbApplyValidations_();
  PropertiesService.getDocumentProperties().setProperties({DB_VERSION:DB.VERSION,DB_UPDATED:new Date().toISOString()});
  SpreadsheetApp.getActive().toast('Baza GRAFIK PRO jest gotowa.','GRAFIK PRO',5);
}

function dbStart_(){
  const ss=SpreadsheetApp.getActive();let sh=ss.getSheetByName(DB.GUIDE);if(!sh)sh=ss.insertSheet(DB.GUIDE,0);sh.clear();
  sh.getRange('A1:H2').merge().setValue('GRAFIK PRO — USTAWIENIA I BAZA PRACOWNIKÓW').setBackground('#0f172a').setFontColor('#fff').setFontSize(20).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange('A4:H4').merge().setValue('JEDYNE MIEJSCE DO KONFIGURACJI ZESPOŁU I REGUŁ').setBackground('#dbeafe').setFontColor('#1e3a8a').setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange('A6:H12').merge().setValue(
    '1. Otwórz arkusz BAZA_PRACOWNIKÓW.\n2. Jeden pracownik = jeden wiersz.\n3. Kolumny z * są obowiązkowe.\n4. Wartości wielokrotne rozdzielaj przecinkiem.\n5. Nie zmieniaj nazw kolumn ani identyfikatorów.\n6. Uruchom „Sprawdź bazę”.\n7. W Plannerze wybierz „Synchronizuj centrale”.'
  ).setWrap(true).setVerticalAlignment('middle').setFontSize(13).setBackground('#f8fafc');
  sh.getRange('A14:D18').setValues([
    ['STATUS','ZNACZENIE','DZIAŁANIE',''],
    ['GOTOWY','Rekord kompletny','Może być synchronizowany',''],
    ['NIEKOMPLETNY','Brak pola obowiązkowego','Uzupełnij żółte pola',''],
    ['BŁĄD','Nieprawidłowa zależność','Popraw wskazane wartości',''],
    ['NIEAKTYWNY','Pozostaje w historii','Nie będzie planowany','']
  ]);
  sh.setColumnWidths(1,8,145);sh.setHiddenGridlines(true);
}

function dbSeedReferences_(){
  const write=(name,rows)=>{
    const sh=SpreadsheetApp.getActive().getSheetByName(name);if(sh.getLastRow()>1)return;
    const width=sh.getLastColumn();
    const normalized=rows.map((row,index)=>{
      if(row.length>width)throw new Error(`${name}, wiersz DEMO ${index+2}: ${row.length} wartości dla ${width} kolumn.`);
      return row.concat(Array(Math.max(0,width-row.length)).fill(''));
    });
    if(normalized.length)sh.getRange(2,1,normalized.length,width).setValues(normalized);
  };
  write(DB.LOC,[
    ['LOC-CENTRUM','Centrum','Warszawa — Centrum','TAK','kierownik.centrum@demo.pl','#2563eb','A'],
    ['LOC-OGRODY','Ogrody','Warszawa — Ogrody','TAK','kierownik.ogrody@demo.pl','#7c3aed','B']
  ]);
  write(DB.SHIFT,[
    ['RANO','Zmiana poranna','06:00','14:00',8,'PODSTAWOWA','#fde68a','OBSŁUGA'],
    ['POPOŁUDNIE','Zmiana popołudniowa','14:00','22:00',8,'PODSTAWOWA','#c4b5fd','OBSŁUGA'],
    ['EXTRA','Zmiana dodatkowa','10:00','18:00',8,'DODATKOWA','#86efac','OBSŁUGA'],
    ['STANDBY','Dyżur stand-by','06:00','22:00',2,'STANDBY','#fca5a5','']
  ]);
  write(DB.RULES,[
    ['MIN_ODPOCZYNEK_H',11,'Minimalny odpoczynek dobowy','TAK'],
    ['MAX_DNI_Z_RZĘDU',6,'Domyślny limit dni z rzędu','TAK'],
    ['MAX_H_TYDZIEŃ',48,'Domyślny limit godzin tygodniowo','TAK'],
    ['STANDBY_NA_LOKALIZACJĘ',1,'Liczba rezerw na lokal','TAK'],
    ['TRYB_DOMYŚLNY','ZRÓWNOWAŻONY','Domyślny tryb optymalizacji','TAK'],
    ['BLOKUJ_BŁĘDY_TWARDE','TAK','Zakaz publikacji planu z błędami','NIE']
  ]);
  write(DB.SCENARIOS,[
    ['BAZOWY','Bazowy',1,1,'OPTIMAL','OGRANICZONE',0,0,'NIE','TAK','Standardowe warunki operacyjne'],
    ['WYSOKI_RUCH','Wysoki ruch',1.25,1.15,'OPTIMAL','DOZWOLONE',8,0,'OPCJONALNIE','TAK','Zwiększone zapotrzebowanie i budżet'],
    ['REDUKCJA_KOSZTÓW','Redukcja kosztów',1,0.85,'MINIMUM','BLOKOWANE',0,0,'NIE','TAK','Minimalna bezpieczna obsada i ścisły budżet'],
    ['BRAKI_KADROWE','Braki kadrowe',1,1,'MINIMUM','DOZWOLONE',8,15,'SUGEROWANE','TAK','Symulacja ograniczonej dostępności zespołu'],
    ['SEZONOWY','Sezonowy',1.4,1.25,'OPTIMAL','DOZWOLONE',12,0,'TAK','TAK','Sezonowy wzrost ruchu']
  ]);
  write(DB.MODES,[
    ['ZRÓWNOWAŻONY','Zrównoważony',20,20,25,30,5,'TAK','Równowaga kosztu, preferencji i sprawiedliwości'],
    ['MINIMALNY_KOSZT','Minimalny koszt',50,5,10,30,5,'TAK','Najwyższy priorytet kosztowy'],
    ['PREFERENCJE','Preferencje pracowników',10,45,15,25,5,'TAK','Maksymalizacja zgodności z preferencjami'],
    ['RÓWNY_PODZIAŁ','Równy podział',10,10,45,30,5,'TAK','Wyrównanie wykorzystania nominałów'],
    ['MAKSYMALNE_POKRYCIE','Maksymalne pokrycie',5,5,10,75,5,'TAK','Najwyższy priorytet pełnej obsady']
  ]);
  write(DB.LEVELS,[
    ['MINIMUM','Minimalny','MIN',1,100,'TAK','Wymagane minimum bezpieczeństwa'],
    ['OPTIMAL','Optymalny','OPT',1,100,'TAK','Docelowa obsada operacyjna'],
    ['MAXIMUM','Maksymalny','MAX',1,120,'TAK','Górny dozwolony poziom obsady'],
    ['PLUS_10','110% optymalnego','OPT',1.1,115,'TAK','Obsada optymalna zwiększona o 10%'],
    ['BUDGET','Budżetowy','OPT',1,100,'TAK','Najlepsza obsada mieszcząca się w budżecie'],
    ['DYNAMIC','Dynamiczny','OPT',1,110,'TAK','Obsada zależna od wydarzeń i scenariusza']
  ]);
  write(DB.SKILLS,[
    ['OBSŁUGA','Obsługa podstawowa','Podstawowa praca operacyjna','TAK'],
    ['LIDER','Lider zmiany','Może odpowiadać za zmianę','TAK'],
    ['OTWARCIE','Otwarcie lokalu','Posiada uprawnienia do otwarcia','TAK'],
    ['ZAMKNIĘCIE','Zamknięcie lokalu','Posiada uprawnienia do zamknięcia','TAK']
  ]);
}

function dbLoadDemo(){
  dbInstall();
  const first=['Anna','Marta','Julia','Zofia','Aleksandra','Natalia','Katarzyna','Iga','Lena','Maja','Monika','Karolina','Ewa','Alicja','Weronika'];
  const last=['Nowak','Kowalska','Wiśniewska','Wójcik','Kamińska','Lewandowska','Zielińska','Szymańska','Woźniak','Dąbrowska','Kozłowska','Jankowska','Mazur','Krawczyk','Piotrowska'];
  const rows=[];
  for(let i=0;i<60;i++){
    const full=i<40,oneLoc=i%5<2,home=i%2?'LOC-OGRODY':'LOC-CENTRUM';
    const allowed=oneLoc?home:'LOC-CENTRUM,LOC-OGRODY';
    const leader=i%10===0,onlyMorning=i%11===0,noWeekends=i%17===0;
    rows.push([
      `P${String(i+1).padStart(3,'0')}`,'TAK',`${first[i%15]} ${last[(i*7)%15]}`,`pracownik${i+1}@demo.pl`,`500${String(100000+i).slice(-6)}`,
      home,allowed,home==='LOC-CENTRUM'?'kierownik.centrum@demo.pl':'kierownik.ogrody@demo.pl',
      full?'UMOWA O PRACĘ':'CZĘŚĆ ETATU',full?1:0.5,full?168:84,'2025-01-01','',
      onlyMorning?'TAK':'NIE','NIE',noWeekends?'TAK':'NIE',i%7===0?'NIE':'TAK',
      i%13===0?4:6,full?48:32,11,
      i%4===0?'PONIEDZIAŁEK,ŚRODA':'',i%9===0?'NIEDZIELA':'',onlyMorning?'RANO':i%3===0?'POPOŁUDNIE':'',home,
      leader?'OBSŁUGA,LIDER,OTWARCIE,ZAMKNIĘCIE':'OBSŁUGA',leader?'TAK':'NIE',leader?'TAK':'NIE',leader?'TAK':'NIE',
      leader?2:1,i===0||i===1?'KIEROWNIK':'PRACOWNIK',`KAD-${String(i+1).padStart(4,'0')}`,i%8===0?'Elastyczny grafik':'','GOTOWY'
    ]);
  }
  const sh=SpreadsheetApp.getActive().getSheetByName(DB.EMP);
  if(sh.getLastRow()>1)sh.getRange(2,1,sh.getLastRow()-1,EMP_HEADERS.length).clearContent();
  sh.getRange(2,1,rows.length,EMP_HEADERS.length).setValues(rows);
  dbFormatEmployeeSheet_();dbValidate();dbTouchVersion();
  return {ok:true,employees:60};
}

function dbApplyValidations_(){
  const sh=SpreadsheetApp.getActive().getSheetByName(DB.EMP),n=sh.getMaxRows()-1;
  const yn=SpreadsheetApp.newDataValidation().requireValueInList(['TAK','NIE'],true).setAllowInvalid(false).build();
  [2,14,15,16,17,26,27,28].forEach(c=>sh.getRange(2,c,n,1).setDataValidation(yn));
  sh.getRange(2,9,n,1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['UMOWA O PRACĘ','CZĘŚĆ ETATU','ZLECENIE','B2B','TYMCZASOWA'],true).build());
  sh.getRange(2,30,n,1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['PRACOWNIK','KIEROWNIK','KSIĘGOWOŚĆ','ADMIN'],true).build());
}

function dbFormatEmployeeSheet_(){
  const sh=SpreadsheetApp.getActive().getSheetByName(DB.EMP),last=Math.max(2,sh.getLastRow());
  sh.getRange(2,1,last-1,EMP_HEADERS.length).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
  sh.getRange(2,1,last-1,EMP_HEADERS.length).setWrap(true);
  sh.setColumnWidth(3,180);sh.setColumnWidth(7,190);sh.setColumnWidths(21,4,150);sh.setColumnWidth(25,210);sh.setColumnWidth(31,220);
  const req=[1,2,3,4,6,7,9,10,11,12,25];req.forEach(c=>sh.getRange(1,c).setBackground('#1d4ed8'));
}

function dbValidate(){
  const sh=SpreadsheetApp.getActive().getSheetByName(DB.EMP),data=sh.getDataRange().getValues(),headers=data.shift();
  const required=headers.map((h,i)=>String(h).endsWith('*')?i:-1).filter(i=>i>=0);
  const ids=new Set(),locs=new Set(SpreadsheetApp.getActive().getSheetByName(DB.LOC).getDataRange().getValues().slice(1).map(r=>r[0]));
  const statuses=[],errors=[];
  data.forEach((r,i)=>{
    if(!r.some(Boolean)){statuses.push(['']);return}
    let status=String(r[1]).toUpperCase()==='NIE'?'NIEAKTYWNY':'GOTOWY';
    if(required.some(c=>r[c]==='')){status='NIEKOMPLETNY';errors.push(`Wiersz ${i+2}: brak pola obowiązkowego`);}
    if(ids.has(r[0])){status='BŁĄD';errors.push(`Wiersz ${i+2}: zduplikowane ID ${r[0]}`);}ids.add(r[0]);
    String(r[6]||'').split(',').filter(Boolean).forEach(x=>{if(!locs.has(x.trim())){status='BŁĄD';errors.push(`Wiersz ${i+2}: nieznana lokalizacja ${x}`);}});
    statuses.push([status]);
  });
  dbValidateProfiles_(errors);
  if(statuses.length)sh.getRange(2,33,statuses.length,1).setValues(statuses);
  const rules=[SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('GOTOWY').setBackground('#dcfce7').setRanges([sh.getRange('AG2:AG')]).build(),SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('BŁĄD').setBackground('#fee2e2').setRanges([sh.getRange('AG2:AG')]).build(),SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('NIEKOMPLETNY').setBackground('#fef3c7').setRanges([sh.getRange('AG2:AG')]).build()];
  sh.setConditionalFormatRules(rules);dbTouchVersion();
  SpreadsheetApp.getActive().toast(errors.length?`Błędy: ${errors.length}`:'Baza jest poprawna.','Walidacja',6);
  return {ok:errors.length===0,employees:data.filter(r=>r[0]).length,errors};
}

function dbValidateProfiles_(errors){
  const rows=name=>{const v=SpreadsheetApp.getActive().getSheetByName(name).getDataRange().getValues(),h=v.shift();return v.filter(r=>r.some(Boolean)).map(r=>Object.fromEntries(h.map((x,i)=>[x,r[i]])));};
  rows(DB.MODES).forEach((r,i)=>{
    const total=Number(r['WAGA_KOSZT_PROC*']||0)+Number(r['WAGA_PREFERENCJE_PROC*']||0)+Number(r['WAGA_SPRAWIEDLIWOŚĆ_PROC*']||0)+Number(r['WAGA_POKRYCIE_PROC*']||0)+Number(r['WAGA_CIĄGŁOŚĆ_PROC*']||0);
    if(total!==100)errors.push(`Tryby optymalizacji wiersz ${i+2}: suma wag wynosi ${total}%, powinna 100%`);
  });
  rows(DB.SCENARIOS).forEach((r,i)=>{
    if(Number(r['MNOŻNIK_ZAPOTRZEBOWANIA*'])<=0)errors.push(`Scenariusze wiersz ${i+2}: nieprawidłowy mnożnik zapotrzebowania`);
    if(Number(r['MNOŻNIK_BUDŻETU*'])<=0)errors.push(`Scenariusze wiersz ${i+2}: nieprawidłowy mnożnik budżetu`);
  });
  rows(DB.LEVELS).forEach((r,i)=>{
    if(!['MIN','OPT','MAX'].includes(String(r['ŹRÓDŁO_CELU*']).toUpperCase()))errors.push(`Poziomy obsady wiersz ${i+2}: źródło musi być MIN, OPT albo MAX`);
  });
}

function dbTouchVersion(){
  const stamp=new Date().toISOString();
  PropertiesService.getDocumentProperties().setProperty('DB_UPDATED',stamp);
  SpreadsheetApp.getActive().getSheetByName(DB.GUIDE).getRange('A20:B21').setValues([['WERSJA BAZY',DB.VERSION],['OSTATNIA ZMIANA',stamp]]);
  return stamp;
}
