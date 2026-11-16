/* Formula-preserving report. ExcelJS is bundled locally, no external upload. */
function buildGnbWorkbook(ExcelJS, d, matrix) {
 const wb=new ExcelJS.Workbook();wb.creator='INSAFIN';wb.created=new Date();wb.calcProperties.fullCalcOnLoad=true;
 const ws=wb.addWorksheet('Расчёт ГНБ'), ref=wb.addWorksheet('Справочник');
 const url='https://normativ.kontur.ru/document?documentId=500203&moduleId=9';
 ws.columns=[{width:48},{width:30},{width:12},{width:40}];
 const row=(n,a,b,c,e)=>ws.getRow(n).values=[a,b,c,e];
 const formula=(n,f,result)=>ws.getCell('B'+n).value={formula:f,result};
 row(1,'РАСЧЁТ ГНБ');ws.mergeCells('A1:D1');
 row(2,'Подбор установки и потребность в буровом растворе');ws.mergeCells('A2:D2');
 row(3,'Дата формирования',new Date().toLocaleDateString('ru-RU'));
 row(4,'01 / ИСХОДНЫЕ ДАННЫЕ');
 row(5,'Показатель','Значение','Ед. изм.','Основание / пояснение');
 row(6,'Грунт',d.soilText,'','Выбран в калькуляторе');
 row(7,'Группа по буримости',d.group,'','СП 341, приложение И');
 row(8,'Грунтовый коэффициент F',d.F,'','СП 341, таблица Л.1; 0 — не установлен');
 row(9,'Длина перехода L',d.L,'м','Исходные данные проекта');
 row(10,'Наружный диаметр трубы',d.dn,'мм','Исходные данные проекта');
 row(11,'Суммарный крутящий момент ΣM',d.sigmaM,'кН·м','0 — проверка по моменту не выполнена');
 row(12,'Перебур δ',d.deltaPct,'%','СП 341, Л.2; исходный параметр');
 row(13,'Коэффициент потерь Кр',d.Kr,'','СП 341, Л.1; исходный параметр');
 row(14,'Концентрация бентонита',d.cb,'кг/м³','Рецептура раствора; СП 341, Л.4');
 row(15,'Концентрация полимера',d.cp,'кг/м³','Рецептура раствора; СП 341, Л.4');
 row(16,'Голубые ячейки — исходные данные. Расчётные значения связаны формулами.');ws.mergeCells('A16:D16');
 row(17,'02 / ПОДБОР БУРОВОЙ УСТАНОВКИ');row(18,'Показатель','Результат','Ед. изм.','Нормативное основание');
 const descriptions={19:['Диаметр бурового канала Dбур','мм','Таблица 8.3'],20:['Минимальное усилие Pmin','кН','Таблица А.3'],21:['Коэффициент запаса k1','','Таблица А.2'],22:['Требуемая тяга Pт, не менее','кН','А.2.5, формула (А.1)'],23:['Коэффициент запаса k2','','Таблица А.2'],24:['Требуемый момент Mб, не менее','кН·м','А.2.6, формула (А.2)'],25:['Класс по тяге','','Таблица А.1'],26:['Класс по моменту','','Таблица А.1'],27:['Класс по длине','','Таблица А.1'],28:['Класс по диаметру','','Таблица А.1'],29:['РЕКОМЕНДУЕМЫЙ КЛАСС','','Максимум требований; таблица А.1'],33:['Расчётный объём Vбр','м³','Л.2, формула (Л.2)'],34:['Приготовляемый объём Vприг','м³','Л.1, формула (Л.1)'],35:['Потребность в бентоните','кг','Л.4, формула (Л.8)'],36:['Потребность в полимере','кг','Л.4, формула (Л.8)']};
 for(const [n,[label,unit,source]]of Object.entries(descriptions)){row(+n,label,null,unit,'СП 341.1325800.2017, '+source);ws.getCell('D'+n).value={text:'СП 341: '+source,hyperlink:url};}
 formula(19,'ROUND(IF(B10<=200,IF(B9<50,B10*1.2,IF(B9<100,B10*1.3,IF(B9<300,B10*1.4,B10+100))),IF(B10<=599,IF(B9<100,B10*1.3,IF(B9<300,B10*1.4,B10*1.5)),B10+300)),0)',d.Dbur);
 formula(20,"INDEX('Справочник'!B4:H14,MATCH(B9,'Справочник'!A4:A14,1),MATCH(B19,'Справочник'!B3:H3,1))",d.Pmin);
 formula(21,"IF(OR(B7=\"I\",B7=\"II\",B7=\"III\"),1.5,IF(OR(B7=\"IV\",B7=\"V\",B7=\"VI\"),2,2.5))",d.k1);
 formula(22,'B20*B21',d.Pt);formula(23,'IF(B21=1.5,1.2,IF(B21=2,1.35,1.5))',d.k2);
 formula(24,'IF(B11>0,B11*B23,"НЕ ПРОВЕРЕН")',d.sigmaM>0?d.Mb:'НЕ ПРОВЕРЕН');
 const rank=(cell,a,b,c)=>`IF(${cell}<=${a},1,IF(${cell}<=${b},2,IF(${cell}<=${c},3,4)))`;
 const ranks=[rank('B22',100,400,2500),`IF(B11>0,${rank('B24',10,30,100)},0)`,rank('B9',250,750,1000),rank('B19',300,1000,1800)];
 [d.cPull,d.cMom,d.cLen,d.cDiam].forEach((v,i)=>formula(25+i,i===1?`IF(B11>0,CHOOSE(${ranks[i]},"МИНИ","МИДИ","МАКСИ","МЕГА"),"НЕ ПРОВЕРЕН")`:`CHOOSE(${ranks[i]},"МИНИ","МИДИ","МАКСИ","МЕГА")`,v.name));
 const max=`MAX(${ranks.join(',')})`;
 formula(29,`CHOOSE(${max},"МИНИ","МИДИ","МАКСИ","МЕГА")&IF(B11>0,""," — ПРЕДВАРИТЕЛЬНО")`,d.totalClass+(d.sigmaM>0?'':' — ПРЕДВАРИТЕЛЬНО'));
 row(30,'Паспорт выбранной установки должен соответствовать всем показателям отдельно.');ws.mergeCells('A30:D30');
 row(31,'03 / БУРОВОЙ РАСТВОР');row(32,'Показатель','Результат','Ед. изм.','Нормативное основание');
 const valid=d.F>0&&d.maxRank<3;
 formula(33,`IF(OR(B8<=0,${max}>=3),"НЕ ПРИМЕНИМО",0.785*(B19/1000)^2*(B9+B9*B12/100)*B8)`,valid?d.Vbr:'НЕ ПРИМЕНИМО');
 for(const [n,f,v]of [[34,'B33*B13',d.Vprig],[35,'B34*B14',d.mBent],[36,'B34*B15',d.mPoly]])formula(n,`IF(ISNUMBER(B33),${f},"НЕ ПРИМЕНИМО")`,valid?v:'НЕ ПРИМЕНИМО');
 row(38,'04 / ОБЛАСТЬ ПРИМЕНЕНИЯ');
 row(39,'Для МАКСИ/МЕГА нужен поэтапный расчёт по Л.3.1–Л.3.5. При F = 0 объём раствора не определяется. При ΣM = 0 класс предварительный.');ws.mergeCells('A39:D39');
 row(40,'Основание: СП 341.1325800.2017, приложения А, И, Л и таблица 8.3. Коэффициенты, границы диапазонов и алгоритм перенесены из калькулятора.');ws.mergeCells('A40:D40');
 row(41,'Составил: ____________________','Проверил: ____________________');ws.mergeCells('B41:D41');
 ref.addRow(['СП 341.1325800.2017 / СПРАВОЧНЫЕ ДАННЫЕ']);ref.mergeCells('A1:H1');
 ref.addRow(['Таблица А.3: минимальное усилие, кН. Границы диапазонов — как в калькуляторе.']);ref.mergeCells('A2:H2');
 ref.addRow(['L от, м / D от, мм',...matrix.diamBounds]);matrix.values.forEach((values,i)=>ref.addRow([matrix.lengthBounds[i],...values]));
 ref.getCell('A16').value={text:'Открыть текст СП 341.1325800.2017',hyperlink:url};ref.mergeCells('A16:H16');
 ref.getCell('A18').value='Таблица 8.3 — диаметр канала; А.1 — классы; А.2 — запасы; А.3 — усилие; приложение И — группы грунтов; Л.1 — F; Л.1–Л.4 — раствор и компоненты.';ref.mergeCells('A18:H19');
 ref.getCell('A21').value='F и концентрации остаются исходными данными. При смене грунта проверьте группу и F. Классификация за пределами таблиц требует отдельного обоснования.';ref.mergeCells('A21:H22');
 ref.columns.forEach((c,i)=>c.width=i===0?25:13);
 for(const sh of [ws,ref]){
  sh.eachRow((r,n)=>{r.height=25;r.eachCell({includeEmpty:true},c=>{c.font={name:'Calibri',size:11,color:{argb:'FF203348'}};c.alignment={vertical:'middle',wrapText:true};c.border={bottom:{style:'hair',color:{argb:'FFD7E0E7'}}};});});
  sh.pageSetup={paperSize:9,orientation:sh===ws?'portrait':'landscape',fitToPage:true,fitToWidth:1,fitToHeight:sh===ws?0:1,margins:{left:.3,right:.3,top:.4,bottom:.4,header:.15,footer:.15},printArea:sh===ws?'A1:D41':'A1:H22'};
  sh.headerFooter.oddFooter='&C&P / &N&RРасчёт ГНБ';sh.views=[{state:'frozen',ySplit:5,showGridLines:false}];
 }
 for(const n of [1,4,17,31,38]){if(n!==1)ws.mergeCells(`A${n}:D${n}`);const c=ws.getCell('A'+n);c.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF18364B'}};c.font={name:'Calibri',size:n===1?20:12,bold:true,color:{argb:'FFFFFFFF'}};ws.getRow(n).height=n===1?36:27;}
 for(const n of [5,18,32])ws.getRow(n).eachCell(c=>{c.font={name:'Calibri',size:10,bold:true};c.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFE2EAF0'}};});
 for(let n=6;n<=15;n++){ws.getCell('B'+n).fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFEAF4FB'}};ws.getCell('B'+n).font={name:'Calibri',size:11,color:{argb:'FF155B92'}};}
 for(const n of [22,29,34,35,36])ws.getRow(n).eachCell(c=>{c.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFFFF0CC'}};c.font={name:'Calibri',size:11,bold:true,color:{argb:'FF18364B'}};});
 for(let n=8;n<=36;n++)ws.getCell('B'+n).numFmt='#,##0.00';
 for(const n of [9,10,19,20])ws.getCell('B'+n).numFmt='#,##0';
 ws.getRow(6).height=64;ws.getRow(29).height=36;ws.getRow(39).height=42;ws.getRow(40).height=38;
 for(const n of [1,3])ref.getRow(n).eachCell(c=>{c.font={name:'Calibri',size:11,bold:true,color:{argb:'FFFFFFFF'}};c.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF18364B'}};});
 ws.getRow(30).addPageBreak();ws.pageSetup.printTitlesRow="1:3";
 return wb;
}
async function exportGnbExcel(d,matrix){
 try{const wb=buildGnbWorkbook(ExcelJS,d,matrix);const bytes=await wb.xlsx.writeBuffer();const url=URL.createObjectURL(new Blob([bytes],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));const a=document.createElement('a');a.href=url;a.download='Расчёт_ГНБ_'+new Date().toISOString().slice(0,10)+'.xlsx';a.click();setTimeout(()=>URL.revokeObjectURL(url),10000);}catch(e){console.error(e);alert('Не удалось сформировать Excel. Повторите попытку.');}
}
if(typeof module!=='undefined')module.exports={buildGnbWorkbook};
