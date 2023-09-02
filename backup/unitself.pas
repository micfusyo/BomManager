unit unitSelf;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls,
  Grids, ComCtrls, ComObj, Variants, IniFiles, Math;

procedure SetHintFontStyle();
procedure IniFileRead(const FOwner: TForm);
procedure IniFileWrite(const FOwner: TForm);
function GetStringFromHyperLink(Link: string): string;
function SetStringToHyperLink(Link: string; Text: string): String;
procedure StringGridReNumber(SGrid: TStringGrid);
function StringGridFindString(const SGrid: TStringGrid; const Str: string): Boolean;
function StringGridSearchString(SGrid: TStringGrid; Str: string): integer;
procedure StringGridLoadFromXLSFile(SGrid: TStringGrid; XlsName: string; PBar: TProgressBar);
procedure StringGridSaveToXLSFileBom(SGrid: TStringGrid; XlsName: string; XltName: string; PBar: TProgressBar);
procedure StringGridSaveToXLSFileStk(SGrid: TStringGrid; XlsName: string; XltName: string; PBar: TProgressBar);

var
  DefaultBomXltx: string;
  DefaultOrderXltx: string;
  DefaultStockXltx: string;

const
  constTab = #9;
  constCrLf = #13#10;
  BomColID = 1;
  BomColSyoPart = 2;
  BomColName = 3;
  BomColDesignator = 4;
  BomColFootPrint = 5;
  BomColQuantity = 6;
  BomColManufacturerPart = 7;
  BomColManufacturer = 8;
  BomColSupplier = 9;
  BomColSupplierPart = 10;
  BomColSupplierLink = 11;
  BomColPrice = 12;
  BomColSubTotal = 13;
  BomColPins = 14;
  BomColPinTotal = 15;
  BomColRemark = 16;
  BomColReplacement = 17;
  StkColSyoPart = 1;
  StkColName = 2;
  StkColManufacturerPart = 3;
  StkColManufacturer = 4;
  StkColSupplier = 5;
  StkColSupplierPart = 6;
  StkColSupplierLink = 7;
  StkColPrice = 8;
  StkColPins = 9;
  StkColRemark = 10;
  StkColReplacement = 11;
  StkColDependances = 12;
  StrArray: array[0..2] of String = ('配電線材','配電電料','配電電子');
  StrArray111: array[0..80] of string = ('AWG0','AWG1','AWG2','AWG3','AWG4','AWG5','AWG6','AWG7','AWG8','AWG9','AWG10','AWG11','AWG12','AWG13','AWG14',
    'AWG15','AWG16','AWG17','AWG18','AWG19','AWG20','AWG21','AWG22','AWG23','AWG24','AWG25','AWG26','AWG27','AWG28','AWG29','AWG30','AWG31','AWG32',
    'AWG33','AWG34','AWG35','AWG36','AWG37','AWG38','AWG39','AWG40','AWG41','AWG42','AWG43','AWG44','AWG45','AWG46','AWG47','AWG00','AWG000','AWG0000',
    '0.18mm²','0.2mm²','0.25mm²','0.3mm²','0.5mm²','0.75mm²','1.0mm²','1.25mm²','1.5mm²','1.75mm²','2.0mm²','3.5mm²','5.5mm²','8.0mm²','14mm²','22mm²',
    '30mm²','38mm²','50mm²','60mm²','80mm²','100mm²','125mm²','150mm²','200mm²','250mm²','325mm²','400mm²','500mm²','其他線材');
  StrArray112: array[0..56] of string = ('不斷電糸統','電源穩壓器','電源供應器','電源濾波器','電磁接觸器','電磁接觸器保護蓋','無熔絲開關','無熔絲開關保護蓋',
    '微型斷路器','微型斷路器保護蓋','漏電斷路器','漏電斷路器保護蓋','迴路保護器','迴路保護器保護蓋','過電流保護器','過電流保護器保護蓋','繼電器','繼電器座或配件',
    '通用保險絲','保險絲座或配件','接地銅排','電源插座','電源插頭','散熱風扇','風扇鐵網或配件','端子台或端子盤','端子台護蓋或配件','端子短路片','壓接端子',
    'R型端子','Y型端子','線槽','鋁軌','PLC','人機介面','安全元件','伺服馬達','伺服驅動','計時器','計數器','溫度控制器','溫溼度控制器','各式光源','警示燈或指示燈',
    '急停開關','急停開關保護罩或配件','通用開關','通用開關保護罩或配件','特殊開關','連接頭與配件','多功能儀表','通信轉換器','手工具','其他五金配件','其他光學配件',
    '電腦主件','電腦週邊');
  StrArray113: array[0..33] of string = ('開發軟體','程式驅動','人工智能','特規晶片','電源晶片','微處理器','邏輯晶片','外圍晶片','存儲器','模擬晶片','射頻無線電',
    '二極管','傳感器','線性放大器','接口晶片','驅動器','電阻','電容','晶振','光耦發光管','電路板','晶體管','電感磁珠變壓器','蜂鳴器揚聲器咪頭','保險絲',
    '按鍵開關繼電器','五金類與其他','電源電池','線材與焊料','連接器','萬能板','工具與輔材','微電機馬達','濾波器');
  StrArray300: array[0..33] of string = (
    '通用規格','通用規格','通用規格','通用規格',
    '線性穩壓晶片,低壓差點線性穩壓晶片,開關電源晶片,DC-DC晶片,通用電源晶片,電壓基準晶片,電池電壓管理晶片,電池保護晶片,電源監控晶片,功率開關晶片,電源模塊AC-DC,電源模塊DC-DC,其他',
    'TECHNOL,FREESCALE,GIGADEVICE,HISILTC,HITRENDTECH,HOLTEK,INFINEON,MDT,MEGAWIN,MICROCHIP,NUVOTON,NXP,PADAUK,RENESAS,SAMSUNG,SILICON_LABS,SONIX,SST,ST,STC,SYNCMOS,SYNWIT,TI,WCH',
    '74系列邏輯晶片,4000系列邏輯晶片,時基集成晶片,通用邏輯晶片,CPLD/FPGA晶片,編解碼晶片',
    '單晶片監控晶片,端口擴展晶片,驅動晶片,實時時鐘晶片,字庫晶片,通用嵌入式外圍晶片',
    'DDR存儲器,EEPROM存儲器,EPROM存儲器,FIFO存儲器,FLASH存儲器,FRAM存儲器,PROM存儲器,RAM存儲器',
    '模數轉換晶片,數模轉換晶片,通用模擬晶片,模擬開關晶片,電量計晶片,數字電位器晶片,電流監控晶片',
    '無線收發晶片,射頻開關,射頻卡晶片,射頻放大器,射頻混頻器,射頻檢測器,射頻天線,射頻配件,射頻衰減器',
    '穩壓二極管,通用二極管,快恢復二極管,超快恢復二極管,蕭特基二極管,高壓二極管,整流橋,開關二極管,觸發二極管,TVS二極管,放電管,ESD二極管,雪崩二極管,高效率二極管,變容二極管,恆流二極管',
    '通用傳感器,溫度傳感器,顏色傳感器,圖像傳感器,加速度傳感器,觸摸晶片,溫濕度傳感器,環境光傳感器,紅外線傳感器',
    '通用運放,電壓比較器,音頻放大器,通用放大器,高速寬帶運放,儀表運放,精密運放,FET輸入運放,低噪音運放,低功耗運放,差分運放',
    '通用接口晶片,USB晶片,RS232晶片,RS485/RS422晶片,以太網路晶片,隔離器晶片,CAN晶片,LVDS晶片,4-20MA晶片,音頻接口晶片,DDS晶片',
    'LED驅動,LCD驅動,MOS驅動,大電流驅動,驅動晶片,電機驅動,可控硅驅動,IGBT驅動,EL驅動,功率電子開關',
    '壓敏電阻,貼片電阻,貼片排阻,金屬膜電阻,NTC熱敏電阻,水泥電阻,氧化膜電阻,碳膜電阻,直插排阻,精密可調電阻,一般可調電阻,PTC熱敏電阻,光敏電阻,其他可調電阻,保險電阻,繞線電阻,金屬氧化膜電阻,貼片超低阻值電阻,貼片合金電阻,貼片高精密低溫漂電阻,LED燈條電阻,金屬玻璃釉電阻,採樣電阻,MELF電阻',
    '貼片電容,鉭電容,貼片電解電容,直插電解電容,直插瓷片電容,直插獨石電容,安規電容,高壓電容,可調電容,滌綸電容,CBB電容,校正電容,固態電解電容,超級電容,CL21電容,聚酯薄膜電容,氧化鈮電容',
    '圓柱體晶振,49SMD晶振,49S晶振,49U晶振,貼片有源晶振,聲表面諧振器,陶瓷諧振器,直插有源晶振,貼片無源晶振',
    '直插光耦,貼片光耦,紅外接收管,矽光電池,光電開關,發光二極管,LED數碼管,紅外發射管,LED點陣,LCD液晶顯示模塊,光電固態繼電器,光電可控矽,OLED顯示模塊,菲涅爾透鏡,IrDA紅外收發器,VFD顯示模塊',
    '通用電路板,專案電路板,公司電路板,市售電路模塊',
    '三極管,MOS場效應管,可控矽,達林頓管,通用晶體管,數字三極管,IGBT管,DIODE管',
    '低頻電感,貼片電感,貼片磁珠,網口變壓器,功率電感,色環電感,工字電感,電源變壓器,共模電感,濾波器,直插磁珠,天線,高頻電感,互感器',
    '蜂鳴器,揚聲器,咪頭/矽麥',
    '貼片保險絲,直插保險絲,保險絲夾/座,玻璃管保險絲,短路器,斷路器,一次性保險絲',
    '船型開關,按鍵開關,行程開關,輕觸開關,撥碼開關,五向開關,多功能開關,繼電器,撥動開關,旋轉編碼開關,帶燈開關,鍋仔片,溫控開關,開關插座',
    '散熱器,螺絲,電子輔料,導熱膠,螺柱,測試針,矽膠',
    '電池,電池盒/電池座,電源適配器',
    '電子線及相關,焊錫絲,焊錫膏,焊錫球,助焊劑,綠油,紅膠,助焊膏,電線,FFC連接線,信號線/數據線,電源線',
    'IC座,通用連接器,USB接口,接線端子,RJ45,FPC接口,DC接口,音頻接口,牛角接口,CH接口,VH接口,PH接口,XH接口,2EDG接口,15EDG接口,1.25T接口,D-SUB,卡座,RJ11,排針,排母,短路帽,BNC,PS2接口,KF2510接口,AV接口,Micro_USB,USB_Type-C,杜邦,EH接口,光纖接口,ZH接口1.5mm,DVI,壓線端子,RJ22,板對板連接器,IDC,香蕉頭接口,夾子,隔燈柱,HDMI,SCSI,冷壓端子,XA接口,PA接口',
    '萬能板,面包板',
    '萬用組合式零件盒,萬用表/儀表,萬用表表筆,鑷子,酒精瓶,螺絲刀/鉗,PVC膜,治具,通用輔材,液態光學膠,焊接工具,測電筆,膠水,導電銀漿,儀器,黃膠,化工輔材,熱縮管,其他,雙面膠,吸錫帶,刀具,UPS電源,扣件',
    '微電機馬達',
    '共模扼流圈,磁珠,陶瓷濾波器,RF濾波器,SAW濾波器,饋通式電容器'
  );
  TilArray: array[0..17] of string = (
    '', '項次', '自編料號', '物料品名', '位號', '封裝', '數量', '規格型號', '廠牌', '供應商', '供應商料號', '連結', '單價', '小計', '腳位', '腳數', '備註', '可替代品'
  );
  PcbArray: array[0..17] of string = (
    '', '1', '113-20-02-000-00000', '公司電路板', 'PCB', '10x10cm', '1', '公司電路板SYO-XXX-001-V1.0', 'SYO', 'JLC', '10x10cm', 'JLC', '10.0', '10.0', '0', '0', '', ''
  );

  {
  constHead = '配電線材,配電電料,配電電子';
  constSpec = '通用規格';
  constVolt = '通用規格,額定電壓1V,額定電壓2V,額定電壓3V,額定電壓4V,額定電壓5V,額定電壓6V,額定電壓7V,額定電壓8V,額定電壓9V,額定電壓10V,額定電壓11V,額定電壓12V,額定電壓13V,額定電壓14V,額定電壓15V,額定電壓16V,額定電壓17V,額定電壓18V,額定電壓19V,額定電壓20V,額定電壓21V,額定電壓22V,額定電壓23V,額定電壓24V,額定電壓25V,額定電壓26V,額定電壓27V,額定電壓28V,額定電壓29V,額定電壓30V,額定電壓31V,額定電壓32V,額定電壓33V,額定電壓34V,額定電壓35V,額定電壓36V,額定電壓48V,額定電壓96V,額定電壓100V,額定電壓110V,額定電壓200V,額定電壓210V,額定電壓220V';

  StrSub: array[0..2] of UnicodeString = (
    'AWG0,AWG1,AWG2,AWG3,AWG4,AWG5,AWG6,AWG7,AWG8,AWG9,AWG10,AWG11,AWG12,AWG13,AWG14,AWG15,AWG16,AWG17,AWG18,AWG19,AWG20,AWG21,AWG22,AWG23,AWG24,AWG25,AWG26,AWG27,AWG28,AWG29,AWG30,AWG31,AWG32,AWG33,AWG34,AWG35,AWG36,AWG37,AWG38,AWG39,AWG40,AWG41,AWG42,AWG43,AWG44,AWG45,AWG46,AWG47,AWG00,AWG000,AWG0000,0.18mm²,0.2mm²,0.25mm²,0.3mm²,0.5mm²,0.75mm²,1.0mm²,1.25mm²,1.5mm²,1.75mm²,2.0mm²,3.5mm²,5.5mm²,8.0mm²,14mm²,22mm²,30mm²,38mm²,50mm²,60mm²,80mm²,100mm²,125mm²,150mm²,200mm²,250mm²,325mm²,400mm²,500mm²,其他線材',
    '不斷電糸統,電源穩壓器,電源供應器,電源濾波器,電磁接觸器,電磁接觸器保護蓋,無熔絲開關,無熔絲開關保護蓋,微型斷路器,微型斷路器保護蓋,漏電斷路器,漏電斷路器保護蓋,迴路保護器,迴路保護器保護蓋,過電流保護器,過電流保護器保護蓋,繼電器,繼電器座或配件,通用保險絲,保險絲座或配件,接地銅排,電源插座,電源插頭,散熱風扇,風扇鐵網或配件,端子台或端子盤,端子台護蓋或配件,端子短路片,壓接端子,R型端子,Y型端子,線槽,鋁軌,PLC,人機介面,安全元件,伺服馬達,伺服驅動,計時器,計數器,溫度控制器,溫溼度控制器,各式光源,警示燈或指示燈,急停開關,急停開關保護罩或配件,通用開關,通用開關保護罩或配件,特殊開關,連接頭與配件,多功能儀表,通信轉換器,手工具,其他五金配件,其他光學配件,電腦主件,電腦週邊',
    '開發軟體,程式驅動,人工智能,特規晶片,電源晶片,微處理器,邏輯晶片,外圍晶片,存儲器,模擬晶片,射頻無線電,二極管,傳感器,線性放大器,接口晶片,驅動器,電阻,電容,晶振,光耦發光管,電路板,晶體管,電感磁珠變壓器,蜂鳴器揚聲器咪頭,保險絲,按鍵開關繼電器,五金類與其他,電源電池,線材與焊料,連接器,萬能板,工具與輔材,微電機馬達,濾波器'
  );
  StrLv1: UnicodeString = '通用規格,單芯1C,單芯2C,單芯3C,單芯4C,單芯5C,單芯6C,單芯7C,單芯8C,單芯9C,單芯12C,單芯15C,單芯18C,單芯21C,單芯25C,多芯1C,多芯2C,多芯3C,多芯4C,多芯5C,多芯6C,多芯7C,多芯8C,多芯9C,多芯12C,多芯15C,多芯18C,多芯21C,多芯25C,對絞非隔離2對,對絞非隔離3對,對絞非隔離4對,對絞非隔離5對,對絞非隔離6對,對絞非隔離8對,對絞單隔離2對,對絞單隔離3對,對絞單隔離4對,對絞單隔離5對,對絞單隔離6對,對絞單隔離8對,對絞雙隔離2對,對絞雙隔離3對,對絞雙隔離4對,對絞雙隔離5對,對絞雙隔離6對,對絞雙隔離8對,電源線,信號線';
  StrLv2: array[0..56] of UnicodeString = (
  );
  StrLv3: array[0..33] of UnicodeString = (
    constSpec,constSpec,constSpec,constSpec,
    '線性穩壓晶片,低壓差點線性穩壓晶片,開關電源晶片,DC-DC晶片,通用電源晶片,電壓基準晶片,電池電壓管理晶片,電池保護晶片,電源監控晶片,功率開關晶片,電源模塊AC-DC,電源模塊DC-DC,其他',
    'TECHNOL,FREESCALE,GIGADEVICE,HISILTC,HITRENDTECH,HOLTEK,INFINEON,MDT,MEGAWIN,MICROCHIP,NUVOTON,NXP,PADAUK,RENESAS,SAMSUNG,SILICON_LABS,SONIX,SST,ST,STC,SYNCMOS,SYNWIT,TI,WCH',
    '74系列邏輯晶片,4000系列邏輯晶片,時基集成晶片,通用邏輯晶片,CPLD/FPGA晶片,編解碼晶片',
    '單晶片監控晶片,端口擴展晶片,驅動晶片,實時時鐘晶片,字庫晶片,通用嵌入式外圍晶片',
    'DDR存儲器,EEPROM存儲器,EPROM存儲器,FIFO存儲器,FLASH存儲器,FRAM存儲器,PROM存儲器,RAM存儲器',
    '模數轉換晶片,數模轉換晶片,通用模擬晶片,模擬開關晶片,電量計晶片,數字電位器晶片,電流監控晶片',
    '無線收發晶片,射頻開關,射頻卡晶片,射頻放大器,射頻混頻器,射頻檢測器,射頻天線,射頻配件,射頻衰減器',
    '穩壓二極管,通用二極管,快恢復二極管,超快恢復二極管,蕭特基二極管,高壓二極管,整流橋,開關二極管,觸發二極管,TVS二極管,放電管,ESD二極管,雪崩二極管,高效率二極管,變容二極管,恆流二極管',
    '通用傳感器,溫度傳感器,顏色傳感器,圖像傳感器,加速度傳感器,觸摸晶片,溫濕度傳感器,環境光傳感器,紅外線傳感器',
    '通用運放,電壓比較器,音頻放大器,通用放大器,高速寬帶運放,儀表運放,精密運放,FET輸入運放,低噪音運放,低功耗運放,差分運放',
    '通用接口晶片,USB晶片,RS232晶片,RS485/RS422晶片,以太網路晶片,隔離器晶片,CAN晶片,LVDS晶片,4-20MA晶片,音頻接口晶片,DDS晶片',
    'LED驅動,LCD驅動,MOS驅動,大電流驅動,驅動晶片,電機驅動,可控硅驅動,IGBT驅動,EL驅動,功率電子開關',
    '壓敏電阻,貼片電阻,貼片排阻,金屬膜電阻,NTC熱敏電阻,水泥電阻,氧化膜電阻,碳膜電阻,直插排阻,精密可調電阻,一般可調電阻,PTC熱敏電阻,光敏電阻,其他可調電阻,保險電阻,繞線電阻,金屬氧化膜電阻,貼片超低阻值電阻,貼片合金電阻,貼片高精密低溫漂電阻,LED燈條電阻,金屬玻璃釉電阻,採樣電阻,MELF電阻',
    '貼片電容,鉭電容,貼片電解電容,直插電解電容,直插瓷片電容,直插獨石電容,安規電容,高壓電容,可調電容,滌綸電容,CBB電容,校正電容,固態電解電容,超級電容,CL21電容,聚酯薄膜電容,氧化鈮電容',
    '圓柱體晶振,49SMD晶振,49S晶振,49U晶振,貼片有源晶振,聲表面諧振器,陶瓷諧振器,直插有源晶振,貼片無源晶振',
    '直插光耦,貼片光耦,紅外接收管,矽光電池,光電開關,發光二極管,LED數碼管,紅外發射管,LED點陣,LCD液晶顯示模塊,光電固態繼電器,光電可控矽,OLED顯示模塊,菲涅爾透鏡,IrDA紅外收發器,VFD顯示模塊',
    '通用電路板,專案電路板,公司電路板,市售電路模塊',
    '三極管,MOS場效應管,可控矽,達林頓管,通用晶體管,數字三極管,IGBT管,DIODE管',
    '低頻電感,貼片電感,貼片磁珠,網口變壓器,功率電感,色環電感,工字電感,電源變壓器,共模電感,濾波器,直插磁珠,天線,高頻電感,互感器',
    '蜂鳴器,揚聲器,咪頭/矽麥',
    '貼片保險絲,直插保險絲,保險絲夾/座,玻璃管保險絲,短路器,斷路器,一次性保險絲',
    '船型開關,按鍵開關,行程開關,輕觸開關,撥碼開關,五向開關,多功能開關,繼電器,撥動開關,旋轉編碼開關,帶燈開關,鍋仔片,溫控開關,開關插座',
    '散熱器,螺絲,電子輔料,導熱膠,螺柱,測試針,矽膠',
    '電池,電池盒/電池座,電源適配器',
    '電子線及相關,焊錫絲,焊錫膏,焊錫球,助焊劑,綠油,紅膠,助焊膏,電線,FFC連接線,信號線/數據線,電源線',
    'IC座,通用連接器,USB接口,接線端子,RJ45,FPC接口,DC接口,音頻接口,牛角接口,CH接口,VH接口,PH接口,XH接口,2EDG接口,15EDG接口,1.25T接口,D-SUB,卡座,RJ11,排針,排母,短路帽,BNC,PS2接口,KF2510接口,AV接口,Micro_USB,USB_Type-C,杜邦,EH接口,光纖接口,ZH接口1.5mm,DVI,壓線端子,RJ22,板對板連接器,IDC,香蕉頭接口,夾子,隔燈柱,HDMI,SCSI,冷壓端子,XA接口,PA接口',
    '萬能板,面包板',
    '萬用組合式零件盒,萬用表/儀表,萬用表表筆,鑷子,酒精瓶,螺絲刀/鉗,PVC膜,治具,通用輔材,液態光學膠,焊接工具,測電筆,膠水,導電銀漿,儀器,黃膠,化工輔材,熱縮管,其他,雙面膠,吸錫帶,刀具,UPS電源,扣件',
    '微電機馬達',
    '共模扼流圈,磁珠,陶瓷濾波器,RF濾波器,SAW濾波器,饋通式電容器'
  );
  }

implementation

procedure SetHintFontStyle();
begin
  with Screen do
  begin
    HintFont.Size:= 12;
    HintFont.Color:= clRed;
    HintFont.Style:= [fsBold];
  end;
end;

procedure IniFileRead(const FOwner: TForm);
var
  IniFile: TIniFile;
begin
  IniFile:= TIniFile.Create(ChangeFileExt(Application.ExeName, '.ini'));
  with IniFile do
  begin
    FOwner.Top:= ReadInteger('MainFormPlacement', 'Top', FOwner.Top);
    FOwner.Left:= ReadInteger('MainFormPlacement', 'Left', FOwner.Left);
    FOwner.Width:= ReadInteger('MainFormPlacement', 'Width', FOwner.Width);
    FOwner.Height:= ReadInteger('MainFormPlacement', 'Height', FOwner.Height);
    DefaultBomXltx:= ReadString('DefaultFileName', 'BomXltxFile', 'syo_bom.xltx');
    DefaultOrderXltx:= ReadString('DefaultFileName', 'OrderXltxFile', 'syo_order.xltx');
    DefaultStockXltx:= ReadString('DefaultFileName', 'StockXltxFile', 'syo_stock.xltx');
    Free;
  end;
end;

procedure IniFileWrite(const FOwner: TForm);
var
  IniFile: TIniFile;
begin
  IniFile:= TIniFile.Create(ChangeFileExt(Application.ExeName, '.ini'));
  with IniFile do
  begin
    WriteInteger('MainFormPlacement', 'Top', FOwner.Top);
    WriteInteger('MainFormPlacement', 'Left', FOwner.Left);
    WriteInteger('MainFormPlacement', 'Width', FOwner.Width);
    WriteInteger('MainFormPlacement', 'Height', FOwner.Height);
    WriteString('DefaultFileName', 'BomXltxFile', DefaultBomXltx);
    WriteString('DefaultFileName', 'OrderXltxFile', DefaultOrderXltx);
    WriteString('DefaultFileName', 'StockXltxFile', DefaultStockXltx);
    Free;
  end;
end;

function GetStringFromHyperLink(Link: string): string;
var
  Pos: integer;
begin
  Pos:= AnsiPos(',"', Link) + 2;
  Result:= Copy(Link, Pos, (Length(Link)-Pos)-2);
end;

function SetStringToHyperLink(Link: string; Text: string): String;
const
  LinkStr1: string = '=HYPERLINK("https://so.szlcsc.com/global.html?k=';
  LinkStr2: string = '&hot-key=ADXL355BEZ-RL7","';
  LinkStr3: string = '")';
begin
   Result:= LinkStr1 + Link + LinkStr2 + Text + LinkStr3;
end;

procedure StringGridReNumber(SGrid: TStringGrid);
var
  Row: Integer;
begin
  for Row:= 1 to SGrid.RowCount-1 do
    SGrid.Cells[0, Row]:= IntToStr(Row);
  SGrid.AutoSizeColumns;
end;

function StringGridFindString(const SGrid: TStringGrid; const Str: string): Boolean;
var
  Row, Col: Integer;
begin
  Result:= False;
  for Row:= 0 to SGrid.RowCount - 1 do
  begin
    for Col:= 0 to SGrid.ColCount - 1 do
    begin
      // Check if string is found in cell
      if Pos(Str, SGrid.Cells[Col, Row]) > 0 then
      begin
        // Select cell with match
        SGrid.Row := Row;
        SGrid.Col := Col;
        Result := True;
        SGrid.TopRow := Max(0, Row - (SGrid.VisibleRowCount div 2));
        // Exit function as soon as a match is found
        Exit;
      end;
    end;
  end;
end;

function StringGridSearchString(SGrid: TStringGrid; Str: string): integer;
var
  Idx: integer;
begin
  Result:= 0;
  for Idx:= 2 to SGrid.RowCount-1 do
  begin
    if SGrid.Cells[StkColSupplierPart, Idx] = Str then
    begin
      Result:= Idx;
      break;
    end;
  end;
end;

procedure StringGridLoadFromXLSFile(SGrid: TStringGrid; XlsName: string; PBar: TProgressBar);
var
  ExcelApp: Variant;
  Workbook: Variant;
  Sheet: Variant;
  RowCount, ColCount: Integer;
  Row, Col: Integer;
begin
  if not FileExists(XlsName) then
  begin
    ShowMessage('Excel filename failed!');
    exit;
  end;
  ExcelApp := CreateOleObject('Excel.Application');
  Workbook := ExcelApp.Workbooks.Open(XlsName);
  Sheet := Workbook.Sheets[1];
  RowCount:= Sheet.UsedRange.Rows.Count;
  ColCount:= Sheet.UsedRange.Columns.Count;
  PBar.Position:= 0;
  PBar.Min:= 0;
  PBar.Max:= RowCount-1;
  SGrid.RowCount:= RowCount+1;
  SGrid.ColCount:= ColCount+1;
  // Set Title
  for Col:= 1 to ColCount do
    SGrid.Cells[Col, 0]:= ExcelApp.ActiveSheet.Cells[1, Col].Value;
  // Load excel
  for Row:= 1 to RowCount do
  begin
    PBar.Position:= Row;
    for Col:= 1 to ColCount do
    begin
      SGrid.Cells[Col, Row]:= Sheet.Cells[Row, Col].Value;
    end;
  end;
  Workbook.Close;
  ExcelApp.Quit;
  PBar.Position:= 0;
end;

procedure StringGridSaveToXLSFileBom(SGrid: TStringGrid; XlsName: string; XltName: string; PBar: TProgressBar);
var
  ExcelApp: Variant;
  Workbook: Variant;
  Sheet: Variant;
  Row, Col: Integer;
begin
  if not FileExists(XltName) then
  begin
    ShowMessage('Excel template failed!');
    exit;
  end;
  ExcelApp := CreateOleObject('Excel.Application');
  Workbook := ExcelApp.Workbooks.Open(XltName);
  Sheet := Workbook.Sheets[1];
  PBar.Position:= 0;
  PBar.Min:= 0;
  PBar.Max:= SGrid.RowCount-1;
  for Row:= 2 to SGrid.RowCount-1 do
  begin
    PBar.Position:= Row;
    for Col:= 1 to SGrid.ColCount-1 do
    begin
      case Col of
        BomColQuantity:
        begin
          if Row = 2 then
            Sheet.Cells[Row, Col].Value:= SGrid.Cells[Col, Row]
          else
            Sheet.Cells[Row, Col].Formula:= '=F$2*' + SGrid.Cells[Col, Row];
        end;
        BomColSupplierLink:
        begin
          if Copy(SGrid.Cells[BomColSupplierPart, Row], 1, 1) = 'C' then
            Sheet.Cells[Row, Col].Formula:= SetStringToHyperLink(SGrid.Cells[BomColSupplierPart, Row], SGrid.Cells[BomColSupplier, Row])
          else
            Sheet.Cells[Row, Col].Value:= SGrid.Cells[Col, Row];
        end;
        BomColSubTotal:
          Sheet.Cells[Row, Col].Formula:= '=F' + IntToStr(Row) + '*L' + IntToStr(Row);
        BomColPinTotal:
          Sheet.Cells[Row, Col].Formula:= '=F' + IntToStr(Row) + '*N' + IntToStr(Row);
      else
        Sheet.Cells[Row, Col].Value:= SGrid.Cells[Col, Row];
      end;
    end;
  end;
  WorkBook.SaveAs(XlsName);
  WorkBook.Saved:= True;
  Workbook.Close;
  ExcelApp.Quit;
  PBar.Position:= 0;
end;

procedure StringGridSaveToXLSFileStk(SGrid: TStringGrid; XlsName: string; XltName: string; PBar: TProgressBar);
var
  ExcelApp: Variant;
  Workbook: Variant;
  Sheet: Variant;
  Row, Col: Integer;
begin
  if not FileExists(XltName) then
  begin
    ShowMessage('Excel template failed!');
    exit;
  end;
  ExcelApp := CreateOleObject('Excel.Application');
  Workbook := ExcelApp.Workbooks.Open(XltName);
  Sheet := Workbook.Sheets[1];
  PBar.Position:= 0;
  PBar.Min:= 0;
  PBar.Max:= SGrid.RowCount-1;
  for Row:= 2 to SGrid.RowCount-1 do
  begin
    PBar.Position:= Row;
    for Col:= 1 to SGrid.ColCount-1 do
    begin
      if Col = StkColSupplierLink then
      begin
        if Copy(SGrid.Cells[StkColSupplierPart, Row], 1, 1) = 'C' then
          Sheet.Cells[Row, Col].Formula:= SetStringToHyperLink(SGrid.Cells[StkColSupplierPart, Row], SGrid.Cells[StkColSupplier, Row])
        else
          Sheet.Cells[Row, Col].Value:= SGrid.Cells[Col, Row];
      end else
      begin
        Sheet.Cells[Row, Col].Value:= SGrid.Cells[Col, Row];
      end;
    end;
  end;
  WorkBook.SaveAs(XlsName);
  WorkBook.Saved:= True;
  Workbook.Close;
  ExcelApp.Quit;
  PBar.Position:= 0;
end;

end.

