const ss = SpreadsheetApp.getActiveSpreadsheet()
const formNovaDostavka = ss.getSheetByName("нова доставка")
const sheetDostavki = ss.getSheetByName("доставки")
const formNovoProizvodstvo = ss.getSheetByName("ново производство")
const sheetRecepti = ss.getSheetByName("рецепти")
const sheetProizvodstvo = ss.getSheetByName("производство")

/**
 * функция за запаметяване на новите доставки
 */
function saveNewOrder() {
  //взимаме масив от таблицата на досегашните доставки
  var colArray = sheetDostavki.getRange("A2:A"+sheetDostavki.getLastRow()).getValues();
  //взимаме последното число (партида)
  var maxInColumn = colArray.sort(function(a,b){return b-a})[0][0];
  //увеличаваме с 1 за да създадем нова партида като номер
  maxInColumn++
  //взимаме продукта и количеството в масив заедно с новата партида
  const product = formNovaDostavka.getRange("C3").getValue()
  const qty = formNovaDostavka.getRange("C5").getValue()
  const newOrder = [maxInColumn, product, qty]
  //добавяме масива с доставката като нов ред
  sheetDostavki.appendRow(newOrder)

  //готови сме с доставката и нулираме формата
  maxInColumn++
  formNovaDostavka.getRange("C4").setValue(maxInColumn)
  formNovaDostavka.getRange("C3").setValue("---")
  formNovaDostavka.getRange("C5").setValue("---")

}

/**
 * функция за запаметяване на ново производство
 * взима данни от вече изчислената форма НОВО ПРОИЗВОДСТВО
 */
function saveNovoProizvodstvo(){

  //взима всички данни
  var data = formNovoProizvodstvo.getDataRange().getValues();

  //върти редовете в цикъл
  //без 1 ред - имена на колонки
  count=1;
  for( a in data ){

    if(count>1){ //прескачаме първи ред с имената на колонките

    proizvedenProdukt = formNovoProizvodstvo.getRange("A2").getValue();
    broiProizvedenProdukt = formNovoProizvodstvo.getRange("B2").getValue();
    tekushtaSurovinaIme = data[a][2]; //C колона
    kolichestvoNeobhodimaSurovinaZaProizvodstvoEdinBroi = data[a][3]; //D колона
    kolichestvoNeobhodimaSurovinaZaProizvodstvo = data[a][4]; //E колона
    partida1 = data[a][8]; //I колона
    kolichestvoPartida1 = data[a][9]; //J колона
    partida2 = data[a][10]; //K колона
    kolichestvoPartida2 = data[a][11]; //L колона
    partida3 = data[a][12]; //M колона
    kolichestvoPartida3 = data[a][13]; //N колона
    //Logger.log(kolichestvoNeobhodimaSurovinaZaProizvodstvo);

    //в този масив ще се съхранява информацията за производството 
    //преди да се запамети в таблица ПРОИЗВОДСТВО
    //структура масив:
    //0 - Произведен продукт,
    //1 - Суровина име	
    //2 - Суровина партида	
    //3 - Суровина количество 1 бр	
    //4 - Произведени бройки	 
    //5 - общо кол по бр	
    //6 - Суровина използвано количество от тази партида
    proizvodstvoData = []; 
    //проверява количеството от първата партида дали стига за това производство
    if( kolichestvoPartida1 > kolichestvoNeobhodimaSurovinaZaProizvodstvo ){ 
      //ако първа партида стига
      //запаметява данните в масив производство
      proizvodstvoData.push([proizvedenProdukt,
                             tekushtaSurovinaIme,
                             partida1,
                             kolichestvoNeobhodimaSurovinaZaProizvodstvoEdinBroi,
                             broiProizvedenProdukt,
                             kolichestvoNeobhodimaSurovinaZaProizvodstvo,
                             kolichestvoNeobhodimaSurovinaZaProizvodstvo]);
    }else{ 
        //ако не стига първа партида
        //ако не стига, запаметява количеството от тази партида!!!!! в таблица ПРОИЗВОДСТВО
        //запаметява данните в масив производство от 1 партида
        proizvodstvoData.push([proizvedenProdukt,
                              tekushtaSurovinaIme,
                              partida1,
                              kolichestvoNeobhodimaSurovinaZaProizvodstvoEdinBroi,
                              broiProizvedenProdukt,
                              kolichestvoNeobhodimaSurovinaZaProizvodstvo,
                              kolichestvoPartida1]);

        //проверява дали първа + втора партида стигат
        if( (kolichestvoPartida1+kolichestvoPartida2) > kolichestvoNeobhodimaSurovinaZaProizvodstvo ){
          //ако двете партиди стигат
          //и добавя разликата от втората партида
          //тук запаметява необходимата втора партида ако тя стига
          razlikaNeobhodimoKol = kolichestvoNeobhodimaSurovinaZaProizvodstvo - kolichestvoPartida1;
          proizvodstvoData.push([proizvedenProdukt,
                                tekushtaSurovinaIme,
                                partida2,
                                kolichestvoNeobhodimaSurovinaZaProizvodstvoEdinBroi,
                                broiProizvedenProdukt,
                                kolichestvoNeobhodimaSurovinaZaProizvodstvo,
                                razlikaNeobhodimoKol]);

        }else{
          //ако и двете първи партиди не стигат - 
          //първо запаметяваме цялата втора партида
          proizvodstvoData.push([proizvedenProdukt,
                                tekushtaSurovinaIme,
                                partida2,
                                kolichestvoNeobhodimaSurovinaZaProizvodstvoEdinBroi,
                                broiProizvedenProdukt,
                                kolichestvoNeobhodimaSurovinaZaProizvodstvo,
                                kolichestvoPartida2]);
          //проверяваме и заедно с третата 
          if( (kolichestvoPartida1+kolichestvoPartida2+kolichestvoPartida3) > kolichestvoNeobhodimaSurovinaZaProizvodstvo ){
            //ако третата партида ще стигне
            razlikaNeobhodimoKol2 = kolichestvoNeobhodimaSurovinaZaProizvodstvo - (kolichestvoPartida1+kolichestvoPartida2);
            //запаметяваме колкото е необходимо (разликата) като партида 3
            proizvodstvoData.push([proizvedenProdukt,
                                  tekushtaSurovinaIme,
                                  partida2,
                                  kolichestvoNeobhodimaSurovinaZaProizvodstvoEdinBroi,
                                  broiProizvedenProdukt,
                                  kolichestvoNeobhodimaSurovinaZaProizvodstvo,
                                  razlikaNeobhodimoKol2]);
          } //проверка за трета партида
        } //проверка втора
    } //проверка дали първа партида не стига

    //запаметява всичко в таблица ПРОИЗВОДСТВО
    finalSaveProizvodstvoArray(proizvodstvoData);

    } //ако реда е след първи (от втори нататък)
    count++;
  } //цикъл върти всички редове от рецептата във форма НОВО ПРОИЗВОДСТВО
}

/**
 * запаметява готовия масив за таблица ПРОИЗВОДСТВО
 * масива е всеки елемент съдържа по един ред с масив инфо за реда
 *  //структура масив:
 *  //0 - Произведен продукт,
 *  //1 - Суровина име	
 *  //2 - Суровина партида	
 *  //3 - Суровина количество 1 бр	
 *  //4 - Произведени бройки	 
 *  //5 - общо кол по бр	
 *  //6 - Суровина използвано количество от тази партида
 */
function finalSaveProizvodstvoArray(proizvodstvoArray){

  //трябва ни дата
  currentDate = Utilities.formatDate(new Date(), "GMT+2", "dd.MM.yyyy");
  //Logger.log(proizvodstvoArray);
  //цикъл въртим масива
  for( a in proizvodstvoArray ){

      //взимаме масив от таблицата на досегашните производства - 
      //поредния номер е в колонка А
      var colArray = sheetProizvodstvo.getRange("A2:A"+sheetProizvodstvo.getLastRow()).getValues();
      //взимаме последното число (партида)
      var maxInColumn = colArray.sort(function(a,b){return b-a})[0][0];
      //увеличаваме с 1 за да създадем нов номер на ред за таблица производство
      maxInColumn++

      //подготвяме нов масив, който ще е ред за запис в таблица ПРОИЗВОДСТВО
      //Пореден номер	
      //Дата	
      //Произведен продукт	
      proizvedenProdukt = proizvodstvoArray[a][0];
      //Суровина име	
      surovinaIme = proizvodstvoArray[a][1];
      //Суровина партида	
      surovinaPartida = proizvodstvoArray[a][2];
      //Суровина количество 1 бр	
      surovinaKolichestvo1br = proizvodstvoArray[a][3];
      //Произведени бройки	
      proizvedeniBroiki =  proizvodstvoArray[a][4];
      //общо кол по бр	
      obshtoKolichestvoPoBroi = proizvodstvoArray[a][5];
      //Суровина използвано количество от тази партида
      kolichestvoSurovinaOtTaziPartida = proizvodstvoArray[a][5];

      rowInProizvodstvo = [
        maxInColumn,
        currentDate,
        proizvedenProdukt,
        surovinaIme,
        surovinaPartida,
        surovinaKolichestvo1br,
        proizvedeniBroiki,
        obshtoKolichestvoPoBroi,
        kolichestvoSurovinaOtTaziPartida
      ];

      //Logger.log(rowInProizvodstvo);

      //добавяме масива с доставката като нов ред
      sheetProizvodstvo.appendRow(rowInProizvodstvo);
  } //цикъл нов ред производство

  //изтриваме старите данни от формата за ново производство
  cleanFormNovoProizvodstvo();

  return true;
}

/**
 * функция за взимане на рецептата за производство на един брой продукт
 * приема име на продукт, за който ни трябва рецептата
 * връща масив име,суровина,количество
 */
function getProductRecepta(product){
  //трябва да вземем рецептата за един брой от този продукт
  //от таблица рецепти
  var data = sheetRecepti.getDataRange().getValues();
  var productRecepta = [];
  for( i in data ){
    //data[i][0] MATSAN DOYCH
    //data[i][1] Vitamin C
    //data[i][2] 0.1
    //рецептата я наливаме в нов масив само за избрания продукт
    if( data[i][0] == product ){
      prName = data[i][0];
      prSurovina = data[i][1];
      prKolichestvo = data[i][2];
      productRecepta.push( [prName, prSurovina, prKolichestvo] );
    }
  }
  //Logger.log(productRecepta) 
  //[[MATSAN DOYCH, Vitamin C, 0.1], [MATSAN DOYCH, Kurkuma, 0.2], [MATSAN DOYCH, Lecitin, 0.3]]
  return productRecepta;
}

/**
 * функция за почистване на таблицата преди калкулация и нанасяне на нови стоности
 * почистване на клетки C2-C30 и D2-D30
 * 
 */
function cleanFormNovoProizvodstvo(){
  //почистване на клетки C2-C30 и D2-D30
  for(count=2; count<30; count++){
    formNovoProizvodstvo.getRange("C"+count).setValue('');
    formNovoProizvodstvo.getRange("D"+count).setValue('');
    formNovoProizvodstvo.getRange("E"+count).setValue('');
    formNovoProizvodstvo.getRange("F"+count).setValue('');
    formNovoProizvodstvo.getRange("G"+count).setValue('');
    formNovoProizvodstvo.getRange("H"+count).setValue('');
    formNovoProizvodstvo.getRange("I"+count).setValue('');
    formNovoProizvodstvo.getRange("J"+count).setValue('');
    formNovoProizvodstvo.getRange("K"+count).setValue('');
    formNovoProizvodstvo.getRange("L"+count).setValue('');
    formNovoProizvodstvo.getRange("M"+count).setValue('');
    formNovoProizvodstvo.getRange("N"+count).setValue('');
  }
  return true;
}

/**
 * функция за изчисляване на количества суровини спрямо избран продукт
 * от бутон calculate при НОВО ПРОИЗВОДСТВО
 * нанася резултатите в таблицата във форма "ново производство"
 */
function calculateNovoProizvodstvo() {

  //взимаме избраният продукт от падащото меню
  const product = formNovoProizvodstvo.getRange("A2").getValue()

  //взимаме рецептата за избрания продукт
  productRecepta = getProductRecepta(product);

  //взимаме количеството, което ще се произвежда - от клетка B2
  const qtyToBeMade = formNovoProizvodstvo.getRange("B2").getValue();

  //въртим в цикъл всяка суровина от рецептата
  var count = 2;
  for( a in productRecepta ){

    imeSurovina = productRecepta[a][1];
    kolSurovinaEdinBroi = productRecepta[a][2];

    //добавяне на рецептата във формата
    formNovoProizvodstvo.getRange("C"+count).setValue(imeSurovina)
    //добавяне на бройките за един брой произведен продукт към формата за ново производство
    formNovoProizvodstvo.getRange("D"+count).setValue(kolSurovinaEdinBroi)
    //количество от рецепта по бройки за производство
    qtyReceiptByQtyToBeMade = qtyToBeMade * kolSurovinaEdinBroi;
    //нанасяме умножените бройки по количество в таблицата
    formNovoProizvodstvo.getRange("E"+count).setValue(qtyReceiptByQtyToBeMade);
    //изчисляваме колко количество има доставено от съответната суровина
    kgDostaveniObshto = getSurovinaDostaveniKg(imeSurovina);
    //нанасяме количествата в таблицата
    formNovoProizvodstvo.getRange("F"+count).setValue(kgDostaveniObshto);
    //изчисляваме колко количество досега от суровината е използвано в производство
    allUsedQtyBySurovina = getAllUsedQtyBySurovina(imeSurovina);
    //оставяме го в колонка G
    formNovoProizvodstvo.getRange("G"+count).setValue(allUsedQtyBySurovina);
    //изчисляваме колко има в момента от тази суровина (доставено-използвано)
    obshtoNalichnoKolOtSurovina = kgDostaveniObshto - allUsedQtyBySurovina;
    //поставяме наличното в момента количество в колонка H
    formNovoProizvodstvo.getRange("H"+count).setValue(obshtoNalichnoKolOtSurovina);

    //взимаме масив с трите партиди с неизразходвани количества сортирани по номер партида 
    //най-старата първа подред, намираме ги по име на суровина
    lastPartidesWithQty = getLastThreePartidesWithQtyBySurovina(imeSurovina);

    //подреждаме трите партиди и количества вдясно
    formNovoProizvodstvo.getRange("I"+count).setValue(lastPartidesWithQty[0][0]);
    formNovoProizvodstvo.getRange("J"+count).setValue(lastPartidesWithQty[0][1]);

    formNovoProizvodstvo.getRange("K"+count).setValue(lastPartidesWithQty[1][0]);
    formNovoProizvodstvo.getRange("L"+count).setValue(lastPartidesWithQty[1][1]);

    formNovoProizvodstvo.getRange("M"+count).setValue(lastPartidesWithQty[2][0]);
    formNovoProizvodstvo.getRange("N"+count).setValue(lastPartidesWithQty[2][1]);

    count++
  }
}

/**
 * функция която връща трите последни партиди с количества, сортирани по  номер партида
 */
function getLastThreePartidesWithQtyBySurovina(imeSurovina){

  //взимаме всички партиди от текущата суровина
  surovinaDostaveniPartidi = getSurovina(imeSurovina);
  //array [dostavkaLot, dostavkaProdukt, dostavkaQty]

  lastThreePartidesWithQty = [];
  //въртим ги всички в цикъл
  for (c in surovinaDostaveniPartidi){

    dostavenaPartida = surovinaDostaveniPartidi[c][0];
    dostavenaSurovinaIme = surovinaDostaveniPartidi[c][1];
    //за всяка една взимаме доставено количество
    dostavenoKolichestvo = surovinaDostaveniPartidi[c][2];
    //Logger.log(dostavenaPartida);
    //и използвано количество
    izpolzvanoKolichestvo = getProizvodstvoByPartidi(dostavenaPartida);
    //ако използвано е по-малко от доставено, го добавяме към новия масив
    ostanaloKolichestvo = dostavenoKolichestvo-izpolzvanoKolichestvo;
    if( ostanaloKolichestvo>0 ){
      lastThreePartidesWithQty.push( [dostavenaPartida,ostanaloKolichestvo] );
    }
  }
  //сортираме масива по номер партида
  Logger.log(lastThreePartidesWithQty);

  //ако има по-малко от 3 елемента в масива - добавяне на 0 количества и партиди
  lastThreePartidesWithQty.push( [0,0],[0,0],[0,0] );
  return lastThreePartidesWithQty;

}

/**
 * функция за пресмятане на всички количества от конкретна суровина,
 * участвали в производство
 * приема име на суровина
 * 
 * връща кг
 */
function getAllUsedQtyBySurovina(imeSurovina){

  //взимаме всички данни
  var data = sheetProizvodstvo.getDataRange().getValues();
  var izpolzvanoKolichestvo = 0;
  for(i in data){
    if( data[i][2] == imeSurovina ){
      izpolzvanoKolichestvo = izpolzvanoKolichestvo+data[i][6];
    }
  }
  return izpolzvanoKolichestvo;
}

/**
 * функция за взимане на всички партиди с количества и имена
 * от доставки за конкретна рецепта
 * ??? не знам дали ще се използва
 * 
 * връща масив с елементи партида, име продукт, количество
 */
function getAllDostavkiByRecepta(productRecepta){
  //от тук започваме да изчисляваме какви количества суровини имаме по партиди
  //първо взимаме всички доставени продукти и партиди от рецептата
  var currentSurovinaDostavki = [];
  for( a in productRecepta ){
    currentSurovinaName = productRecepta[a][1];
    //Logger.log(currentSurovinaName)
    currentSurovinaDostavki.push( [getSurovina(currentSurovinaName)] );

  }
  //тук имаме масив с всички партиди, суровини и доставени количества за конкретния краен продукт за производство
  //	[[[[1001.0, Vitamin C, 25.0]]], [[[1002.0, Kurkuma, 30.0], [1003.0, Kurkuma, 30.0], [1004.0, Kurkuma, 44.0], 
  //[1005.0, Kurkuma, 66.0], [1006.0, Kurkuma, 77.0], [1007.0, Kurkuma, 88.0], [1008.0, Kurkuma, 99.0], 
  //[1009.0, Kurkuma, 678.0]]], [[]]]
  //Logger.log(currentSurovinaDostavki);
  return currentSurovinaDostavki
}

/**
 * функция за взимане на количества 
 * от конкретна партида
 * участвала в производството досега
 * или колко от нея е използвано досега
 */
function getProizvodstvoByPartidi(partida){
  //за тест 
  //partida = 1001;
  //взимаме всички данни
  var data = sheetProizvodstvo.getDataRange().getValues();
  //Logger.log(data)
  var izpolzvanoKolichestvo = 0;
  for(i in data){
    //партидите са в трета колонка
    if( data[i][3] == partida ){
      izpolzvanoKolichestvo = izpolzvanoKolichestvo+data[i][6];
    }
  }

  //Logger.log(izpolzvanoKolichestvo)
  return izpolzvanoKolichestvo;

}

/**
 * функцията връща общо килограми доставени от дадена суровина
 * приема името на суровина
 * връща кг
 */
function getSurovinaDostaveniKg(surovinaName){
  var data = sheetDostavki.getDataRange().getValues();
  var kg = 0;
  for( i in data ){
    if( data[i][1]==surovinaName ){
      kg = kg + data[i][2];
    }
  }
  return kg;
}


/**
 * функция за взимане на всички доставени партиди и количества от дадена суровина
 * по име на суровината
 * връща масив от елементи (партида,име,количество)
 */
function getSurovina(surovinaName){
  //var surovinaDostavki = ['test', 1, 1005];
  var data = sheetDostavki.getDataRange().getValues();
  //Logger.log(surovinaName);
  var surovinaDostavki = [];
  for( i in data ){
    if( data[i][1] == surovinaName ){
      //Logger.log(surovinaName);
      dostavkaLot = data[i][0];
      dostavkaProdukt = data[i][1];
      dostavkaQty = data[i][2];
      surovinaDostavki.push( [dostavkaLot, dostavkaProdukt, dostavkaQty] );
    }
  }
  return surovinaDostavki;
}
