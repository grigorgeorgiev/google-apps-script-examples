/**
*
* файл от видео https://www.youtube.com/watch?v=o3AL7ASI_cA
* при който формата е създадена в Google Forms
* а не директно в таблицата
*/

//spreadsheet id of table
var ssID = "1U5vSyKbbJc5S2NVwilC8FsPK03OaPl8R0xKNZPye3qE";
//form nova dostavka id
var formNDID = "1Dhp5uLEXUkG4q2dKfAyJALX3Yz-9h1CPZIi7ZWeIoOk";

// взимаме информация от таблица products
var wsData = SpreadsheetApp.openById(ssID).getSheetByName("produkti");
var form = FormApp.openById(formNDID);

function main(){

  //взима заглавките в страницата
  var labels = wsData.getRange(1,1,1,wsData.getLastColumn()).getValues()[0];
  //върти заглавията в цикъл
  labels.forEach(function(label,i){
    //взима списъка с опции за дропдаун менюто от под всяка заглавка в таблица produkti
    var options = wsData
                  .getRange(2,i+1,wsData.getLastRow()-1,1)
                  .getValues()
                  .map(function(o){ return o[0] })
                  .filter(function(o){ return o !== "" });
    //ъпдейтва падащото меню с инфото под всеки тайтъл в таблица produkti
    updateNDDropdownUsingTitle(label,options);
    // Logger.log(label);
    // Logger.log(options);
  });

  //Logger.log(labels);

}

function updateNDDropdownUsingTitle(title,values) {

  //взима елементите от формата за нова доставка
  //в случая са две дропдаун менюта
  var items = form.getItems();
  //Logger.log(items); //[Item, Item]
  //взимаме техните заглавни редове и ги мапваме в масив ??? как работи това
  var titles = items.map(function(item){
    return item.getTitle();
  });
  //Logger.log(titles); //[Продукт, second question]
  //Logger.log(title); //product name
  //взимаме позицията им в масива - ако са две вероятно е 0 и 1
  //за съответното заглавие, за което е извикана функцията
  var pos = titles.indexOf(title);
  //Logger.log(pos); //-1.0

  var item = items[pos];
  //Logger.log(item); //null
  var itemID = item.getId();

  updateNDDropdown(itemID,values);

}

function updateNDDropdown(id, values) {

  var item = form.getItemById(id);
  item.asListItem().setChoiceValues(values);
  //var items = form.getItems();
  //Logger.log(items[0].getId().toString());

}
