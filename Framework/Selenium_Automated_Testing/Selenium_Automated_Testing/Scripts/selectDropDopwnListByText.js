function  selectDropDopwnListByText($element, text){var kDropDownList = $element.data('kendoDropDownList');if(kDropDownList !== undefined){var arrayLi = kDropDownList.ul.find('li');for(var i = 0, n = arrayLi.length; i < n; ++i){var $li = $(arrayLi[i]);if($li.html() === text) {
console.log($li);$li.click();return $li;}}}};


