//
// JUMP MENU
// onchange('?..... &page=',1-2-3-4-5);
// 

function jump(url1,txt){
    eval("window.location.href='" + url1 + txt.options[txt.selectedIndex].value+"'");
}


//
// SELECT MENU
// onchange('[tablename]'); //append to textarea
// 

function selectel(txt){
   document.forms[0].memo.value+=' ' + txt.options[txt.selectedIndex].value + ' ';
}