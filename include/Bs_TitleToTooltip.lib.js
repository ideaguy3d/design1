
var bs_ttt_class;var bs_ttt_style;function bs_ttt_showTitle() {
var elm = event.srcElement;if (bs_isEmpty(elm.title)) return;var div = bs_ttt_getTitleElm();if (div == false) return;div.innerHTML     = elm.title;elm.title = '';var pos = getAbsolutePos(elm);var elmHeight = elm.offsetHeight;div.style.left    = pos.x;div.style.top     = parseInt(pos.y) + parseInt(elmHeight) +0;div.style.display = 'block';}
function bs_ttt_hideTitle() {
var elm = event.srcElement;var div = bs_ttt_getTitleElm();if (div == false) return;elm.title = div.innerHTML;div.innerHTML     = '';div.style.display = 'none';}
function bs_ttt_getTitleElm() {
var elm = document.getElementById('bs_ttt_dynamicTitle');if (elm == null) {
try {
var divTagStr = '<div';divTagStr += ' id="bs_ttt_dynamicTitle"';if (typeof(bs_ttt_class) != 'undefined') {
divTagStr += ' class="' + bs_ttt_class + '"';}
divTagStr += ' style="position:absolute; display:none; cursor:default;';if (typeof(bs_ttt_style) != 'undefined') {
divTagStr += bs_ttt_style;} else if (typeof(bs_ttt_class) == 'undefined') {
divTagStr += ' background-color:#FFFFE7; font-family:arial,helvetica; font-size:11px; border:1px solid black; padding:1px;';}
divTagStr += '"></div>';var bodyTag = document.getElementsByTagName('body');bodyTag = bodyTag[0];bodyTag.insertAdjacentHTML('beforeEnd', divTagStr);elm = document.getElementById('bs_ttt_dynamicTitle');if (elm == null) return false;} catch (e) {
return false;}
}
return elm;}
function bs_ttt_initAll() {
var elms = document.getElementsByTagName('input');for (var i=0; i<elms.length; i++) {
if (!bs_isEmpty(elms[i].title)) {
elms[i].attachEvent('onfocus', bs_ttt_showTitle);elms[i].attachEvent('onblur',  bs_ttt_hideTitle);}
}
}
