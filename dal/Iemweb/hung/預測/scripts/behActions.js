// Copyright 1998,1999 Macromedia, Inc. All rights reserved.

function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
function MM_changeProp(objStrNS,objStrIE,theProp,theValue) { //v2.0
  var NS = (navigator.appName == 'Netscape');
  var objStr = (NS)?objStrNS:objStrIE;
  if (( NS && (objStr.indexOf('document.layers[')!=0 || document.layers!=null)) ||
      (!NS && (objStr.indexOf('document.all[')   !=0 || document.all   !=null))) {
    var obj = eval(objStr);
    if ((obj != null) && (theProp.indexOf("style.") != 0 || obj.style != null)) {
      eval(objStr+'.'+theProp + '="'+theValue+'"');
  } }
}
function MM_checkBrowser(NSvers,NSpass,NSnoPass,IEvers,IEpass,IEnoPass,OBpass,URL,altURL) { //v2.0
  var newURL = '', version = parseFloat(navigator.appVersion);
  if (navigator.appName.indexOf('Netscape') != -1) {
    if (version >= NSvers) {if (NSpass>0) newURL = (NSpass==1)?URL:altURL;}
    else {if (NSnoPass>0) newURL = (NSnoPass==1)?URL:altURL;}
  } else if (navigator.appName.indexOf('Microsoft') != -1) {
    if (version >= IEvers) {if (IEpass>0) newURL = (IEpass==1)?URL:altURL;}
    else {if (IEnoPass>0) newURL = (IEnoPass==1)?URL:altURL;}
  } else if (OBpass>0) newURL = (OBpass==1)?URL:altURL;
  if (newURL) {
    window.location = unescape(newURL);
    document.MM_returnValue = false;
  }
}
function MM_checkPlugin(plugin, theURL, altURL, IEGoesToURL) { //v2.0
  if ((navigator.plugins && navigator.plugins[plugin]) || //if NS, or
      (IEGoesToURL &&  //if flag set, and MSIE browser for Win95/NT (ActiveX)
       navigator.appName.indexOf('Microsoft') != -1 &&
       navigator.appVersion.indexOf('Mac') == -1 &&
       navigator.appVersion.indexOf('3.1') == -1)) {
    if (theURL.length>2) window.location = theURL;
  } else {
    if (altURL.length>2) window.location = altURL;
  }
  document.MM_returnValue = false;
}
function MM_controlShockwave(objStrNS,objStrIE,cmdName,frameNum) { //v2.0
  var objStr = (navigator.appName=='Netscape')?objStrNS:objStrIE;
  if ((objStr.indexOf('document.layers[')==0 && document.layers==null) ||
      (objStr.indexOf('document.all[')   ==0 && document.all   ==null))
    objStr = 'document'+objStr.substring(objStr.lastIndexOf('.'),objStr.length);
  if (eval(objStr) != null)
    eval(objStr+'.'+cmdName+'('+((cmdName=='GotoFrame')?frameNum:'')+')');
}
function MM_controlSound(sndAction,_sndObj) { //v2.0
  var sndObj = eval( _sndObj );
  if (sndObj != null) {
    if (sndAction=='stop') {
      sndObj.stop();
    } else {
      if (navigator.appName == 'Netscape' ) {
         sndObj.play(false);
      } else {
         if (document.MM_WMP_DETECTED == null) {
            document.MM_WMP_DETECTED = false;
            var i;
            for( i in sndObj )
               if ( i == "ActiveMovie" ) {
                  document.MM_WMP_DETECTED = true;
                  break; }
         }
         if (document.MM_WMP_DETECTED)
            sndObj.play();
         else if ( sndObj.FileName )
            sndObj.run();
}}}}
function MM_displayStatusMsg(msgStr) { //v2.0
  status=msgStr;
  document.MM_returnValue = true;
}
function MM_goToURL() { //v2.0
  for (var i=0; i< (MM_goToURL.arguments.length - 1); i+=2) //with arg pairs
    eval(MM_goToURL.arguments[i]+".location='"+MM_goToURL.arguments[i+1]+"'");
  document.MM_returnValue = false;
}
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
function MM_popupMsg(theMsg) { //v2.0
  alert(theMsg);
}
function MM_preloadImages() { //v2.0
  if (document.images) {
    var imgFiles = MM_preloadImages.arguments;
    if (document.preloadArray==null) document.preloadArray = new Array();
    var i = document.preloadArray.length;
    with (document) for (var j=0; j<imgFiles.length; j++) if (imgFiles[j].charAt(0)!="#"){
      preloadArray[i] = new Image;
      preloadArray[i++].src = imgFiles[j];
  } }
}
function MM_showHideLayers() { //v2.0
  var i, visStr, args, theObj;
  args = MM_showHideLayers.arguments;
  for (i=0; i<(args.length-2); i+=3) { //with arg triples (objNS,objIE,visStr)
    visStr   = args[i+2];
    if (navigator.appName == 'Netscape' && document.layers != null) {
      theObj = eval(args[i]);
      if (theObj) theObj.visibility = visStr;
    } else if (document.all != null) { //IE
      if (visStr == 'show') visStr = 'visible'; //convert vals
      if (visStr == 'hide') visStr = 'hidden';
      theObj = eval(args[i+1]);
      if (theObj) theObj.style.visibility = visStr;
  } }
}
function MM_swapImgRestore() { //v2.0
  if (document.MM_swapImgData != null)
    for (var i=0; i<(document.MM_swapImgData.length-1); i+=2)
      document.MM_swapImgData[i].src = document.MM_swapImgData[i+1];
}
function MM_swapImage() { //v2.0
  var i,j=0,objStr,obj,swapArray=new Array,oldArray=document.MM_swapImgData;
  for (i=0; i < (MM_swapImage.arguments.length-2); i+=3) {
    objStr = MM_swapImage.arguments[(navigator.appName == 'Netscape')?i:i+1];
    if ((objStr.indexOf('document.layers[')==0 && document.layers==null) ||
        (objStr.indexOf('document.all[')   ==0 && document.all   ==null))
      objStr = 'document'+objStr.substring(objStr.lastIndexOf('.'),objStr.length);
    obj = eval(objStr);
    if (obj != null) {
      swapArray[j++] = obj;
      swapArray[j++] = (oldArray==null || oldArray[j-1]!=obj)?obj.src:oldArray[j];
      obj.src = MM_swapImage.arguments[i+2];
  } }
  document.MM_swapImgData = swapArray; //used for restore
}
function MM_validateForm() { //v2.0
  var i,objStr,field,theCheck,atPos,theNum,colonPos,min,max,errors='';
  for (i=0; i<(MM_validateForm.arguments.length-2); i+=3) {
    objStr = MM_validateForm.arguments[(navigator.appName == 'Netscape')?i:i+1];
    if ((objStr.indexOf('document.layers[')==0 && document.layers==null) ||
        (objStr.indexOf('document.all[')   ==0 && document.all   ==null))
      objStr = 'document'+objStr.substring(objStr.substring(0,objStr.lastIndexOf('.')).
                 lastIndexOf('.'),objStr.length);  //fix layer ref if not supp
    field = eval(objStr);
    field.name = (field.name)?field.name:objStr;
    theCheck = MM_validateForm.arguments[i+2];
    if (field.value) { //IF NOT EMPTY FIELD
      if (theCheck.indexOf('isEmail') != -1) { //CHECK EMAIL
        atPos = field.value.indexOf('@');
        if (atPos < 1 || atPos == (field.value.length - 1))
          errors += '- '+field.name+' must contain an e-mail address.\n';
      } else if (theCheck != 'R') { //START NUM CHECKS
        theNum = parseFloat(field.value);
        if (field.value != ''+theNum) errors += '- '+field.name+' must contain a number.\n';
        if (theCheck.indexOf('inRange') != -1) { //CHECK RANGE
          colonPos = theCheck.indexOf(':');
          min = theCheck.substring(8,colonPos);
          max = theCheck.substring(colonPos+1,theCheck.length);
          if (theNum < min || max < theNum) //bad range
            errors += '- '+field.name+' must contain a number between '+min+' and '+max+'.\n';
    } } }
    else if (theCheck.charAt(0) == 'R') errors += '- '+field.name+' is required.\n';
  }
  if (errors) alert('The following error(s) occurred:\n'+
                    errors);
  document.MM_returnValue = (errors == '')
}
