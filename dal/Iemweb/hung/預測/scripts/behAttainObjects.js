// Copyright 1998,1999 Macromedia, Inc. All rights reserved.

var MSG_notracking = "Tracking system not found";
function MM_cmiCheckInstall(win) {
  // Done since NS blows away window variables on resize
  if (win==null) win = window;
  if (win) {
    if (window.CMIIsPresent == null) {
      if (findcmiframe != null) {
        var frm = findcmiframe(null);
        if (frm != null) frm.installcmi(win);
        else if (installcmi != null) {
          installcmi(win);
          cmiinit(win);
      } }
      if (win.CMIInitialize != null) win.CMIInitialize();
  } }
  return (win.CMIIsPresent != null);
}
function MM_cmiSendInteractionInfo(date, time, intid, objid, intrtype, correct, student, result, weight, latency) { //v1.22
  var aDt= new Date();
  var curHr=aDt.getHours()+'', curMin=aDt.getMinutes()+'', curSec=aDt.getSeconds()+'';
  var curDay=aDt.getDate()+'', curMonth=aDt.getMonth()+1+'', curYear=aDt.getYear(), dmy;

  if (curYear < 1900) curYear += 1900;
  if (curDay.toString().length==1) curDay = '0'+curDay;
  if (curMonth.toString().length==1) curMonth = '0'+curMonth;
  if (curHr.toString().length==1) curHr = '0'+curHr;
  if (curMin.toString().length==1) curMin = '0'+curMin;
  if (curSec.toString().length==1) curSec = '0'+curSec;
  tim=curHr+":"+curMin+":"+curSec;
  dmy=curDay+"/"+curMonth+"/"+curYear;
  if (date=='') date=dmy;
  if (time=='') time=tim;

  if (MM_cmiCheckInstall(window)) {
    if (CMIIsPresent()) {
      CMIAddInteraction(date, time, intid, objid, intrtype, correct, student, result, weight, latency);
      return;
  } }
}
function MM_cmiSendObjectiveInfo(theInt, index, objid, score, status) { //v1.22
  if (MM_cmiCheckInstall(window)) {
    if (CMIIsPresent()) {
      if (theInt) {
        objid = eval(theInt+".trackObjectiveId");
        score = eval(theInt+".score");
      }
      CMISetObj(index, objid, score, status, '', '', '', '');
      return;
  } }
}
function MM_cmiSendScore(theInt, theScore) { //v1.22
  if (MM_cmiCheckInstall(window)) {
    if (CMIIsPresent()) {
      if (theInt) theScore = eval(theInt+'.score');
      CMISetScore(theScore);
      return;
  } }
}
function MM_cmiSetLessonStatus(theStatus) { //v1.22
  if (MM_cmiCheckInstall(window)) {
    if (CMIIsPresent()) {
      CMISetStatus(theStatus);
      return;
  } }
}
function MM_cmiSetTime(theInt, theSeconds) { //v1.22
  if (MM_cmiCheckInstall(window)) {
    if (CMIIsPresent()) {
      if (theInt) theSeconds = eval(theInt+'.getTime()');
      CMISetTime(theSeconds);
      return;
  } }
}
function MM_judgeAttainObj(interactionId) { //v1.0
  eval(interactionId+".judge()");
}
function MM_resetAttainObj(interactionId,method,item) { //v1.0
  if (item!=null && item)
    if (method=="resetElems") {method="e['"+item+"'].reset"; item=""}
    else item = "'"+item+"'";
  else item="";
  eval(interactionId+"."+method+"("+item+")");
}
function MM_setAttainObjProps(jsStr) { //v1.0
  eval(jsStr);
}
function MM_initAttainObj() { //v1.0
  if (window.initAttainObjFns) {eval(window.initAttainObjFns); window.initAttainObjFns = '';}
}
function MM_writeToFrame(frameRef,newHTML,preserveBg) { //v2.0
  var bodyAttr="", frameObj=eval(frameRef);
  if (frameObj) with (frameObj.document) { //if frame found
    if (preserveBg) bodyAttr = " BGCOLOR='"+bgColor+"' TEXT='"+fgColor+"'";
    write("<HTML><BODY"+bodyAttr+">"+unescape(newHTML)+"</BODY></HTML>");
    close();
  }
}
function MM_writeToLayer(NSlayerRef,IElayerRef,newText) { //v2.0
  var layerObj, NS = (navigator.appName=="Netscape");
  if (document.layers || document.all) layerObj = eval((NS)?NSlayerRef:IElayerRef);
  if (layerObj) with (layerObj) { //if layer found
    if (NS) {document.write(unescape(newText)); document.close();}
    else innerHTML = unescape(newText);
  }
}
function MM_writeToPopup(msg) { //v2.0
  alert(msg);
}
function MM_writeToTextfield(NSobj,IEobj,newText) { //v2.0
  var theObj = eval((navigator.appName=="Netscape")?NSobj:IEobj);
  if (theObj) theObj.value = newText;
}
