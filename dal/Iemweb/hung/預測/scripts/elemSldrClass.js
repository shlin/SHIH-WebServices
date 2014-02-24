// Copyright 1998,1999 Macromedia, Inc. All rights reserved.

//Constructs a slider element
function AO_sldr(theParent, theName, theInitVal,theMin,theMax,theUseFloat,theIsVertical) {
  // properties
  this.initialValue = Math.min(Math.max(Math.min(theMin,theMax),theInitVal),
                      Math.max(theMax,theMin)); //keep in legal range, may be reversed
  this.value = '';
  this.disabled = true;

  this.min = theMin;
  this.max = theMax;
  this.useFloat = theUseFloat;

  this._parent = theParent;
  this._name = theName;
  this._obj = '';
  this._isVertical = theIsVertical;
  this._trkObj = '';
  this._txtObj = '';
  this._trackLen = 0;
  this._trackPos = 0;
  this._self = theParent._self+".e['"+theName+"']";
  this.c = new Array();

  // member functions
  this.reset = AO_sldrReset;
  this.init = AO_sldrInit;
  this.enable = AO_sldrEnable;
  this.disable = AO_sldrDisable;
  this.setDisabled = AO_sldrSetDisabled;
  this.update = AO_sldrUpdate;
  this.setValue = AO_sldrSetValue;
}

//Resets the element
function AO_sldrReset() {
  var i, NS = (navigator.appName == "Netscape");
  if (this._obj) with (this) {
    if (_parent.disabled) disable();
    else enable();
    value = initialValue;
    for (i in c) c[i].validValue();
    if (_txtObj) _txtObj.value = initialValue; //load text object if there
    setValue(initialValue);
  }
}

//Enables the element
function AO_sldrEnable() {
  if (this._obj) with (this) {
    disabled = false;
    _obj.MM_dragOk = true;
  }
}

//Disables the element
function AO_sldrDisable() {
  this.disabled = true;
  this._obj.MM_dragOk = null;
}

//Calls the approppriate disable or enable function
function AO_sldrSetDisabled(theDisabled) {
  if (theDisabled) this.disable();
  else this.enable();
}

//Called when the layer is dragged or dropped, updates all values.
function AO_sldrUpdate(dragging) {
  var i;
  if (!this.disabled) with (this) {
    if (_isVertical) value = (min+((max-min)*(_trackLen-_obj.MM_UPDOWN-_trackPos)/_trackLen));
    else value = (min+(max-min)*(1-(_trackLen-_obj.MM_LEFTRIGHT-_trackPos)/_trackLen));
    if (!useFloat) value = Math.round(value);
    if (dragging) {
      if (_txtObj) _txtObj.value = value;
      if (this.onDrag != null) onDrag(_parent._self+_name, value);
    } else {
      for (i in c) c[i].validValue(); // walk choices, setting selected
      if (this.onDrop != null) onDrop(_parent._self+_name, value);
      _parent.update(); // call the parent's update
  } }
}

//Initializes slider objects. Looks for drag thumb, track image, and optional textbox.
//Makes the layer draggable by calling MM_dragLayer.
function AO_sldrInit() {
  var baseName,thm,diff,U=0,D=0,L=0,R=0,NS = (navigator.appName == "Netscape");
  with (this) {
    baseName = _parent._self+_name; //assemble base name
    _obj = AO_findObject(baseName+"Inp"); //find slider layer object
    thm = AO_findObject(baseName+"Thm"); //find slider layer object
    if (_obj && thm) {
      _trkObj = AO_findObject(baseName+'Trk');  //find trk image object
      if (_trkObj) {
        _txtObj = AO_findObject(baseName+'Val');  //find optional output textfield obj
        _trackLen = Math.max(((_isVertical)?
          ((NS)?_trkObj.height-thm.height : _trkObj.height-thm.height):
          ((NS)?_trkObj.width-thm.width  : _trkObj.width-thm.width)),1);
        diff = this.max-this.min;
        _trackPos = Math.round(((this.initialValue - this.min)/((diff)?diff:1))*_trackLen);
        if (_isVertical) _trackPos = _trackLen - _trackPos;
        setValue(this.initialValue);
        if (_isVertical) {U = _trackPos; D = _trackLen-_trackPos;}
        else {L = _trackPos; R = _trackLen-_trackPos;}
        MM_dragLayer(_obj,_obj,0,0,0,0,false,false,U,D,L,R,false,false,0,
                     _self+'.update()',true,_self+'.update(true)');
  } } }
}


//Moves the drag thumb to the given position.
function AO_sldrSetValue(newValue) {
  var i, newPos, diff, newPos, posObj, NS = (navigator.appName == "Netscape");
  with (this) {
    newValue = Math.min(Math.max(Math.min(min,max),newValue),
               Math.max(max,min)); //keep in legal range, may be reversed
    diff = max-min;
    newPos = Math.round(((newValue - min)/((diff)?diff:1))*_trackLen);
    if (_isVertical) newPos = _trackLen - newPos;
    posObj = (NS)?_obj:_obj.style;
    posObj[_isVertical?(NS?'top':'pixelTop'):(NS?'left':'pixelLeft')] = newPos;
    if (_txtObj) _txtObj.value = newValue; //load text object if there
    value = newValue;
    for (i in c) c[i].validValue(); // walk choices, setting selected
    if (this.onDrag != null) onDrag(_parent._self+_name, value);
    if (this.onDrop != null) onDrop(_parent._self+_name, value);
    _parent.update(true); // update int, but don't judge
  }
}

//------------------------------------------------------------------

//Constructs a slider range element.
function AO_sldrRnge(theParent, theElement, theName,
               theExpectedValue, theIsCorrect, theScore) {
  // properties
  this.isCorrect = theIsCorrect;
  this.expectedValue = theExpectedValue;
  this.score = theScore;
  this.selected = false;
  this.disabled = false;

  this._elem = eval(theParent._self+".e['"+theElement+"']");
  this._isChoice = true;

  // method
  this.validValue = AO_sldrRgneValidValue;
  this.setSelected = AO_sldrRgneSetSelected;
  this.setDisabled = AO_sldrRgneSetDisabled;
  this.getExpRangeValue = AO_sldrRgneGetExpRangeValue;
}


//Checks if the current slider value is within the current range.
function AO_sldrRgneValidValue() {
  var myValue,expVal,colonPos,fromStr,fromNum,toStr,toNum,retVal=false;
  if (!this.disabled) with (this) {
    fromNum = getExpRangeValue(0);
    toNum   = getExpRangeValue(1);
    myValue = _elem.value;
    if (fromNum <= toNum)
      retVal = (fromNum <= myValue && myValue <= toNum);
    else
      retVal = (toNum <= myValue && myValue <= fromNum);
    selected = retVal;
  }
  return retVal
}


//Averages the range and sets the slider to that average
function AO_sldrRgneSetSelected(theSelected) {
  with (this) {
    if (theSelected) {
      fromNum = getExpRangeValue(0);
      toNum   = getExpRangeValue(1);
      avgNum = fromNum
      if      (fromNum > toNum) avgNum = toNum + (fromNum - toNum)/2;
      else if (fromNum < toNum) avgNum = fromNum + (toNum - fromNum)/2;
      if (!this._elem.useFloat) avgNum = Math.round(avgNum);
      _elem.setValue(avgNum);
    } else {
      selected = false;
      _elem._parent.update(true);
    }
  }
}


function AO_sldrRgneSetDisabled(theDisabled) {
  with (this) {
    disabled = theDisabled;
    if (disabled) selected = false;
    else validValue();
    _elem._parent.update(true);
  }
}


//Given colon separated num string, returns one of the numbers.
//Pass 0 for first num, 1 for second: "33:46" => 33  or 46
function AO_sldrRgneGetExpRangeValue(numIndex) {
  var expVal, colonPos, retVal;

  expVal = this.expectedValue;
  colonPos = expVal.indexOf(":");
  if (colonPos != -1) { //if colon separated, split strings
    retVal = expVal.substring((numIndex)?colonPos+1:0,(numIndex)?expVal.length:colonPos);
  } else { //else, theres a single number, no range
    retVal = expVal;
  }
  return (this._elem.useFloat)?parseFloat(retVal):parseInt(retVal);
}
