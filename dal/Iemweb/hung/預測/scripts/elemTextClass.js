// Copyright 1998,1999 Macromedia, Inc. All rights reserved.

//Constructs a text entry element
function AO_text(theParent, theName, theInitialValue) {
  // properties
  this.initialValue = theInitialValue?theInitialValue:'';
  this.value = null;
  this.disabled = true;
  
  this._parent = theParent;
  this._name = theName;
  this._obj = '';
  
  this.c = new Array();

  // member functions
  this.init = AO_textInit;
  this.reset = AO_textReset;
  this.enable = AO_textEnable;
  this.disable = AO_textDisable;
  this.setDisabled = AO_textSetDisabled;
  this.update = AO_textUpdate;
  this.redraw = AO_textRedraw;
  this.setValue = AO_textSetValue;
  this.focus = AO_textFocus;
}

//Initializes the element
function AO_textInit() {
  with (this) _obj = AO_findObject(_parent._self + _name + "Inp");
}

//Resets the element
function AO_textReset() {
  with (this) {
    _parent.disabled ? disable() : enable();
    setValue(initialValue);
  }
}

//Enables the element
function AO_textEnable() {
  var i;
  if (this._obj) with (this) {
    disabled = false;
    for (i in c) if (i != 'length') c[i].disabled = false;
    setValue(_obj.value);
  }
}

//Disables the element
function AO_textDisable() {
  this.disabled = true;
  this.redraw();
}

//Calls the approppriate disable or enable function
function AO_textSetDisabled(theDisabled) {
  if (theDisabled) this.disable();
  else this.enable();
}

//Called by onClick event to update this elements value
//We also store the object reference of the control, so that
// we can reset it later.
function AO_textUpdate() {
  var isChanged = '';
  if (!this.disabled) with (this) {
    _obj.value = AO_textStripSpaces(_obj.value);
    isChanged = (value != _obj.value);
    value = _obj.value;
    for (var i in c) if (i != 'length') c[i].validValue();
    if (isChanged && this.onChange != null) onChange(_parent._self+_name, value);
    _parent.update();   // call the parent's update
  }
}

// Sets the text value
function AO_textRedraw() {
  if (this._obj) with (this) {
    if (_obj.disabled != null) _obj.disabled = disabled;
    if (_obj.value != null) _obj.value = value;
  }
}

function AO_textSetValue(theValue) {
  var isChanged = '';
  with (this) {
    theValue = AO_textStripSpaces(theValue);
    isChanged = (value != theValue);
    value = theValue;
    redraw();
    for (var i in c) if (i != 'length') c[i].validValue();
    if (isChanged && this.onChange != null) onChange(_parent._self+_name, value);
    _parent.update(true);  // update int, no judge
  }
}

function AO_textFocus() {
  if (this.disabled) with (this._parent)
    if (browserIsNS && browserVersion >= 4.0 && osIsWindows) this._obj.blur();
}


//////////////////////////////////////////
//Create a string choice object
function AO_textComp(theParent, theElement, theName,
                     theExpectedValue, theIsCorrect, theScore,
                     theMatchCase, theMatchAll)
{
  // properties
  this.expectedValue = theExpectedValue;
  this.isCorrect = theIsCorrect;
  this.score = theScore;
  this.selected = false;
  this.disabled = false;
  
  this.matchCase = theMatchCase;
  this.matchAll = theMatchAll;
  
  this._elem = eval(theParent._self+".e['"+theElement+"']");
  this._isChoice = true;

  // methods
  this.validValue = AO_textCompValidValue;
  this.setSelected = AO_textCompSetSelected;
  this.setDisabled = AO_textCompSetDisabled;
  this.deencrypt = AO_textDeencrypt;

}

function AO_textCompValidValue() {
  var theValue, expValue;
  with (this) {
    selected = false;
    if (!disabled) {
      theValue = _elem.value;
      expValue = deencrypt(expectedValue);

      if (!matchCase) {
        theValue = theValue.toUpperCase();
        expValue = expValue.toUpperCase();
      }

      if (theValue != '') {
        if (matchAll)  selected = (theValue == expValue);
        else  selected = (theValue.indexOf(expValue) != -1);
  } } }
  return this.selected;
}

function AO_textCompSetSelected(theSelected) {
  with (this) {
    if (theSelected) {
      _elem.setValue(deencrypt(expectedValue));
    } else {
      selected = false;
      _elem._parent.update(true); // update int, no judge
  } }
}

function AO_textCompSetDisabled(theDisabled) {
  with (this) {
    disabled = theDisabled;
    validValue();
    _elem._parent.update(true); // update interaction, no judge
  }
}

function AO_textDeencrypt(theStr) {
  var decipher='',i,key,clength,part='-###-',keyOffsetObscure,keyOffset='',hyphen;
  
  if (theStr.indexOf(part)!=-1) {
    strStart = theStr.indexOf(part)+part.length;
  hyphen=theStr.indexOf('-');
    key = theStr.substring(0,hyphen);
  keyOffsetObscure=theStr.substring(hyphen+1, theStr.indexOf('-',hyphen+1));
    for (i=keyOffsetObscure.length-1;i>=0;i--)
    keyOffset+=keyOffsetObscure.charAt(i);
    clength = key-keyOffset;
    retVal = theStr.substring(strStart, theStr.length);
  for (i=0;i<retVal.length;i+=(clength+1))
    decipher=decipher+retVal.charAt(i);

  retVal = decipher;
  }
  else retVal = theStr;
    
  return retVal;
}

function AO_textStripSpaces(theStr) {
  var c,b=0,e=theStr.length-1,E=e;
  while (b<E&&((c=theStr.charAt(b))==" " || c=="\r" || c=="\n")) b++;
  while (e>0&&((c=theStr.charAt(e))==" " || c=="\r" || c=="\n")) e--;
  return theStr.substring(b,e+1);
}
