// Copyright 1998,1999 Macromedia, Inc. All rights reserved.

//Constructs a hot area element
function AO_hota(theParent, theName, theInitialValue,
                 theExpectedValue, theIsCorrect, theScore) {
  // properties
  this.initialValue = theInitialValue;
  this.value = '';
  this.disabled = false;

  this.expectedValue = theExpectedValue;
  this.isCorrect = theIsCorrect;
  this.score = theScore;
  this.selected = false;

  this._parent = theParent;
  this._name = theName;
  this._self = theParent._self+".e['"+theName+"']";
  this._obj = '';
  
  this.c = new Array(this); // NOTE: choice info stored on the element.

  // member functions
  this.reset = AO_hotaReset;
  this.init = AO_hotaInit;
  this.initBackdrop = AO_hotaInitBackdrop;
  this.enable = AO_hotaEnable;
  this.disable = AO_hotaDisable;
  this.setDisabled = AO_hotaSetDisabled;
  this.update = AO_hotaUpdate;
  this.validValue = AO_hotaValidValue;
  this.changeValue = AO_hotaChangeValue;
  this.setValue = AO_hotaSetValue;
  this.setSelected = AO_hotaSetSelected;
  
}

//Resets the element
function AO_hotaReset() {
  var isChanged = '';
  if (this._obj) with (this) {
    isChanged = (value != initialValue);
    value = initialValue;
    disabled = _parent.disabled;
    validValue();
    if (isChanged && this.onChange != null) onChange(_parent._self+_name, value);
  }
}

function AO_hotaInit() {
  with (this) {
    _obj = AO_findObject(_parent._self + _name); //find slider layer object
    if (_obj)
      MM_dragLayer(_obj,_obj,0,0,0,0,false,false,0,0,0,0,false,false,0,_self + '.update()',true);
    initBackdrop();
  }
}

function AO_hotaInitBackdrop() {
  var theObj;
  if (this._parent.bdInited == null) with (this) {
    _parent.bdInited = true;
    theObj = AO_findObject(_parent._self + "backdrop"); // find backdrop layer object
    if (theObj) {
      document.allLayers = null;
      MM_dragLayer(theObj,theObj,0,0,0,0,false,false,0,0,0,0,false,false,0,
                   "AO_hotaUpdateBackdrop(" + _parent._self + ")",true);
  } }
}

//Enables the element
function AO_hotaEnable() {
  if (this._obj) with (this) {
    disabled = false;
  }
}

//Disables the element
function AO_hotaDisable() {
  this.disabled = true;
}

//Calls the approppriate disable or enable function
function AO_hotaSetDisabled(theDisabled) {
  if (theDisabled) this.disable();
  else this.enable();
}

//Called by onClick event to update this elements value
function AO_hotaUpdate() {
  if (!this.disabled) with (this) {
    changeValue((_parent.allowMultiSel) ? !value : true);
    _parent.update();  // call the parent's update
  }
}

function AO_hotaValidValue() {
  this.selected = (this.value == this.expectedValue);
  return this.selected;
}

function AO_hotaChangeValue(theValue) {
  var i, isChanged = '', isReset = '';
  with (this) {
    isChanged = (value != theValue);
    if (!_parent.allowMultiSel) {
      value = theValue;
      for (i in _parent.e) if (i != 'length') with (_parent) {
        if (e[i] != this) {
          isReset = (e[i].value != false);
          e[i].value = false;
        }
        e[i].validValue();
        if (e[i] != this && isReset && e[i].onChange != null) 
          e[i].onChange(e[i]._parent._self+e[i]._name, e[i].value);
      }
    } else {
      value = theValue;
      validValue();
    }
    if (isChanged && this.onChange != null) onChange(_parent._self+_name, value);
  }
}

function AO_hotaSetValue(theValue) {
  with (this) {
    changeValue(theValue);
    _parent.update(true); // update int, but don't judge
  }
}

function AO_hotaSetSelected(theSelected) {
  if (theSelected)
    this.setValue(this.expectedValue);
  else
    this.setValue(!this.expectedValue);
}

function AO_hotaUpdateBackdrop(theParent) {
  if (!theParent.allowMultiSel) {
    theParent.resetElems();
    theParent.update();
  }
}
