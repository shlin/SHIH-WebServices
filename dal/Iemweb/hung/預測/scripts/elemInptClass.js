// Copyright 1998,1999 Macromedia, Inc. All rights reserved.

//Constructs a multiple choice element
function AO_inpt(theParent, theName, theInitialValue,
                 theExpectedValue, theIsCorrect, theScore) {
  // properties
  this.initialValue = theInitialValue;
  this.value = '';
  this.disabled = true;
  
  this.expectedValue = theExpectedValue;
  this.isCorrect = theIsCorrect;
  this.score = theScore;
  this.selected = false;
  
  this.isRadioList = false;
  
  this._parent = theParent;
  this._name = theName;
  this._obj = '';
  
  this.c = new Array(this); // NOTE: choice info stored on the element.

  // member functions
  this.init = AO_inptInit;
  this.reset = AO_inptReset;
  this.enable = AO_inptEnable;
  this.disable = AO_inptDisable;
  this.update = AO_inptUpdate;
  this.setDisabled = AO_inptSetDisabled;
  this.redraw = AO_inptRedraw;
  this.validValue = AO_inptValidValue;
  this.setValue = AO_inptSetValue;
  this.setSelected = AO_inptSetSelected;
  this.changeValue = AO_inptChangeValue;
}

// Initializes the element, special case radio lists
function AO_inptInit() {
  var rlist, i, pos=0;
  with (this) { 
    _obj = AO_findObject(_parent._self + _name + "Inp");
    if (!_obj) { // assume radio
      rlist = AO_findObject(_parent._self + "RadioInp");
      if (rlist && rlist.length != null) {
          for (i in _parent.e) if (i != 'length') // get our element position
            if (_parent.e[i] == this) break; else pos++;
          if (pos < rlist.length) _obj = rlist[pos];  // get radio at same position
          isRadioList = true;
  } } } 
}

//Resets the element
function AO_inptReset() {
  var isChanged = '';
  with (this) {
    isChanged = (value != initialValue);
    value = initialValue;
    _parent.disabled ? disable() : enable();
    validValue();
    redraw();
    if (isChanged && this.onChange != null) onChange(_parent._self+_name, value);
  }
}

//Enables the element
function AO_inptEnable() {
  if (this._obj) with (this) {
    disabled = false;
    redraw();
  }
}

//Calls the approppriate disable or enable function
function AO_inptSetDisabled(theDisabled) {
  if (theDisabled) this.disable();
  else this.enable();
}

//Disables the element
function AO_inptDisable() {
  this.disabled = true;
  this.redraw();
}

//Called by onClick event to update this elements value
function AO_inptUpdate() {
  var noJudge = false;
  with (this) {
    if (disabled) {
      if (!isRadioList) 
        redraw();
      else
        for (var i in _parent.e) if (i != 'length')
          _parent.e[i].redraw();
      return;
    }
  
    if (_obj.checked != null) {
      if (isRadioList && value == _obj.checked) noJudge = true; //IE3.0 oddity
      changeValue((_obj.checked) ? true : false);  //IE3.0 oddity
    } else
      changeValue(_parent.allowMultiSel ? !value : true);
  
    // call the parent's update
    _parent.update(noJudge);
  }
}

//Sets the checked state of the form element
function AO_inptRedraw() {
  if (this._obj) with (this) {
    if (_obj.disabled != null) _obj.disabled = disabled;
    if (isRadioList) {
      if (value) _obj.checked = true;
    } else if (_obj.checked != null) _obj.checked = value;
  }
}

//Checks the value with the expectedValue
function AO_inptValidValue() {
  this.selected = (this.value == this.expectedValue);
  return this.selected;
}

//Internal routine for changing element value
function AO_inptChangeValue(theValue) {
  var i, isChanged = '', isReset = '';
  with (this) {
    isChanged = (value != theValue);
    if (!_parent.allowMultiSel || isRadioList || _obj.type == 'radio') {
      value = theValue;
      for (i in _parent.e) if (i != 'length') with (_parent) {
        if (e[i] != this) {
          isReset = (e[i].value != false);
          e[i].value = false;
        }
        e[i].validValue();
        e[i].redraw();
        if (e[i] != this && isReset && e[i].onChange != null)
          e[i].onChange(e[i]._parent._self+e[i]._name, e[i].value);
      }
    } else {
      value = theValue;
      validValue();
      redraw();
    }
    if (isChanged && this.onChange != null) onChange(_parent._self+_name, value);
  }
}

//Sets the state of the element to the given value
function AO_inptSetValue(theValue) {
  with (this) {
    changeValue(theValue);
    _parent.update(true); // update int, but don't judge
  }
}

//Sets this element to its selected state
function AO_inptSetSelected(theSelected) {
  if (theSelected)
    this.setValue(this.expectedValue);
  else
    this.setValue(!this.expectedValue);
}