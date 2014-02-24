// Copyright 1998,1999 Macromedia, Inc. All rights reserved.

//Constructs a image button element
function AO_ibtn(theParent, theName, theInitialValue,
                 theExpectedValue, theIsCorrect, theScore,
                 theIsToggle, theStateMask)
{
  // properties
  this.initialValue = theInitialValue;
  this.value = '';
  this.disabled = true;

  this.expectedValue = theExpectedValue;
  this.isCorrect = theIsCorrect;
  this.score = theScore;
  this.selected = false;

  this.isToggle = theIsToggle;

  this._parent = theParent;
  this._name = theName;
  this._obj = '';
  this._stateMask = theStateMask ? theStateMask : '';
  this._mouseOver = false;
  this._mouseDown = false;
  this._filePre = '';
  this._fileExt = '';

  this.c = new Array(this); // NOTE: choice info stored on the element.

  // member functions
  this.init = AO_ibtnInit;
  this.reset = AO_ibtnReset;
  this.enable = AO_ibtnEnable;
  this.disable = AO_ibtnDisable;
  this.setDisabled = AO_ibtnSetDisabled;
  this.update = AO_ibtnUpdate;
  this.redraw = AO_ibtnRedraw;
  this.validValue = AO_ibtnValidValue;
  this.changeValue = AO_ibtnChangeValue;
  this.setValue = AO_ibtnSetValue;
  this.setSelected = AO_ibtnSetSelected;
  this.setIsToggle = AO_ibtnSetIsToggle;
}

// Initializes the element
function AO_ibtnInit() {
  var theSrc, extIndex, maskArray, extArray, i;
  with (this) {
    if (! _stateMask) return;

    _obj = AO_findObject(_parent._self + _name + "Btn");
    if (_obj && _obj.src != null) {
      theSrc = _obj.src;
      extIndex = theSrc.lastIndexOf(".");
      if (extIndex != -1) { // save the extension
        _fileExt = theSrc.substring(extIndex, theSrc.length);
        theSrc = theSrc.substring(0, extIndex); // remove extension
      }
      if (theSrc.lastIndexOf("_sel") != -1)
        _filePre = theSrc.substring(0, theSrc.lastIndexOf("_sel"));
      else if (theSrc.lastIndexOf("_hlt") != -1)
        _filePre = theSrc.substring(0, theSrc.lastIndexOf("_hlt"));
      else if (theSrc.lastIndexOf("_dis") != -1)
        _filePre = theSrc.substring(0, theSrc.lastIndexOf("_dis"));
      else
        _filePre = theSrc;
        
      // preload the images
      maskArray = new Array('s', 'S', 'h', 'H', 'd', 'D');
      extArray = new Array("", "_sel", "_hlt", "_sel_hlt", "_dis", "_sel_dis"); 
      for (i=0; i < maskArray.length; i++)
        if (_stateMask.indexOf(maskArray[i]) != -1)
          MM_preloadImages(_filePre + extArray[i] + _fileExt);
  } }
}

//Resets the element
function AO_ibtnReset() {
  var isChanged = '';
  with (this) {
    isChanged = (value != initialValue);
    _mouseOver = false;
    _mouseDown = false;
    value = initialValue;
    _parent.disabled ? disable() : enable();
    validValue();
    redraw();
    if (isChanged && this.onChange != null) onChange(_parent._self+_name, value);
  }
}

//Enables the element
function AO_ibtnEnable() {
  if (this._obj) with (this) {
    disabled = false;
    redraw();
  }
}

//Disables the element
function AO_ibtnDisable() {
  this.disabled = true;
  this.redraw();
}

//Calls the approppriate disable or enable function
function AO_ibtnSetDisabled(theDisabled) {
  if (theDisabled) this.disable();
  else this.enable();
}

//Called by the onClick, onMouseOver, onMouseDown, and onMouseOut
// events of the A tag, to change the button image and state
function AO_ibtnUpdate(theEvent) {
  if (!this.disabled) with (this) {
    if (theEvent != "onclick") {
      if (theEvent == "onmouseover") {
        if (_mouseOver) return;
        _mouseOver = true;
      } else if (theEvent == "onmouseout") {
        if (!_mouseDown && !_mouseOver) return;
        _mouseDown = false;
        _mouseOver = false;
      } else if (theEvent == "onmousedown") {
        if (_mouseDown) return;
        _mouseDown = true;
      }
      redraw();
    } else { // onclick
      _mouseDown = false;
      changeValue((_parent.allowMultiSel) ? !value : true);
      _parent.update();
  } }
}

// Sets the image based on the button state
function AO_ibtnRedraw() {
  if (this._filePre) with (this) {
    var imageIndex = 's';
    var imageExt = '';
    if (disabled) {
      imageIndex = 'd';
      imageExt += "_dis";
    } else if (_mouseOver) {
      imageIndex = 'h';
      imageExt += "_hlt";
    }
    if ((value && isToggle) || _mouseDown) {
      if (_stateMask.indexOf(imageIndex.toUpperCase()) != -1) {
        imageExt = "_sel" + imageExt;
      } else if (_stateMask.indexOf(imageIndex) == -1) {
        // unselected images not found
        if (_stateMask.indexOf('S') != -1) imageExt = "_sel";
        else imageExt = '';
      }
    } else if (_stateMask.indexOf(imageIndex) == -1) imageExt = '';

    var currImageName = _obj.src;
    var imageName = _filePre + imageExt + _fileExt;
    if (currImageName != imageName) _obj.src = imageName;
  }
}

function AO_ibtnValidValue() {
  this.selected = (this.value == this.expectedValue);
  return this.selected;
}

function AO_ibtnChangeValue(theValue) {
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

function AO_ibtnSetValue(theValue) {
  with (this) {
    changeValue(theValue);
    _parent.update(true); // update int, but don't judge
  }
}

function AO_ibtnSetSelected(theSelected) {
  if (theSelected)
    this.setValue(this.expectedValue);
  else
    this.setValue(!this.expectedValue);
}

function AO_ibtnSetIsToggle(theIsToggle) {
  with (this) {
    isToggle = theIsToggle;
    redraw();
  }
}