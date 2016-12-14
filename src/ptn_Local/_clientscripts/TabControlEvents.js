//<script language="javascript">

// ------------------------------------------------------------
// TTabControl client script support
// ------------------------------------------------------------

//
//--- function to activate the relvant tab, this function 
//--- is called as an event from the various tabs on the onclick 
//--- events of those tabs. It does all the CSS gimmickry
//--- and achieves the desired effect of tab activation
//
function tabActivate( prmCanvasLoc, prmtabHdrLoc, prmtabCntLoc ) 
{
  var i, nChildCount;
  var strChildId;
  var oChild;

  //--- get the number of children in canvas
  nChildCount = prmCanvasLoc.children.length;
  
  //--- loop thru the child objects to manage CONTENT-TABS
  for( i = 0; i < nChildCount; i++ ) 
  {
    //--- retrive child object and id
    oChild = prmCanvasLoc.children(i);
    strChildId = oChild.id;
    
    //--- process only if object is content tab
    if( strChildId.substr(0,6) == "tabCnt" ) 
    { 
      if( oChild.id==prmtabCntLoc.id )
      {
        oChild.style.visibility= 'inherit';
        oChild.style.zIndex= 2;
      }
      else 
      oChild.style.visibility= 'hidden';
    }
  }
  
  //--- loop thru the child objects to manage HEADER-TABS
  for( i = 0; i < nChildCount; i++ ) 
  {
    //--- retrive child object and id
    oChild = prmCanvasLoc.children(i); strChildId = oChild.id

    //--- process only if object is header tab
    if( strChildId.substr(0,6) == "tabHdr" ) 
    { 
      if(oChild.id==prmtabHdrLoc.id) 
      {
        oChild.style.top="4px"; oChild.style.width= "108px";
        oChild.style.borderBottomColor= "buttonface";
        oChild.style.borderBottomStyle= "solid";
        oChild.style.borderBottomWidth= "1px";
        oChild.style.cursor= 'default';
        oChild.style.zIndex= 4;
      }
      else 
      {
        oChild.style.top="8px"; oChild.style.width= "100px";
        oChild.style.cursor= 'hand';
        oChild.style.zIndex= -1;
      }
    }
  }
}
