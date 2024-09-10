//
//----------
// Open and close click methods for admin nav ajax menu
//----------
//
function AdminNavOpenClick(OffNode,OnNode,ContentNode,NodeID,onEmptyShow,onEmptyHide) {
	document.getElementById(OffNode).style.display='none';
	document.getElementById(OnNode).style.display='block';
	var e=document.getElementById(ContentNode);
	if(e.ok) {//already populated
		cj.ajax.addon('AdminNavigatorOpenNode','nodeid='+NodeID)
		navBindNodes();
	}else{
		e.ok='ok';
		var arg = {contentNode:ContentNode};
		arg.onEmptyHide = onEmptyHide;
		arg.onEmptyShow = onEmptyShow;
		cj.ajax.addonCallback("AdminNavigatorGetNode",'nodeid='+NodeID,AdminNavOpenClickCallback,arg);
		//cj.ajax.addon('AdminNavigatorGetNode','nodeid='+NodeID,'',ContentNode,onEmptyHide,onEmptyShow)
	}
	e.style.display='block';
}
function AdminNavOpenClickCallback(serverResponse, arg ) {
	if (serverResponse == '') {
		if (document.getElementById(arg.onEmptyHide)) { document.getElementById(arg.onEmptyHide).style.display = 'none' }
		if (document.getElementById(arg.onEmptyShow)) { document.getElementById(arg.onEmptyShow).style.display = 'block' }
	} else {
		var el1 = document.getElementById(arg.contentNode);
		if (el1) {
			el1.innerHTML = serverResponse
			navBindNodes();
			};
	}
}
function AdminNavCloseClick(OffNode,OnNode,ContentNode,NodeID,EmptyNode) {
		document.getElementById(OffNode).style.display='none';
		document.getElementById(OnNode).style.display='block';
		document.getElementById(ContentNode).style.display='none';
		cj.ajax.addon('AdminNavigatorCloseNode','nodeid='+NodeID);
}
/*
*	moved to dashboard.js to keep it together. kept this incase dash updated and not adminNav
*/
function navBindNodes() {
	if(!dashBindNavNodes) {
		/* dashBindNavNodes installed, navDrag binding already handled */
	}
	jQuery(".navDrag").each(function(){
		jQuery(this).draggable({
			stop: function(event, ui){
				console.log("adminNav.navBindNodes draggable:stop");
				navDrop(this.id,ui.offset.left,ui.offset.top);
			}
			,helper: "clone"
			,revert: "invalid"
			,zIndex: 0
			,hoverClass: "droppableHover"
			,opacity: 0.50
			,cursor: "move"
		});
	});	
}
jQuery( document ).ready(function(){
	/*
	* bind to icon nodes
	*/
	navBindNodes();
});
/*
* open/close
*/
var AdminNavPop=false;
/* 
* open nav when created closed 
*/
function OpenAdminNav() {
	SetDisplay('AdminNavHeadOpened','block');
	SetDisplay('AdminNavHeadClosed','none');
	SetDisplay('AdminNavContentOpened','block');
	SetDisplay('AdminNavContentMinWidth','block');
	SetDisplay('AdminNavContentClosed','none');
	cj.ajax.setVisitProperty('','AdminNavOpen','1');
	navBindNodes();
	if(!AdminNavPop){
		cj.ajax.addonCallback("AdminNavigatorGetNode",'',OpenAdminNavCallback);
		AdminNavPop=true;
	}else{
		cj.ajax.addon('AdminNavigatorOpenNode');
	}
}
function OpenAdminNavCallback(serverResponse){
	if (serverResponse != '') {
		var el1 = document.getElementById("AdminNavContentOpened");
		if (el1) {
			el1.innerHTML = serverResponse
			SetDisplay('AdminNavContentOpened','block');
			navBindNodes();
			};
	}
}
/* 
* close nav when created closed 
*/
function reCloseAdminNav() {
	SetDisplay('AdminNavHeadOpened','none');
	SetDisplay('AdminNavHeadClosed','block');
	SetDisplay('AdminNavContentOpened','none');
	SetDisplay('AdminNavContentMinWidth','none');
	SetDisplay('AdminNavContentClosed','block');
	cj.ajax.setVisitProperty('','AdminNavOpen','0')
}
/* 
* open nav when created open
*/
function closeAdminNav() {
	SetDisplay('AdminNavHeadOpened','none');
	SetDisplay('AdminNavContentOpened','none');
	SetDisplay('AdminNavHeadClosed','block');
	SetDisplay('AdminNavContentClosed','block');
	// var allowSaveState = true;
	// var $allowSaveStateInput=$('input[name=allowAdminNavSaveState');
	// console.log('$allowSaveStateInput.length [' + $allowSaveStateInput.length + ']');
	// if($allowSaveStateInput.length && $allowSaveStateInput.val()=='false') {
	// 	allowSaveState=false;
	// }
	// console.log('allowSaveState [' + allowSaveState + ']');
	//if (allowSaveState) { 
		cj.ajax.setVisitProperty('','AdminNavOpen','0'); 
	//}
}
/* 
* close nav when created open
*/
function reOpenAdminNav() {
	SetDisplay('AdminNavHeadOpened','block');
	SetDisplay('AdminNavContentOpened','block');
	SetDisplay('AdminNavHeadClosed','none');
	SetDisplay('AdminNavContentClosed','none');
	navBindNodes();
	cj.ajax.setVisitProperty('','AdminNavOpen','1')
}

