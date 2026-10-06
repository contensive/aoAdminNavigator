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
		fetch('/AdminNavigatorOpenNode?nodeid=' + NodeID);
		navBindNodes();
	}else{
		e.ok='ok';
		var arg = {contentNode:ContentNode};
		arg.onEmptyHide = onEmptyHide;
		arg.onEmptyShow = onEmptyShow;
		fetch('/AdminNavigatorGetNode?nodeid=' + NodeID).then(r => r.text()).then(text => AdminNavOpenClickCallback(text, arg));
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
		fetch('/AdminNavigatorCloseNode?nodeid=' + NodeID);
}
/*
*	moved to dashboard.js to keep it together. kept this incase dash updated and not adminNav
*/
function navBindNodes() {
	if(typeof dashBindNavNodes !== 'undefined' && dashBindNavNodes) {
		/* dashBindNavNodes installed, navDrag binding already handled by dashboard.js */
		return;
	}
	jQuery(".navDrag").each(function(){
		jQuery(this).draggable({
			stop: function(event, ui){
				console.log("adminNav.navBindNodes draggable:stop");
				if(typeof navDrop === 'function') {
					navDrop(this.id,ui.offset.left,ui.offset.top);
				}
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
	try {
		localStorage.setItem('AdminNavOpen', '1');
	} catch(e) {
		// Ignore localStorage errors (e.g., private browsing mode)
	}
	navBindNodes();
	if(!AdminNavPop){
		fetch('/AdminNavigatorGetNode').then(r => r.text()).then(text => OpenAdminNavCallback(text));
		AdminNavPop=true;
	}else{
		fetch('/AdminNavigatorOpenNode');
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
	try {
		localStorage.setItem('AdminNavOpen', '0');
	} catch(e) {
		// Ignore localStorage errors (e.g., private browsing mode)
	}
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
		try {
			localStorage.setItem('AdminNavOpen', '0');
		} catch(e) {
			// Ignore localStorage errors (e.g., private browsing mode)
		}
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
	try {
		localStorage.setItem('AdminNavOpen', '1');
	} catch(e) {
		// Ignore localStorage errors (e.g., private browsing mode)
	}
}

