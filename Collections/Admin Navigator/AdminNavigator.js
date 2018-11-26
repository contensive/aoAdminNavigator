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
	console.log("AdminNavOpenClickCallback, contentNode ["+arg.contentNode+"], onEmptyHide ["+arg.onEmptyHide+"], onEmptyShow ["+arg.onEmptyShow+"]");
	if (serverResponse == '') {
		if (document.getElementById(onEmptyHideID)) { document.getElementById(onEmptyHideID).style.display = 'none' }
		if (document.getElementById(onEmptyShowID)) { document.getElementById(onEmptyShowID).style.display = 'block' }
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
*	bind nodes
*/
function navBindNodes() {
	jQuery(".navDrag").each(function(){
		console.log("navBindNodes, node [" + this.id + "]");
		jQuery(this).draggable({
			stop: function(event, ui){
				console.log("navDrag:stop, htmlId ["+this.id+"]");
				navDrop(this.id,ui.position.left,ui.position.top);;
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

/*
* OnReady
*/
jQuery( document ).ready(function(){
	/*
	* bind to icon nodes
	*/
	navBindNodes();
});
