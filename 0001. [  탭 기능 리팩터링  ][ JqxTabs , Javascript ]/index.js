

function addTab( newTabId , newTabName, newTabUrl ){
		let metchTabIndex = findMatchTabIndex(newTabName);
		if( metchTabIndex > -1 ){
				goTab( metchTabIndex );
		} else{
				createTab( newTabId , newTabUrl , newTabName  );
		};
};


function findMatchTabIndex( newTabName ){
	let result ;
	for(let i=0 ; i < currentTabLength() ; i++){
		if( newTabName == currentTabName(i) ) return result = i;
	};
	return result ;
};


function currentTabLength(){
	return $('#jqxTabs').jqxTabs('length');
};

function currentTabName( cuttentTabIndex){
	return $('#jqxTabs').jqxTabs('getTitleAt', cuttentTabIndex )
};


function goTab( selectNo ){
	$('#jqxTabs').jqxTabs('val', selectNo);
};


function createTab( newTabId , newTabUrl , newTabName  ){
	let content = '<iframe id="'+newTabId+'" frameborder="0"  src="'+newTabUrl+'" width="100%" height="100%" marginwidth="0" marginheight="0" ></iframe>';
	$('#jqxTabs').jqxTabs('addLast', newTabName , content);
	$('#jqxTabs').jqxTabs('ensureVisible', -1);
};

