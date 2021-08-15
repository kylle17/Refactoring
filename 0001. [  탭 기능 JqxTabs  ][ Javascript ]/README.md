# [  탭 기능 JqxTabs  ][ Javascript ] <br><br>

- 설명 : <br>

     JqxTabs를 사용해서 구현한 탭 기능을 리팩터링 해서 가독성을 향상시켰다. <br>
<br><br><br><br>





### 1. 탭 기능의 전체 흐름 . <br>

탭 기능의 경우 현재 열린 탭이 있으면 해당 탭으로 이동하고 . <br>
열린 탭이 없으면 탭을 생성한다. <br>


```javascript
function actTab( newTabId , newTabName, newTabUrl ){
	if( findMatchTabIndex(newTabName) > -1 ){
		goTab( metchTabIndex );
	} else{
		createTab( newTabId , newTabUrl , newTabName  );
	};
};
```

- findMatchTabIndex() 함수에서 일치하는 탭이 있는지 검색한다. <br>
- 일치하는 탭이 있으면 goTab() 함수에서 해당 탭으로 이동한다. <br>
- 없으면 createTab() 함수에서 탭을 생성한다. <br>
<br><br><br><br>





### 2. 일치하는 탭 검색하기 . <br>


- 현재 열려 있는 탭들의 이름과 요청한 탭의 이름을 비교한다. <br>

```javascript
function findMatchTabIndex( newTabName ){
	let result ;
	for(let i=0 ; i < currentTabLength() ; i++){
		if( newTabName == currentTabName(i) ) return result = i;
	};
	return result ;
};
```

- 현재 열려 있는 탭들의 갯수 구하기. <br>

```javascript
function currentTabLength(){
	return $('#jqxTabs').jqxTabs('length');
};
```

- 현재 열려 있는 탭들의 이름 가져오기. <br>

```javascript
function currentTabName( cuttentTabIndex ){
	return $('#jqxTabs').jqxTabs('getTitleAt', cuttentTabIndex )
};
```
<br>

<br><br><br><br>





### 3. 열린 탭으로 이동하기. <br>

jqxTabs 라이브러리를 사용해서 해당 탭으로 이동한다. <br>

```javascript
function goTab( selectNo ){
	$('#jqxTabs').jqxTabs('val', selectNo);
};
```
<br><br><br><br>





### 4. 탭 신규로 생성하기. <br>

jqxTabs 라이브러리를 사용해서 탭을 생성한다. <br>

```javascript
function createTab( newTabId , newTabUrl , newTabName  ){
	let content = '<iframe id="'+newTabId+'" frameborder="0"  src="'+newTabUrl+'" width="100%" height="100%" marginwidth="0" marginheight="0" ></iframe>';
	$('#jqxTabs').jqxTabs('addLast', newTabName , content);
	$('#jqxTabs').jqxTabs('ensureVisible', -1);
};
```
<br><br><br><br>






### 마무리 <br>

기존 코드의 경우 한로직에 작성되어 있어서 큰흐름 파악이 어려웠는데 <br>
위와 같이 로직을 분리한 결과 큰흐름을 파악하기 쉬워졌다. <br>
























<br><br><br><br>




```java
 
```









