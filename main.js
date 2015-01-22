/*
*  for the array input and output setArray and getArray needed, which are defined in the JS.bas class 
*/

function main(arr){
	return setArray(_.shuffle(getArray(arr)));
}

function main1(){
	return setArray(_.shuffle([1,2,3,4,5]));
}

function mystr(){
	return {name : "hi", age : "1"};
}

function myzipp(arr1, arr2){
	return setArray(_.zip(getArray(arr1),getArray(arr2)));
}

function myrange(n){
	return setArray(_.range(n));
}
