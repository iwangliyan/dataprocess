// 导入依赖库
var express = require('express');
var fs = require('fs'); 
var async = require('async');
var xlsx = require('node-xlsx');
var fse = require('fs-extra');
var klaw = require('klaw');

Date.prototype.Format = function (fmt) {
	var o = {
	    "M+": this.getMonth() + 1, 
	    "d+": this.getDate(), //日 
	    "h+": this.getHours(), //小时 
	    "m+": this.getMinutes(), //分 
	    "s+": this.getSeconds(), //秒 
	    "q+": Math.floor((this.getMonth() + 3) / 3), //季度 
	    "S": this.getMilliseconds() //毫秒 
	};
	if (/(y+)/.test(fmt)) fmt = fmt.replace(RegExp.$1, (this.getFullYear() + "").substr(4 - RegExp.$1.length));
	for (var k in o)
	if (new RegExp("(" + k + ")").test(fmt)) fmt = fmt.replace(RegExp.$1, (RegExp.$1.length == 1) ? (o[k]) : (("00" + o[k]).substr(("" + o[k]).length)));
	return fmt;
}

var rows = [];
var head = ["time_stamp", "day", "hour", "locId"];

rows.unshift(head);
var startDate = new Date(2017, 00, 01);
var monthCount = 12;
var count = 0;

for (var i = 11; i < monthCount; i++) {
	var date = new Date(2017, i + 1, 0);
	var dateCount = date.getDate();

	for (var d = 1; d <= dateCount; d++) {
		var currentDate = new Date(2017, i, d);
		console.log("generate: " + currentDate.Format("yyyy-MM-dd"))
		generateDateLocId(currentDate.Format("yyyy-MM-dd"))
	}
}

function generateDateLocId(date) {
	for (var i = 0; i <= 23; i++) {
		for (var j = 1; j <= 33; j++ ) {
			var i = i + "";
			var hour = i.length == 1 ? "0" + i : i;
			count++;
			var arr = [];
			arr.push(date + " " + hour)
			arr.push(date);
			arr.push(hour)
			arr.push(j)

			rows.push(arr);
		}
	}
}

var buffer = xlsx.build([{name: "auto_date", data: rows}]);
fs.writeFileSync('auto_date_12.xlsx', buffer, 'binary');


