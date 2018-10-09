// 导入依赖库
var express = require('express');
var fs = require('fs'); 
var async = require('async');
var xlsx = require('node-xlsx');
var fse = require('fs-extra');
var klaw = require('klaw');

var files = [];
var rows = [];

var file_path = __dirname + "/西红柿种植环境数据.xlsx";
parseFile(file_path);
generateExcel(rows);

// 保存每行记录的数据结构
function Record(id) {
  this.id = id;
  // 空气温度 1
  this.air_temperature = [];
  // 空气湿度 2
  this.air_humidity = [];
  // 照度 3
  this.light = [];
  // 土壤温度 4
  this.soil_temperature = [];
  // 土壤湿度 5
  this.soil_humidity = [];
  // CO2 20
  this.CO2 = [];
}
function addOne(record, arr) {
  record.air_temperature.push(arr[3]);
  record.air_humidity.push(arr[6]);
  record.light.push(arr[9]);
  record.soil_temperature.push(arr[14]);
  record.soil_humidity.push(arr[17]);
  record.CO2.push(arr[20]);
};

function parseFile(file) {
  // 根据路径读取excel文件为对象
  console.log("start file ...." + file)
  var workSheetsFromFile = xlsx.parse(file);
  // 从读到的对象中提取出文件内容
  var rows_data = workSheetsFromFile[0];
  var data = rows_data.data;
  var len = data.length;
  var map = {};

  console.log("len: " + len);
  
  // 遍历excel内容的每一行，取出来的数据接口是一个数组
  // 其中第一个数据代表时间(对应excel的第一列)，以此类推
  for (var i = 0; i < len; i++) {
		var item_arr = data[i];
		var timeItem = item_arr[0];

    // 过滤掉每个xls文件标题列，只处理事件
    // 根据第一项是否包含数字来判断
    var p = /[0-9]/
    if (!p.test(timeItem)) {
      continue;
    }
    var time_arr = timeItem.split(" ");
    var key = time_arr[0];

    // 使用时间作为key，每个时间生成一条记录
		if (!map[key]) {
		  var record = new Record(key);
		  addOne(record, item_arr);
		  map[key] = record;
		} else {
      // 如果时间重复，直接把数据加到原来的记录上面
		  var origin_record = map[key];
		  addOne(origin_record, item_arr);
		}
  }
  // console.log("maps: ", map);
  for (id in map) {
		rows.push(generateRow(map[id]));
  }
}
function generateExcel(rows) {
  var head = ["时间", "空气温度(max)","空气温度(min)", "平均空气温度", "活动空气温度", "空气湿度(max)", "空气湿度(min)",
	  "平均空气湿度", "照度(max)", "照度(min)", "平均照度", "土壤温度(max)",
	  "土壤温度(min)",  "平均土壤温度", "活动土壤温度", 
	  "土壤湿度(max)", "土壤湿度(min)", "平均土壤湿度",
	  "二氧化碳(max)", "二氧化碳(min)", "平均二氧化碳"];
  rows.unshift(head);

  var buffer = xlsx.build([{name: "数据整合处理", data: rows}]);
  fs.writeFileSync('日期统计数据.xlsx', buffer, 'binary');
}
// 根据Record生成excel文件里面对应的一列数据
// 最早push上去的是第一行，以此类推
function generateRow(record) {
  var row = [];
  row.push(record.id);

  row.push(getMax(record.air_temperature));
  row.push(getMin(record.air_temperature));
  row.push(getEven(record.air_temperature));
  row.push(getUpEven(record.air_temperature));

  row.push(getMax(record.air_humidity));
  row.push(getMin(record.air_humidity));
  row.push(getEven(record.air_humidity));

  row.push(getMax(record.light));
  row.push(0);
  row.push(getEven(record.light));

  row.push(getMax(record.soil_temperature));
  row.push(getMin(record.soil_temperature));
  row.push(getEven(record.soil_temperature));
  row.push(getUpEven(record.soil_temperature));

  row.push(getMax(record.soil_humidity));
  row.push(getMin(record.soil_humidity));
  row.push(getEven(record.soil_humidity));

  row.push(getMax(record.CO2));
  row.push(getMin(record.CO2));
  row.push(getEven(record.CO2));

  return row;
}

// 根据数组里面的值计算平均值
function getEven(arr) {
  if (!arr || arr.length == 0) {
		return 0;
  }
  var arr = fliterInvalidValue(arr);
  var total = 0;
  for (var i = 0; i < arr.length; i++) {
		total += parseFloat(arr[i]);
  }
  return Math.round(total * 1000 / arr.length) / 1000;
}
function getMax(arr) {
	if (!arr || arr.length == 0) {
		return 0;
  }
  var arr = fliterInvalidValue(arr);
  var maxValue = 0;
  for (var i = 0; i < arr.length; i++) {
  	if (arr[i] > maxValue) {
  		maxValue = arr[i];
  	}
  }
  return maxValue;
}
function getMin(arr) {
	if (!arr || arr.length == 0) {
		return 0;
  }
  var arr = fliterInvalidValue(arr);
  var minValue = arr[0];
  for (var i = 0; i < arr.length; i++) {
  	if (arr[i] < minValue) {
  		minValue = arr[i];
  	}
  }
  return minValue;
}
function getUpEven(arr) {
  if (!arr || arr.length == 0) {
    return 0;
  }

  var even = getEven(arr);
  var len = 0;
  var total = 0;

  for (var i = 0; i < arr.length; i++) {
    if (arr[i] > even) {
      total += parseFloat(arr[i]);
      len += 1
    }
  }
  return Math.round(total * 1000 / len) / 1000;
}
function fliterInvalidValue(arr) {
	if (!arr || arr.length == 0) {
		return [];
  }
	var newArr = [];
	for (var i = 0; i < arr.length; i++) {
		if (arr[i] != 0) {
			newArr.push(arr[i])
		}
	};
	return newArr
}

