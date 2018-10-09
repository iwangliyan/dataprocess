// 导入依赖库
var express = require('express');
var fs = require('fs'); 
var async = require('async');
var xlsx = require('node-xlsx');
var fse = require('fs-extra');
var klaw = require('klaw');

var files = [];
var rows = [];

console.log("aaaaa")

// 遍历data目录下的所有文件的路径
klaw(__dirname + "/data").on('data', function (item) {
  console.log("item.path: " + item)
  var index = item.path.lastIndexOf(".");
  var length = item.path.length;
  // 获取文件的后缀名字，只处理.xls后缀的文件
  var suffix = item.path.substring(index + 1, length);
  if (suffix == "xls") {
		files.push(item.path);
  }
}).on('end', function () {
  // 遍历结束之后，触发end事件，会执行这里的回调函数
  // 这里的逻辑循环处理每个文件
  for (var i = 0; i < files.length; i++) {
		console.log("==> 开始处理文件：" + files[i]);
		// 开始解析文件
    parseFile(files[i]);
  };
  generateExcel(rows);
});

function parseFile(file) {
  // 根据路径读取excel文件为对象
  var workSheetsFromFile = xlsx.parse(file);
  // 从读到的对象中提取出文件内容
  var rows_data = workSheetsFromFile[0];
  var data = rows_data.data;
  var len = data.length;
  var map = {};
  
  // 遍历excel内容的每一行，取出来的数据接口是一个数组
  // 其中第一个数据代表时间(对应excel的第一列)，以此类推
  for (var i = 0; i < len; i++) {
		var item_arr = data[i];
		var key = item_arr[0];

    // 过滤掉每个xls文件标题列，只处理事件
    // 根据第一项是否包含数字来判断
    var p = /[0-9]/
    if (!p.test(key)) {
      continue;
    }
    // 使用时间作为key，每个时间生成一条记录
		if (!map[key]) {
		  var record = new Record(key);
		  record.add(item_arr[2], item_arr[3], item_arr[4]);
		  map[key] = record;
		} else {
      // 如果时间重复，直接把数据加到原来的记录上面
		  var origin_record = map[key];
		  origin_record.add(item_arr[2], item_arr[3], item_arr[4]);
		}
  }
  // console.log("maps: ", map);
  for (id in map) {
		rows.push(generateRow(map[id]));
  }
}
function generateExcel(rows) {
  var head = ["时间", "空气温度(1-1)","空气温度(1-2)", "平均空气温度","空气湿度(2-1)", "空气湿度(2-2)",
	  "平均空气湿度", "照度(3-1)", "照度(3-2)", "平均照度", "土壤温度(4-1)",
	  "土壤温度(4-2)", "土壤温度(4-3)", "土壤温度(4-4)", "平均土壤温度", 
	  "土壤湿度(5-1)", "土壤湿度(5-2)", "平均土壤湿度",
	  "二氧化碳(20-1)", "二氧化碳(20-2)", "平均二氧化碳"];
  rows.unshift(head);

  var buffer = xlsx.build([{name: "数据整合处理", data: rows}]);
  fs.writeFileSync('西红柿种植环境数据.xlsx', buffer, 'binary');
}
// 根据Record生成excel文件里面对应的一列数据
// 最早push上去的是第一行，以此类推
function generateRow(record) {
  var row = [];
  row.push(record.id);

  row.push(record.air_temperature[0]);
  row.push(record.air_temperature[1]);
  row.push(getEven(record.air_temperature));

  row.push(record.air_humidity[0]);
  row.push(record.air_humidity[1]);
  row.push(getEven(record.air_humidity));

  row.push(record.light[0]);
  row.push(record.light[1]);
  row.push(getEven(record.light));

  row.push(record.soil_temperature[0]);
  row.push(record.soil_temperature[1]);
  row.push(record.soil_temperature[2]);
  row.push(record.soil_temperature[3]);
  row.push(getEven(record.soil_temperature));

  row.push(record.soil_humidity[0]);
  row.push(record.soil_humidity[1]);
  row.push(getEven(record.soil_humidity));

  row.push(record.CO2[0]);
  row.push(record.CO2[1]);
  row.push(getEven(record.CO2));

  return row;
}

// 根据数组里面的值计算平均值
function getEven(arr) {
  if (!arr || arr.length == 0) {
		return 0;
  }
  var total = 0;
  for (var i = 0; i < arr.length; i++) {
		total += parseFloat(arr[i]);
  }
  return Math.round(total * 1000 / arr.length) / 1000;
}
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
Record.prototype.add = function(sensor, channel, value) {
  // 不是我的传感器数据，不处理
  if ([1,2,3,4].indexOf(parseInt(sensor)) == -1) {
		return;
  }
  var keyConfig = this.getKey(channel);
  if (keyConfig) {
		var prop_value_arr = this[keyConfig.name];
		if (prop_value_arr.length > keyConfig.num) {
		  console.warn("==> 该类型数据个数超过预定值：" + JSON.stringify(keyConfig));
		} else {
		  prop_value_arr.push(value);
		}
  }
};
Record.prototype.getKey = function(tag) {
  // 为了便于理解牺牲性能
  for (key in this.key_config) {
		if (this.key_config[key]["tag"] == parseInt(tag)) {
		  return this.key_config[key];
		}
  }
  console.warn("==> 检测到不存在的通道，请确认数据格式是否变化，通道：" + tag);
};
Record.prototype.key_config = {
  air_temperature: {
		name: "air_temperature",
		tag: 1,
		num: 2
  },
  air_humidity: {
		name: "air_humidity",
		tag: 2,
		num: 2 
  },
  light: {
		name: "light",
		tag: 3,
		num: 2
  },
  soil_temperature: {
		name: "soil_temperature",
		tag: 4,
		num: 4
  },
  soil_humidity: {
		name: "soil_humidity",
		tag: 5,
		num: 2
  },
  CO2: {
		name: "CO2",
		tag: 20,
		num: 2
  }
};
