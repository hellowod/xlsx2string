# xlsx2string

该工具提供简单高效的命令行和图形界面的方式导出xlsx格式为文本数据。

## 窗口模式

![form](/res/form.png)

## 命令模式

```
xlsx2string 1.0.0.0
Copyright ? abaojin 2017

  -e, --excel       Required. 输入的excel文件路径.

  -j, --json        指定输出的json文件路径.

  -t, --txt         指定输出的txt文件路径.

  -c, --csv         指定输出的csv文件路径.

  -l, --lua         指定输出的lua数据定义代码文件路径.

  -s, --sql         指定输出的sql文件路径.

  -p, --csharp      指定输出的c#数据定义代码文件路径.

  -j, --java        指定输出的java数据定义代码文件路径.

  -+, --cpp         指定输出的cpp数据定义代码文件路径.

  -g, --go          指定输出的go数据定义代码文件路径.

  -h, --header      Required. 表格中有几行是表头.

  -c, --encoding    (Default: utf8-nobom) 指定编码的名称.

  -l, --lowcase     (Default: False) 字段名称自动转换为小写
```

## 插件扩展

工具提供的是一种通用的导出方案，每一种导出类型都提供单独的导出
器如果这些导出器（exporter），你可以非常方便的修改这些导出器的代码，来满足
自身产品的需求，也可以增加新的导出类型和导出器。


任何问题或者错误欢迎给我提问题，我会尽力去解决。
