import 'dart:io';

import 'package:desktop_drop/desktop_drop.dart';
import 'package:file_picker/file_picker.dart';
import 'package:flutter/material.dart';
import 'package:cross_file/cross_file.dart';
import 'package:intl/intl.dart';
import 'PathLabel.dart';
import 'dart:io';
import 'package:excel/excel.dart';
import 'package:path/path.dart' as path;

void main() {
  runApp(const MyApp());
}

class MyApp extends StatelessWidget {
  const MyApp({super.key});

  // This widget is the root of your application.
  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'Flutter Demo',
      theme: ThemeData(
        colorScheme: ColorScheme.fromSeed(seedColor: Colors.deepPurple),
        useMaterial3: true,
      ),
      home: const MyHomePage(title: 'Flutter Demo Home Page'),
    );
  }
}

class MyHomePage extends StatefulWidget {
  const MyHomePage({super.key, required this.title});

  final String title;

  @override
  State<MyHomePage> createState() => _MyHomePageState();
}

class _MyHomePageState extends State<MyHomePage> {
  String _xlsxFilePath = "";
  List<InlineSpan> spans = [];
  @override
  Widget build(BuildContext context) {
    return Scaffold(
        body: DropTarget(
      onDragDone: (detail) {
        _onDragDone(detail);
      },
      child: Container(
        padding: const EdgeInsets.fromLTRB(10, 10, 10, 10),
        child: buildColumn(),
      ),
    ));
  }

  /// 拖拽完毕
  void _onDragDone(DropDoneDetails detail) async {
    if (detail.files.length != 1) {
      return;
    }
    XFile aFile = detail.files[0];
    FileSystemEntityType type = FileSystemEntity.typeSync(aFile.path);
    if (type == FileSystemEntityType.file && aFile.name.endsWith(".xlsx")) {
      setState(() {
        _xlsxFilePath = aFile.path;
        spans = List.from(spans)
          ..add(const TextSpan(
              text: '\n 拖入文件', style: TextStyle(color: Colors.white)));
      });
    }
  }

  Column buildColumn() {
    return Column(
      mainAxisAlignment: MainAxisAlignment.start,
      children: [
        const SizedBox(height: 16),
        PathLabel(text: "已选择文件：$_xlsxFilePath"),
        const SizedBox(height: 16),
        Expanded(
          child: SingleChildScrollView(
            child: Container(
              color: Colors.black,
              width: double.infinity,
              child: RichText(
                text: TextSpan(children: spans),
              ),
            ),
          ),
        ),
        const SizedBox(height: 16),
        Row(
          mainAxisAlignment: MainAxisAlignment.end,
          children: [
            ElevatedButton(
              onPressed: _startWork,
              child: const Text(' 开始处理表格'),
            ),
            const SizedBox(width: 30),
            ElevatedButton(
              onPressed: _pickXlsxFile,
              child: const Text(' 选择或拽入单个xlsx文件'),
            ),
          ],
        ),
        const SizedBox(height: 16),
      ],
    );
  }

  _startWork() {
    setState(() {
      spans = List.from(spans)
        ..add( TextSpan( text: "处理中...", style: TextStyle(color: Colors.yellow)));
    });
    var bytes = File(_xlsxFilePath).readAsBytesSync();
    var excel = Excel.decodeBytes(bytes);

    DateFormat format = DateFormat("yyyy-MM-dd HH:mm:ss");
    bool isAddTime = false;

    for (var table in excel.tables.keys) {
      var rows = excel.tables[table]!.rows;

      for (int i = 3; i < rows.length; i++) {
        var row = rows[i];
        // 需求是将F列的数据生成到G列，日期加1天。并且22:30之前的再加1个半小时。以后得减1个半小时
        // F列就是5 2023-06-26 09:53:18
        if(row[5] == null){
          continue;
        }

        var readDateStr = row[5]!.value.toString();


        DateTime readDate = format.parse(readDateStr.trim());
        // 如果时间在22:30之前，则加1个半小时，否则减1个半小时
        DateTime writeDate;
        if (readDate.hour < 22 || (readDate.hour == 22 && readDate.minute < 30)) {
          writeDate = readDate.add(const Duration(days:1, hours: 1, minutes: 30));
          isAddTime = true;
        } else {
          writeDate = readDate.subtract(const Duration(days:1, hours: 1, minutes: 30));
          isAddTime = false;
        }
        String writeDateStr = writeDate.toString().replaceAll(".000", "");

        String actionStr = isAddTime ? "增加":"减少";
        String logStr =  "\n $readDateStr $actionStr 为: $writeDateStr";
        Color logColor = isAddTime ? Colors.red : Colors.blue;
         setState(() {
           spans = List.from(spans)
             ..add( TextSpan( text: logStr, style: TextStyle(color: logColor)));
         });
         // 这里为什么i要加个1才正常，奇怪了
        excel.updateCell(table, CellIndex.indexByString('G${i+1}'), writeDateStr);

      }
      var fileBytes = excel.save();
      File originFile = File(_xlsxFilePath);
      String originFolder = path.dirname(originFile.path);
      String savedFilePath = path.join(originFolder, 'saved.xlsx');
      File(savedFilePath).writeAsBytesSync(fileBytes!);
      setState(() {
        spans = List.from(spans)
          ..add( TextSpan( text: "\n 保存文件至：$savedFilePath", style: const TextStyle(color: Colors.white)));
      });

    }
  }

  /// 选取 文件
  Future<void> _pickXlsxFile() async {
    FilePickerResult? result = await FilePicker.platform.pickFiles(
      type: FileType.custom,
      allowedExtensions: ['xlsx'],
    );

    if (result != null) {
      setState(() {
        _xlsxFilePath = result.files.single.path!;
        spans = List.from(spans)
          ..add(const TextSpan(
              text: '\n 手动选择文件', style: TextStyle(color: Colors.white)));
      });

    }
  }
}
