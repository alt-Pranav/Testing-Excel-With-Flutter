import 'dart:io';

import 'package:excel/excel.dart';
import 'package:flutter/material.dart';
import 'package:path_provider/path_provider.dart';
import 'package:open_file/open_file.dart';

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
        // This is the theme of your application.
        //
        // Try running your application with "flutter run". You'll see the
        // application has a blue toolbar. Then, without quitting the app, try
        // changing the primarySwatch below to Colors.green and then invoke
        // "hot reload" (press "r" in the console where you ran "flutter run",
        // or simply save your changes to "hot reload" in a Flutter IDE).
        // Notice that the counter didn't reset back to zero; the application
        // is not restarted.
        primarySwatch: Colors.blue,
      ),
      home: const MyHomePage(title: 'Flutter Demo Home Page'),
    );
  }
}

class MyHomePage extends StatefulWidget {
  const MyHomePage({super.key, required this.title});

  // This widget is the home page of your application. It is stateful, meaning
  // that it has a State object (defined below) that contains fields that affect
  // how it looks.

  // This class is the configuration for the state. It holds the values (in this
  // case the title) provided by the parent (in this case the App widget) and
  // used by the build method of the State. Fields in a Widget subclass are
  // always marked "final".

  final String title;

  @override
  State<MyHomePage> createState() => _MyHomePageState();
}

class _MyHomePageState extends State<MyHomePage> {
  TextEditingController posn = TextEditingController();
  TextEditingController data = TextEditingController();

  /// stores file path to save at
  late String savePath;

  /// stores file path + file name
  late String filename;

  @override
  void initState() {
    // TODO: implement initState
    super.initState();
    setPaths();
  }

  Future<void> setPaths() async {
    savePath = (await getApplicationDocumentsDirectory()).path;
    filename = '$savePath/output.xlsx';
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: Text("Excel with Flutter"),
      ),
      body: Center(
        child: Column(
          children: [
            Row(
              mainAxisAlignment: MainAxisAlignment.spaceEvenly,
              children: [
                SizedBox(
                  height: 30,
                  width: 30,
                  child: TextField(
                    controller: posn,
                  ),
                ),
                SizedBox(
                  height: 30,
                  width: 30,
                  child: TextField(
                    controller: data,
                  ),
                )
              ],
            ),
            ElevatedButton(
              child: Text("Create Excel file"),
              onPressed: createExcel,
            ),
            ElevatedButton(
                onPressed: appendExcel, child: Text("Append to Excel file")),
          ],
        ),
      ),
    );
  }

  void createExcel() {
    var excel = Excel.createExcel();

    Sheet sheet1 = excel["Sheet_for_data"];

    var cell =
        sheet1.cell(CellIndex.indexByColumnRow(columnIndex: 4, rowIndex: 1));

    cell.value = 24;

    excel.encode().then((value) {
      File(filename)
        ..createSync(recursive: true)
        ..writeAsBytes(value);
    });
  }

  void appendExcel() {
    var bytes = File(filename).readAsBytesSync();
    var excel = Excel.decodeBytes(bytes);

    Sheet sheet1 = excel["Sheet_for_data"];

    var cell =
        sheet1.cell(CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: 1));

    cell.value = "Appended!";

    excel.encode().then((value) {
      File(filename)
        ..createSync(recursive: true)
        ..writeAsBytes(value);
    });

    OpenFile.open(filename);
  }
}
