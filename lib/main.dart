import 'dart:io';

import 'package:flutter/material.dart';
import 'package:path_provider/path_provider.dart';
import 'package:open_file/open_file.dart';
import 'package:syncfusion_flutter_xlsio/xlsio.dart' as xlsio;

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
          ],
        ),
      ),
    );
  }

  Future<void> createExcel() async {
    /// creates an Excel Workbook with 1 worksheet
    final xlsio.Workbook workbook = xlsio.Workbook();

    /// access the worksheet of the workbook to store data in
    /// 0th index stores first worksheet
    final xlsio.Worksheet sheet = workbook.worksheets[0];
    sheet.getRangeByName('A1').setText("Hello World");
    sheet
        .getRangeByIndex(int.parse(posn.text[0]), int.parse(posn.text[2]))
        .setText("${data.text}");

    /// storing the bytes of the workbook
    final List<int> bytesOfExcelWorkbook = workbook.saveAsStream();

    workbook.dispose();

    /// stores file path to save at
    final String savePath = (await getApplicationSupportDirectory()).path;
    final String docDir = (await getApplicationDocumentsDirectory()).path;

    print(
        savePath); // C:\Users\...\AppData\Roaming\com.example\excel_in_flutter
    print(docDir); // C:\Users\...\OneDrive\Documents

    /// stores file path + file name
    final String filename = '$savePath/output.xlsx';
    final String docFileName = '$docDir/output.xlsx';

    /// create the file here and write the bytes to it
    File file = File(filename);
    await file.writeAsBytes(bytesOfExcelWorkbook, flush: true);

    file = File(docFileName);
    await file.writeAsBytes(bytesOfExcelWorkbook, flush: true);

    OpenFile.open(filename);
  }
}
