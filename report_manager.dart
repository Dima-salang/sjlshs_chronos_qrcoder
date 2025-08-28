import 'dart:io';

import 'package:cloud_firestore/cloud_firestore.dart';
import 'package:isar/isar.dart';
import 'package:logger/logger.dart';
import 'package:excel/excel.dart';




class ReportManager {
  final Isar isar;
  final FirebaseFirestore firestore;
  final Logger logger;

  ReportManager({required this.isar, required this.firestore, required this.logger});



  // get all student records from firestore for report generation
  Future<QuerySnapshot<Map<String, dynamic>>> getStudentRecords(DateTime start, DateTime end, String? section) async {
    try {
      var query = firestore.collection('attendance').where('timestamp', isGreaterThanOrEqualTo: start, isLessThanOrEqualTo: end);

      // Only add the section filter if it's not null
      if (section != null) {
        query = query.where('studentSection', isEqualTo: section);
      }

      return await query.get();
    } catch (e) {
      logger.e('Error getting student records: $e');
      throw Exception('Error getting student records: $e');
    }
  }




  // make the excel file
  Future<void> writeReport(DateTime start, DateTime end, String filePath, String? section) async {
    try {
      final records = await getStudentRecords(start, end, section);
      final excel = Excel.createExcel();
      final months = getMonthsBetween(start, end);
      
      final sheets = <Sheet>[];
      for (final month in months.keys) {
        sheets.add(excel[month]);
      }

      final filledExcel = await fillExcelFile(excel, records, sheets, months);

      // write excel file
      try {
        final bytes = filledExcel.encode();
        if (bytes != null) {
          final file = File(filePath);
          await file.writeAsBytes(bytes);
        } else {
          logger.e('Failed to encode excel file');
          throw Exception('Failed to encode excel file');
        }
      } catch (e) {
        logger.e('Error writing excel file: $e');
        throw Exception('Error writing excel file: $e');
      }
      
    } catch (e) {
      logger.e('Error writing report: $e');
      throw Exception('Error writing report: $e');
    }
  }


  // get the days where the student was present based on their absences from firestore




  // fill excel file with student records
Future<Excel> fillExcelFile(
  Excel excel,
  QuerySnapshot<Map<String, dynamic>> records,
  List<Sheet> sheets,
  Map<String, DateTime> months,
) async {
  try {
    // Group by student
    final studentRecords = <String, List<Map<String, dynamic>>>{};

    for (final doc in records.docs) {
      final data = doc.data();
      final lrn = data['lrn'];
      studentRecords.putIfAbsent(lrn, () => []).add(data);
    }

    for (final sheet in sheets) {
      final month = months[sheet.sheetName];
      if (month == null) continue;
      final days = getDaysForMonth(month);

      // Write headers
      final headers = ['LRN', 'Last Name', 'First Name', 'Year', 'Section'] + days.map((d) => d.day.toString()).toList() + ['Total Absences', 'Total Present'];
      for (int col = 0; col < headers.length; col++) {
        sheet.cell(CellIndex.indexByColumnRow(columnIndex: col, rowIndex: 0)).value = TextCellValue(headers[col]);
      }

      // Write rows
      int row = 1;
      for (final entry in studentRecords.entries) {
        final firstRecord = entry.value.first;
        final dayToAbsence = <int, bool>{};

        // Fill absence days
        for (final record in entry.value) {
          final timestamp = (record['timestamp'] as Timestamp).toDate();
          if (timestamp.year == month.year && timestamp.month == month.month) {
            dayToAbsence[timestamp.day] = record['isAbsent'] ?? false;
          }
        }

        final cells = [
          firstRecord['lrn'],
          firstRecord['lastName'],
          firstRecord['firstName'],
          firstRecord['studentYear'],
          firstRecord['studentSection'],
          ...days.map((d) => dayToAbsence[d.day] == true ? 'A' : ''),
          // summary total of absences per student
          dayToAbsence.values.where((absence) => absence).length,
          // total present days
          days.length - dayToAbsence.values.where((absence) => absence).length,
        ];

        for (int col = 0; col < cells.length; col++) {
          sheet.cell(CellIndex.indexByColumnRow(columnIndex: col, rowIndex: row)).value = TextCellValue(cells[col].toString());
        }

        row++;
      }
    }

    return excel;
  } catch (e) {
    logger.e('Error filling excel file: $e');
    throw Exception('Error filling excel file: $e');
  }
}


  // get all days for the specific month
List<DateTime> getDaysForMonth(DateTime month) {
  final nextMonth = DateTime(month.year, month.month + 1, 1);
  final lastDay = nextMonth.subtract(Duration(days: 1)).day;
  return List<DateTime>.generate(lastDay, (i) => DateTime(month.year, month.month, i + 1));
}



  // get months between start and end
Map<String, DateTime> getMonthsBetween(DateTime start, DateTime end) {
  final months = <String, DateTime>{};
  var current = DateTime(start.year, start.month);

  while (current.isBefore(end) || current.isAtSameMomentAs(end)) {
    months["${current.year}-${current.month.toString().padLeft(2, '0')}"] = current;
    current = DateTime(current.year, current.month + 1);
  }

  return months;
}



}
