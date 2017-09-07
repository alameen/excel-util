package com.aim.airhacks;

import java.io.File;
import java.io.IOException;
import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableCell;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

public class Manipulator {
  private static File sourceFile;
  private static File destinationFile;
  private static final int SOURCE_EMAIL_COL = 2;
  private static final int DESTINATION_EMAIL_COL = 13;
  private static final int SOURCE_ADM_COL  = 5;
  private static final int SOURCE_UNI_COL  = 6;
  private static final int SOURCE_COURSE_COL  = 7;
  private static final int DEST_ADM_COL  = 17;
  private static final int DEST_UNI_COL  = 18;
  private static final int DEST_COURSE_COL  = 19;
  private static Sheet sourceSheet;
  private static WritableSheet writableSheet;
  private Workbook sourceBook;
  private Workbook destinationBook;
  private WritableWorkbook writableWorkbook;

  private void run() throws IOException, BiffException, WriteException {
    System.out.println("Begin..");
    this.sourceFile = new File("/home/aim/Dropbox/Saves/PTDF_Email.xls");
    this.destinationFile = new File("/home/aim/Dropbox/Saves/msc_credentials.xls");

    assert sourceFile.exists();
    assert destinationFile.exists();

    sourceBook = Workbook.getWorkbook(sourceFile);
    sourceSheet = sourceBook.getSheet(0);

    destinationBook = Workbook.getWorkbook(destinationFile);
    writableWorkbook = Workbook.createWorkbook(destinationFile, destinationBook);
    writableSheet = writableWorkbook.getSheet(0);

    Cell[] writableColumns = writableSheet.getColumn(DESTINATION_EMAIL_COL);
    Cell[] sourceEmailCol = sourceSheet.getColumn(SOURCE_EMAIL_COL);

    for (Cell c : sourceEmailCol) {
      if (c.getContents().isEmpty() || getAdmissionStatus(c.getRow()).isEmpty()) continue;
      writeInformation(getRowInfo(c.getContents(), writableColumns));
    }

    writableWorkbook.write();
    writableWorkbook.close();
    destinationBook.close();
  }

  private int getRowInfo(String contents, Cell[] emailCells) {
    for (Cell emailCell : emailCells) {
      if (contents.equals(emailCell.getContents())) {
        System.out.println("Email: " + contents + " Row: " + emailCell.getRow());
        return emailCell.getRow();
      }
    }

    return 0;
  }

  private void writeInformation(int row) throws IOException, WriteException {
    WritableCell admNoCell = writableSheet.getWritableCell(DEST_ADM_COL, row);
    System.out.println("Cell.. " + admNoCell.getContents());

    if (admNoCell.getContents().isEmpty()) return;
    WritableCell uniCell = writableSheet.getWritableCell(DEST_UNI_COL, row);
    WritableCell courseCell = writableSheet.getWritableCell(DEST_COURSE_COL, row);

    Label admValue = new Label(DEST_ADM_COL, row, getAdmissionStatus(row));
    Label uniValue = new Label(DEST_UNI_COL, row, getUniName(row));
    Label courseValue = new Label(DEST_COURSE_COL, row, getCourse(row));

    admValue.setString(getAdmissionStatus(row));
    uniValue.setString(getUniName(row));
    courseValue.setString(getCourse(row));

    writableSheet.addCell(admNoCell);
    writableSheet.addCell(uniCell);
    writableSheet.addCell(courseCell);

    System.out.println("The admission is: " + writableSheet.getWritableCell(DEST_ADM_COL, row));
  }

  private String getAdmissionStatus(int row) {
    return sourceSheet.getCell(SOURCE_ADM_COL, row).getContents();
  }

  private String getUniName(int row) {
    return sourceSheet.getCell(SOURCE_UNI_COL, row).getContents();
  }

  private String getCourse(int row) {
    return sourceSheet.getCell(SOURCE_COURSE_COL, row).getContents();
  }

  public static void main(String[] args) throws IOException, BiffException, WriteException {
    Manipulator manipulator = new Manipulator();
    manipulator.run();
    System.out.println("Done!");
  }
}
