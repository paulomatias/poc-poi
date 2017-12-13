package com.paulo;

public class CsvFile {
    private int lineNumber;
    private int columnNumber;
    private String[][] array;

    public CsvFile(int lineNumber, int columnNumber) {
        this.lineNumber = lineNumber;
        this.columnNumber = columnNumber;
        this.array = new String[lineNumber][columnNumber];

        for (int line = 0; line < lineNumber; line++) {
            for (int column = 0; column < columnNumber; column++) {
                char x = (char) ('a' + line);
                array[line][column] = "cell_value_" + x + column;
            }
        }
    }

    public int getLineNumber() {
        return lineNumber;
    }

    public int getColumnNumber() {
        return columnNumber;
    }

    public String[][] getArray() {
        return array;
    }

}
