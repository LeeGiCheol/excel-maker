package me.excelmaker;


import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import javax.servlet.http.HttpServletResponse;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.*;

/**
 * VO에 저장된 필드명과 해당 VO의 데이터를 통해
 * 엑셀을 쉽게 만든다.
 *
 * @author LEEGICHEOL
 * @since 2021.03.26
 */
public class ExcelMaker {

    /**
     * 파일, 시트 이름
     */
    private String sheetName = "download";

    /**
     * 확장자명
     */
    private String fileExtension = ".xlsx";

    /**
     * 사용하지 않는 필드명
     */
    private String removeField;

    /**
     * 사용하지 않는 필드 리스트
     */
    private List<String> removeFields = new ArrayList<>();

    /**
     * 엑셀에 표기될 필드 이름
     */
    private final Map<String, String> changeFieldName = new HashMap<>();


    public ExcelMaker setSheetName(String sheetName) {
        this.sheetName = sheetName;
        return this;
    }

    public ExcelMaker setFileExtension(String fileExtension) {
        this.fileExtension = fileExtension;
        return this;
    }

    public ExcelMaker setRemoveField(String removeField) {
        this.removeFields.add(removeField);
        return this;
    }

    private List<String> getRemoveFields() {
        return removeFields;
    }

    public ExcelMaker setRemoveFields(List<String> removeFields) {
        this.removeFields = removeFields;
        return this;
    }

    private String getChangeFieldName(String fieldName) {
        return changeFieldName.get(fieldName);
    }

    public ExcelMaker setChangeFieldName(String oldFieldName, String newFieldName) {
        this.changeFieldName.put(oldFieldName, newFieldName);
        return this;
    }


    /**
     * 엑셀을 만든다.
     *
     * @param response    웹사이트 저장 용도
     * @param voClass     VO 클래스
     * @param dataList    엑셀에 표기할 데이터 리스트
     * @param columnSize  컬럼 가로 길이
     *
     */
    public void makeExcel(HttpServletResponse response, Class<?> voClass, List<?> dataList, int columnSize) {
        SXSSFWorkbook wb = new SXSSFWorkbook();

        try (OutputStream output = response.getOutputStream()) {
            int cellNum = 0;
            int currentRow = 0;
            String colTitle;

            Sheet sh = wb.createSheet(sheetName);
            Row row = sh.createRow(currentRow++);
            Cell cell;

            Field[] allFieldsName = voClass.getDeclaredFields();
            List<Field> useFieldName = getUseField(allFieldsName);

            int fieldSize = useFieldName.size();
            mergeRowRegion(sh, fieldSize, columnSize);

            int fieldNumber = 0;

            for (int i = 0; i < fieldSize * columnSize; i+=columnSize) {
                cell = row.createCell(cellNum + i);
                colTitle = useFieldName.get(fieldNumber++).getName();

                if (getChangeFieldName(colTitle) != null) {
                    colTitle = getChangeFieldName(colTitle);
                }

                cell.setCellValue(colTitle);
            }

            for (Object data : dataList) {
                row = sh.createRow(currentRow++);
                setFieldValues(sh, row, allFieldsName, useFieldName, data, columnSize);
            }

            String fileName = this.sheetName;
            fileName = new String(fileName.getBytes("UTF-8"), "ISO-8859-1");

            response.reset();
            response.setContentType("application/octet-stream");
            response.setHeader("Content-Disposition", "attachment;filename=\"" + fileName + this.fileExtension + "\"");

            wb.write(output);
        } catch(Exception e) {
            e.printStackTrace();
        }
        finally {
            wb.dispose();
        }
    }


    /**
     * 컬럼 가로를 columnSize만큼 병합한다.
     *
     * @param sh         시트
     * @param fieldSize  컬럼 개수
     * @param columnSize 컬럼 가로 길이
     */
    private void mergeRowRegion(Sheet sh, int fieldSize, int columnSize) {
        if (columnSize == 1) {
            return;
        }

        int currentColumn = 0;
        for (int i = 0; i < fieldSize; i++) {
            sh.addMergedRegion(new CellRangeAddress(0, 0, currentColumn, currentColumn + columnSize - 1));
            currentColumn += columnSize;
        }
    }


    /**
     * VO 필드 중 실제 사용되는 필드를 찾는다.
     *
     * @param allFieldsName VO의 전체 필드
     * @return              VO 필드 중 사용되는 필드
     */
    private List<Field> getUseField(Field[] allFieldsName) {
        List<Field> useFieldName = new ArrayList<>();

        if (getRemoveFields().size() == 0) {
            return Arrays.asList(allFieldsName);
        }

        boolean flag;
        for (Field value : allFieldsName) {
            flag = true;
            for (int i = 0; i < getRemoveFields().size(); i++) {
                if (value.getName().equals(getRemoveFields().get(i))) {
                    flag = false;
                    break;
                }
            }

            if (flag) {
                useFieldName.add(value);
            }
        }

        return useFieldName;
    }


    /**
     * 데이터를 삽입한다.
     *
     * @param sh                시트
     * @param row               열
     * @param allFieldsName     VO의 전체 필드
     * @param useFieldName      VO 필드 중 사용되는 필드
     * @param data              입력할 데이터
     * @param columnSize        컬럼 가로 길이
     * @throws IllegalAccessException  Object cellValue = useFieldName.get(i).get(data); -> 참조할 수 없는 필드일 경우 Exception 발생
     */
    private void setFieldValues(Sheet sh, Row row, Field[] allFieldsName, List<Field> useFieldName, Object data, int columnSize) throws IllegalAccessException {
        Cell cell;
        int fieldNumber = 0;
        int cellnum = 0;

        for (int i = 0; i < useFieldName.size(); i++) {
            if (!allFieldsName[fieldNumber++].equals(useFieldName.get(i))) {
                i--;
                continue;
            }

            useFieldName.get(i).setAccessible(true);

            Object cellValue = useFieldName.get(i).get(data);

            mergeRowRegion(sh, cellnum, columnSize);

            cell = row.createCell(cellnum);
            cell.setCellValue(String.valueOf(cellValue));
            cellnum += columnSize;
        }
    }

}