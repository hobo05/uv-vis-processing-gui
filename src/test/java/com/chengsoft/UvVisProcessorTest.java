package com.chengsoft;

import org.junit.Ignore;
import org.junit.Test;

import static org.junit.Assert.*;

/**
 * Created by tcheng on 4/2/16.
 */
public class UvVisProcessorTest {

    @Test
    @Ignore
    public void processAndWriteExcel() throws Exception {
//        String inputExcel = "src/test/resources/sample input.xls";
        String inputExcel = "src/test/resources/input file fluorescence 384 well plate.xls";
        String outputExcel = "src/test/resources/java-output.xlsx";
        UvVisProcessor.processAndWriteExcel(inputExcel, outputExcel);
    }
}