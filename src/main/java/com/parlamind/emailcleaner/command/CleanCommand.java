package com.parlamind.emailcleaner.command;

import edu.stanford.nlp.ling.TaggedWord;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.springframework.shell.standard.ShellComponent;
import org.springframework.shell.standard.ShellMethod;
import org.springframework.shell.standard.ShellOption;

import java.io.StringReader;
import java.util.*;
import edu.stanford.nlp.tagger.maxent.MaxentTagger;
import edu.stanford.nlp.ling.HasWord;
import java.io.FileWriter;
import java.io.IOException;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;



import java.io.File;
import java.io.IOException;

@ShellComponent
public class CleanCommand {

    @ShellMethod("Picks the path of the excel")
    public String clean(@ShellOption({"-F", "--filename"}) String path) throws IOException, InvalidFormatException {
        final String SAMPLE_XLSX_FILE_PATH = path;

        //Creating a Workbook from an excel file
        Workbook workbook = WorkbookFactory.create(new File(SAMPLE_XLSX_FILE_PATH));
        // Retrieving the number of sheets in the Workbook
        //return String.format("Workbook has " + workbook.getNumberOfSheets() + " Sheets : ");
        Sheet sheet = workbook.getSheetAt(0);
        Row row = sheet.getRow(0); //assume first row is column header row

        String columnWanted = "body";
        Integer columnNo = null;
        //output all not null values to the list
        List<Cell> cells = new ArrayList<Cell>();
        Cell c = null;

        //get first cell
        Row row1 = sheet.getRow(0);
        //Cell cell = row.getCell(0);
        for (Cell cell : row1){
            if(cell.getStringCellValue().equals(columnWanted)){
                columnNo = cell.getColumnIndex();
            }
        }
        if (columnNo!=null){
            for(Row row2: sheet){
                c = row2.getCell(columnNo);
                if(c == null|| c.getCellType() == Cell.CELL_TYPE_BLANK){
                    //Nothing in the cell in this row, skip it
                }else{
                    cells.add(c);
                    //System.out.println(cells);
                }
            }
        }else{
            System.out.println("could not find column " + columnWanted);
        }


        // Create Standford tagger
        MaxentTagger tagger = new MaxentTagger("models/german-fast.tagger");
        double threshold=0.9;
        List<TaggedWord> tSentence = null;
        //JSONObject tagDetails = new JSONObject();
        JSONArray cleanedarray = new JSONArray();



        for (int i = 0; i < cells.size(); i++) {
            double prob = 0.0;
            int sum = 0;

            String email = cells.get(i).toString();
            //List<String> sentences = Arrays.asList(email.strip().split(String.valueOf('\n')));
            List<List<HasWord>> sentences = MaxentTagger.tokenizeText(new StringReader(email));
            JSONObject emailDetails = new JSONObject();

            for (List<HasWord> sentence : sentences) {
                tSentence = tagger.tagSentence(sentence);
                prob = prob_block(tSentence, sentence);

                //System.out.println(SentenceUtils.listToString(tSentence));
                if (prob < threshold) {
                    System.out.println("CleanedText");

                    if (emailDetails.get("cleaned") != null) {
                        emailDetails.put("cleaned", emailDetails.get("cleaned").toString() + sentence);
                    } else {
                        emailDetails.put("cleaned", sentence);
                    }

                    emailDetails.put("uncleaned", email);
                    cleanedarray.add(emailDetails);
                }
            }

        }
        //Write into JSON
        try (FileWriter file = new FileWriter("email.json")){
            file.write(String.valueOf(cleanedarray));
            file.flush();
        } catch (IOException e) {
            e.printStackTrace();
        }

        return String.format("Done");

    }
    public double prob_block(List<TaggedWord> tSentence, List<HasWord> sentence) {
        /**
         * Calculates probablity based on occurence of "VERB" tag
         *
         * @param  sum   counter for occurence of VERB
         * @return       (Double) probablity
         */
        int sum = 0;
        double prob = 0.0;

        for (TaggedWord tw : tSentence) {
            if (!tw.tag().equalsIgnoreCase("VVFIN") && !tw.tag().equalsIgnoreCase("VVPP")
                    && !tw.tag().equalsIgnoreCase("VVIMP")) {
                sum = sum + 1;
            }
        }

        prob = sum / (tSentence.size());
        return prob;
    }
 }

