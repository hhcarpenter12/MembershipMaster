import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import javax.swing.JFileChooser;
import javax.swing.filechooser.FileNameExtensionFilter;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;




public class Excel {

	private static ArrayList <String> output = new ArrayList <String>();
	private static double goodService = 0;
	private static double goodFellowship = 0;

	private static double uS = 0;
	private static double uF = 0;
	
	
	public static String chooseFile() {
		JFileChooser chooser = new JFileChooser();
		chooser.setDialogTitle("Please Select Membership Progress Report.");
		FileNameExtensionFilter filter = new FileNameExtensionFilter(
				"Excel Spreadsheet - APO Hours", "xls", "xlsx");
		chooser.setFileFilter(filter);
		int returnVal = chooser.showOpenDialog(null);
		if(returnVal == JFileChooser.APPROVE_OPTION) {
			String FILE_NAME = chooser.getSelectedFile().getAbsolutePath(); 
			return FILE_NAME;
		}
		return "";
	}



	public ArrayList <String> hoursUpdate(int currWeek, int minService, int minFellowship, int semLength, int maxChapters, String firstName, ArrayList <String> mandatoryEvents, String status) {
		try {
			FileInputStream excelFile = new FileInputStream(new File(chooseFile()));
			Workbook workbook = new XSSFWorkbook(excelFile);
			Sheet datatypeSheet = workbook.getSheetAt(0);
			Iterator<Row> iterator = datatypeSheet.iterator();
			int headerCount = 0;
		
			
			while (iterator.hasNext()) {
				
				Row currentRow = iterator.next();
				Iterator<Cell> cellIterator = currentRow.iterator();
				int count = 0;
				while (cellIterator.hasNext()) {

					Cell currentCell = cellIterator.next();
					if (headerCount != 0) {
						if (count == 0) {
							output.add("Hi " + currentCell.getStringCellValue() + ",");
							output.add(" ");
							output.add("Here is your week " + currWeek + " update.");
							output.add("Current Status: " + status);
							output.add(" ");
							output.add("Recommended Hours:");
							double urgency = currWeek/semLength;
							goodService = ((currWeek * minService) / semLength);
							goodFellowship = ((currWeek * minFellowship) / semLength);
							
							if (urgency < 0.5) {
								uS = 1;
								uF = 0.66;
							}
							else if (urgency >= 0.5 && urgency < 0.75) {
								uS = 2;
								uF = 1.32;
							}
							else {
								uS = 3;
								uF = 1.98;
							}
							
							
							
							
							
							
							output.add("Green Zone (Service): " + (goodService) + " hours and above");
							output.add("Yellow Zone (Service): < " + (goodService) + " hours and >= " + (goodService/2 + uS) + " hours.");
							output.add("Red Zone (Service): < " + (goodService/2 + uS) + " hours.");
							output.add("Green Zone (Fellowship): " + (goodFellowship) + " hours and above");
							output.add("Yellow Zone (Fellowship): < " + (goodFellowship) + " hours and >= " + (goodFellowship/2 + uF) + " hours.");
							output.add("Red Zone (Fellowship): < " + (goodFellowship/2 + uF) + " hours.");
							
							
							
							output.add(" ");
						}
						if (count == 2) {
							double chapterMeetings = currentCell.getNumericCellValue();
							checkChapterMeetings(maxChapters, chapterMeetings);
							output.add(" ");
						}
						if (count == 3) {
							double serviceHours = currentCell.getNumericCellValue();
							checkService(currWeek, minService, serviceHours, semLength);
							output.add(" ");

						}
						if (count == 4) {
							double fellowshipHours = currentCell.getNumericCellValue();
							checkFellowship(currWeek, fellowshipHours, minFellowship, semLength);
							output.add(" ");
						}


						if (count >= 5 && count-5 < mandatoryEvents.size()) {
							double checkMand = currentCell.getNumericCellValue();
							checkMandEvent(checkMand, mandatoryEvents.get(count-5));
							output.add(" ");
						}
						count++;
					}


				}
				headerCount++;
				output.add("If you have any questions or concerns, feel free to let me know.");
				output.add(" ");
				output.add("Thanks and LFS,");
				output.add(" ");
				output.add(firstName);
				output.add(" ");
				output.add("P.S: LOG YOUR HOURS!!!");


				output.add("-----------------------------------------------------------------------------------------");
				output.add(" ");

			}

			workbook.close();

		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		return output;
	}



	public static void checkMandEvent(double mandEvent, String eventName) {
		output.add(eventName + " " + mandEvent + "/" + 1);
		if (mandEvent >= 1.0) {
			output.add("Good job!");
		}
		else {
			output.add("-Don't forget to complete and log your " + eventName + " hour before the break!");
			output.add("-If you didn't attend, don't forget to ask the event organizer how to make it up!");
			output.add("-Please note that you can't miss more than 3 mandatory events without make-up to be in good standing.");
			output.add("-Please note that chapter meetings also count as mandatory events.");
		}


	}





	public static void checkChapterMeetings(int mChapters, double chapterMeetings) {

		double maxChapters = (double) mChapters;
		output.add("Chapter Meetings:  " + chapterMeetings + "/" + maxChapters);


		if (chapterMeetings >= (double) mChapters) {
			output.add("Good job! Keep it up!");
		}

		else if (chapterMeetings == (double) mChapters-1.0) {
			output.add("-There is no make-up required for one missed chapter.");
			output.add("-However, missing one chapter counts as 1 missed mandatory event. Missing more than 3 mandatory events will lead to probation.");
			output.add("-For three missed chapters, you need to do an extra hour of fellowship or service, or attend an exec meeting.");
		}

		else if (chapterMeetings == (double) mChapters-2.0) {
			output.add("-There is no make-up required for two missed chapters.");
			output.add("-However, missing two chapters counts as 2 missed mandatory events. Missing more than 3 mandatory events will lead to probation.");
			output.add("-For three missed chapters, you need to do an extra hour of fellowship or service, or attend an exec meeting.");
		}
		else if (chapterMeetings == (double) mChapters-3.0) {
			output.add("-You need to make up 3 unexcused absences by attending an exec meeting");
			output.add("-Alternatively, you can do an extra hour of service or fellowship.");
			output.add("-Make sure all your other mandatory events (e.g. Pledge service event, Weenie Roast, Book co-op) are completed to stay in good standing.");
			output.add("-To stay in good standing, try to make all the other chapters.");
		}

		else if (chapterMeetings <= (double) mChapters-4.0) {
			output.add("-You will be on probation status next term.");
			output.add("-If you have any questions or concerns, please come to an exec meeting.");
		}

	}








	public static void checkService(int inputUpdate, int minService, double serviceHours, int semLength) {
		

		output.add("Service Hours:  " + serviceHours + "/" + minService);

		

		if (serviceHours >= (goodService)) {
			output.add("Awesome job! Just keep serving!");
		}
		else if (serviceHours < (goodService) && serviceHours >= (goodService/2 + uS)) {
			warningMessageService();
		}
		else {
			dangerMessageService();
		}
	}

	public static void checkFellowship(int inputUpdate, double fellowshipHours, int minFellowship, int semLength) {
		


		output.add("Fellowship Hours:" + fellowshipHours + "/" + minFellowship);

		if (fellowshipHours >= (goodFellowship)) {
			output.add("Awesome job! Just keep fellowshipping!");
		}
		else if (fellowshipHours < (goodFellowship) && fellowshipHours >= (goodFellowship/2 + uF)) {
			warningMessageFellowship();
		}
		else {
			dangerMessageFellowship();
		}



	}	

	public static void dangerMessageService() {

		output.add("-I am quite worried about your hours requirements. Be sure to start logging them if you want to stay in good standing!");
		output.add("-Be sure to go to APO Online to find service events that fit into your schedule.");
		output.add("-Try to attend some service events as soon as possible."); 
		output.add("-If you ever have difficulties finding service events or are unsure about what is internal/external, feel free to contact the VP(s) of Service. They are more than happy to help you out.");
		output.add("-If you have any difficulties logging your hours, feel free to contact me.");

	}

	public static void warningMessageService() {
		output.add("-You are slightly behind on your hours progress. Be sure to go to APO Online to find service events that fit into your schedule.");
		output.add("-If you ever have difficulties finding service events or are unsure about what is internal/external, feel free to contact the VP(s) of Service. They are more than happy to help you out.");
		output.add("-If you have any difficulties logging your hours, feel free to contact me.");
	}


	public static void dangerMessageFellowship() {
		output.add("-I am quite worried about your hours requirements. Be sure to start logging them if you want to stay in good standing!");
		output.add("-Be sure to go to APO Online to find fellowship events that fit into your schedule.");
		output.add("-Try to attend some fellowship events as soon as possible."); 
		output.add("-If you ever have difficulties finding fellowship events or are unsure about what is internal/external, feel free to contact the VP(s) of Fellowship. They are more than happy to help you out.");
		output.add("-If you have any difficulties logging your hours, feel free to contact me.");
	}

	public static void warningMessageFellowship() {
		output.add("-You are slightly behind on your hours progress. Be sure to go to APO Online to find fellowship events that fit into your schedule.");
		output.add("-If you ever have difficulties finding fellowship events or are unsure about what is internal/external, feel free to contact the VP(s) of Fellowship. They are more than happy to help you out.");
		output.add("-If you have any difficulties logging your hours, feel free to contact me.");
	}

}













	
