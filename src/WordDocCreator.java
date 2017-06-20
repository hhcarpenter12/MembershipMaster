import java.util.ArrayList;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;



import java.io.FileOutputStream;
import java.io.IOException;

public class WordDocCreator { 
	//private tokensFinder finder = new tokensFinder();
	//private ArrayList <String> tokens = finder.getTokens();
	ArrayList <String> labels = new ArrayList <String>();







	public void createAndShowGUI(QuestionDoc ans) throws IOException {


		JTextArea area = new JTextArea(9, 45);
		area.setLineWrap(true);
		area.setWrapStyleWord(true);





		for (String line: ans.getText()) {
			area.setText(line);
		}


		area.setFont(area.getFont().deriveFont(14.0f));


		try { 
			FileOutputStream out = new FileOutputStream(chooseFile());
			XWPFDocument document = new XWPFDocument();

			for (String line: ans.getText()) {
				XWPFParagraph paragraph = document.createParagraph();
				XWPFRun run = paragraph.createRun();
				run.setText(line);
				run.setFontSize(14);

				
				if (line.startsWith("Red Zone")) {
					run.setColor("ff0000");
					run.setBold(true);
				}
				
				if (line.startsWith("Yellow Zone")) {
					run.setColor("cccc00");
					run.setBold(true);
				}
				
				if (line .startsWith("Green Zone")) {
					run.setColor("006600");	
					run.setBold(true);
				}
				
				
				
				if (line.startsWith("Chapter Meetings:")) {
					run.setBold(true);
					Pattern p = Pattern.compile("(\\d+(?:\\.\\d+))");
					Matcher m = p.matcher(line);
					double chapNum = 0;
					if (m.find()) {
						chapNum = Double.parseDouble(m.group(1));
					}
					if (chapNum >= ans.getMaxChapters() || chapNum == ans.getMaxChapters()-1) {
						run.setColor("006600");
						run.setBold(true);
					}
					else if (chapNum == ans.getMaxChapters()-2)
					{
						run.setColor("cccc00");
						run.setBold(true);

					}
					else {
						run.setColor("ff0000");
						run.setBold(true);
					}	
				}
				if (line.startsWith("Service Hours:")) {
					run.setBold(true);
					Pattern p = Pattern.compile("(\\d+(?:\\.\\d+))");
					Matcher m = p.matcher(line);
					double serviceNum = 0;
					if (m.find()) {
						serviceNum = Double.parseDouble(m.group(1));
					}

					System.out.println(serviceNum);
					int curr = ans.getCurrWeek();
					int sHours = ans.getHoursService();
					int sLength = ans.getSemLength();


					double goodNumber = ((curr * sHours) / sLength);
					double urgency = curr / sLength;
					int u = 0;
					
					if (urgency < 0.5) {
						u = 1;
					}
					else if (urgency >= 0.5 && urgency < 0.75) {
						u = 2;
					}
					else {
						u = 3;
					}

					if (serviceNum >= (goodNumber)) {
						run.setColor("006600");
						run.setBold(true);
					}
					else if (serviceNum < (goodNumber) && serviceNum >= (goodNumber/2 + u)) {
						run.setColor("cccc00");
						run.setBold(true);
					}
					else {
						run.setColor("ff0000");
						run.setBold(true);
					}
				}

				if (line.startsWith("Fellowship Hours:")) {
					run.setBold(true);
					Pattern p = Pattern.compile("(\\d+(?:\\.\\d+))");
					Matcher m = p.matcher(line);
					double fellowshipNum = 0;
					if (m.find()) {
						fellowshipNum = Double.parseDouble(m.group(1));
					}

					System.out.println(fellowshipNum);
					int curr = ans.getCurrWeek();
					int fHours = ans.getHoursFellowship();
					int sLength = ans.getSemLength();


					double goodNumber = ((curr * fHours) / sLength);
					double urgency = curr / sLength;
					double u = 0;
					
					if (urgency < 0.5) {
						u = 0.66;
					}
					else if (urgency >= 0.5 && urgency < 0.75) {
						u = 1.32;
					}
					else {
						u = 1.98;
					}
					

					if (fellowshipNum >= (goodNumber)) {
						run.setColor("006600");
						run.setBold(true);
					}
					else if (fellowshipNum < (goodNumber) && fellowshipNum >= (goodNumber/2 + u)) {
						run.setColor("cccc00");
						run.setBold(true);
					}
					else {
						run.setColor("ff0000");
						run.setBold(true);
					}
				}

				for (int i = 0; i < ans.getMandEvents().size(); i++) {
					if (line.startsWith(ans.getMandEvents().get(i))) {
						run.setBold(true);
						Pattern p = Pattern.compile("(\\d+(?:\\.\\d+))");
						Matcher m = p.matcher(line);
						double mandNum = 0;
						if (m.find()) {
							mandNum = Double.parseDouble(m.group(1));
						}
						if (mandNum >= 1.0) {
							run.setColor("006600");
							run.setBold(true);
						}
						else {
							run.setColor("ff0000");
							run.setBold(true);
						}	

					}


				}





				if (line.equals(ans.getFirstName())) {
					run.setBold(true);
				}







				paragraph.setWordWrapped(true);
				run.setFontSize(12);
			}


			document.write(out);
			document.close();

		}
		catch (Exception e) {
			System.out.println("Failed");
		}



	}

	public static String chooseFile() {
		JFileChooser chooser = new JFileChooser();
		chooser.setApproveButtonText("Select");
		chooser.setDialogTitle("Please Select an Ouput File");
		FileNameExtensionFilter filter = new FileNameExtensionFilter(
				"Word Document", "docx");
		chooser.setFileFilter(filter);
		int returnVal = chooser.showOpenDialog(null);
		if(returnVal == JFileChooser.APPROVE_OPTION) {
			String FILE_NAME = chooser.getSelectedFile().getAbsolutePath(); 
			return FILE_NAME;
		}
		return "";
	}














}