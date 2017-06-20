import javax.swing.*;

import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;

public class QuestionDoc implements ActionListener {
	//private JTextField textField = new JTextField(10);
	JLabel result;
	String currentPattern;
	ArrayList <String> labels = new ArrayList <String>();
	private int currWeek;
	private int hoursService;
	private int hoursFellowship;
	private int semLength;
	private int maxChapters;
	private ArrayList <String> mandatoryEvents;
	private String firstName;
	private String status;

	private ArrayList <String> text;

	public QuestionDoc(ArrayList <String> t, int mChapters, int cWeek, int hService, int hFellowship, int sLength, String fName, ArrayList <String> mEvents, String s) {
		text = t;
		maxChapters = mChapters;
		currWeek = cWeek;
		hoursService = hService;
		hoursFellowship = hFellowship;
		semLength = sLength;
		firstName = fName;
		mandatoryEvents = mEvents;
		status = s;
	}

	public QuestionDoc () {

	}


	private TextField [] textFields;
	/**
	 * Create the GUI and show it.  For thread safety,
	 * this method should be invoked from the
	 * event-dispatching thread.
	 */


	public void createAndShowGUI() {

		labels.add("Current Week:");
		labels.add("Semester Length (weeks):");
		labels.add("Enter Status Group:");
		labels.add("Chapter Attendance Requirements:");
		labels.add("Service Hour Requirements:");
		labels.add("Fellowship Hour Requirements:");
		labels.add("List Mandatory Events:");
		labels.add("First Name:");


		textFields = new TextField[labels.size()];
		int numPairs = labels.size();

		//Create and populate the panel.
		JPanel p = new JPanel(new SpringLayout());
		JLabel l = new JLabel(labels.get(0), JLabel.TRAILING);
		p.add(l);
		textFields[0] = new TextField(10);
		textFields[0].addActionListener(this);


		l.setLabelFor(textFields[0]);
		p.add(textFields[0]);

		for (int i = 1; i < numPairs; i++) {
			l = new JLabel(labels.get(i), JLabel.TRAILING);
			p.add(l);
			textFields[i] = new TextField(10);
			//textFields[i].addActionListener(this);

			l.setLabelFor(textFields[i]);
			p.add(textFields[i]);
		}


		JButton ok = new JButton();
		ok.setText("Generate Updates!");
		ok.addActionListener(this);
		l = new JLabel("", JLabel.TRAILING);
		p.add(l);
		l.setLabelFor(ok);
		p.add(ok);



		//Lay out the panel.
		QuestionDocHelper.makeCompactGrid(p,
				numPairs+1, 2, //rows, cols
				18, 18,        //initX, initY
				18, 18);       //xPad, yPad

		//Create and set up the window.
		JFrame frame = new JFrame("MembershipMaster");
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

		//Set up the content pane.
		p.setOpaque(true);  //content panes must be opaque
		frame.setContentPane(p);

		//Display the window.
		frame.pack();
		frame.setVisible(true);
	}

	public static void main(String[] args) {
		//Schedule a job for the event-dispatching thread:
		//creating and showing this application's GUI.
		javax.swing.SwingUtilities.invokeLater(new Runnable() {
			public void run() {
				QuestionDoc sF = new QuestionDoc();
				sF.createAndShowGUI();
			}
		});
	}
	public int getCurrWeek() {
		return currWeek;
	}


	public int getHoursService() {
		return hoursService;
	}

	public int getHoursFellowship() {
		return hoursFellowship;
	}

	public ArrayList <String> getText() {
		return text;
	}

	public ArrayList <String> getMandEvents() {
		return mandatoryEvents;
	}

	public String getStatus() {
		return status;
	}


	public int getMaxChapters() {
		return maxChapters;
	}

	public int getSemLength() {
		return semLength;
	}

	public String getFirstName() {
		return firstName;
	}

	@Override
	public void actionPerformed(ActionEvent e) {

		currWeek = Integer.parseInt(textFields[0].getText());
		semLength = Integer.parseInt(textFields[1].getText());
		status = textFields[2].getText();
		maxChapters = Integer.parseInt(textFields[3].getText());
		hoursService = Integer.parseInt(textFields[4].getText());
		hoursFellowship = Integer.parseInt(textFields[5].getText());
		mandatoryEvents = new ArrayList<String>(Arrays.asList(textFields[6].getText().split(",")));
		firstName = textFields[7].getText();
		Excel sheet = new Excel();
		text = sheet.hoursUpdate(currWeek, hoursService, hoursFellowship, semLength, maxChapters, firstName, mandatoryEvents, status);
		WordDocCreator foo = new WordDocCreator();
		QuestionDoc sF = new QuestionDoc(text, maxChapters, currWeek, hoursService, hoursFellowship, semLength, firstName, mandatoryEvents, status);
		try {
			foo.createAndShowGUI(sF);
		} catch (IOException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}

	}






}