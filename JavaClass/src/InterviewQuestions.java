
public class InterviewQuestions {
	
	static int a;
	
	public static void show() {
		System.out.println("Show Static Method");
	}
	
	static {
		a = 20;
		System.out.println("Value of a : "+a);
		show();
	}

	public static void main(String[] args) {

		System.out.println("Main Method");
	}

}
