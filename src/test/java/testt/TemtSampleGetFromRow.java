package testt;

import java.io.IOException;

public class TemtSampleGetFromRow {

	public static void main(String[] args) throws IOException {
		
		DataFromRow d = new DataFromRow();
		
		System.out.println(d.getData("Add Profile"));
		System.out.println();
		System.out.println(d.getData("Add Profile").get(2));
	}
}
