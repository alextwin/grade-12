// The "DieTest" class.
public class DieTest
{
    public static void main (String[] args)
    {
	final int num = 8000;
	int k, sum;
	Die d1 = new Die();
	double average;
	
	sum = 0;
	
	for(k = 0; k < num; k++){
	    d1.roll();
	    sum += d1.getValue();
	}
	
	average = (double)sum / num;
	
	System.out.print("Expected value: " + average);
    } // main method
} // DieTest class
