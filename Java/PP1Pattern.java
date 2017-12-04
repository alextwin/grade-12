// The "PP1Pattern" class.
public class PP1Pattern
{
    public static void main (String[] args)
    {
	int num, k, j;
	
	System.out.print("Enter a number: ");
	num = ReadLib.readInt()
	
	for(k = num; k >= 1; k--){
	    for(j = k; j >= 1; j--){
		System.out.print(j + " ");
	    }
	    System.out.println();
	}
    } // main method
} // PP1Pattern class
