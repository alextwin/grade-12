// The "PP1MReverse" class.
public class PP1MReverse
{
    public static void main (String[] args)
    {
	String name, reversed;
	int k;
	
	reversed = "";
	
	System.out.print("Type name: ");
	name = ReadLib.readString();
	
	for(k = name.length() - 1; k >= 0; k--){
	    reversed = reversed + name.charAt(k);
	}
	
	System.out.print("Reversed: " + reversed);
    } // main method
} // PP1MReverse class
