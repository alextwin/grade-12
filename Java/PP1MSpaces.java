// The "PP1MSpaces" class.
public class PP1MSpaces
{
    public static void main (String[] args)
    {
	String st;
	char ch;
	int k, counter;
	
	counter = 0;
	
	System.out.print("Enter sentence: ");
	st = ReadLib.readString();
	
	for(k = 0; k <= st.length() - 1; k++){
	    ch = st.charAt(k);
	    if(ch == ' '){
		counter++;
	    }
	}
	
	System.out.print("Number of spaces: " + counter);
    } // main method
} // PP1MSpaces class
