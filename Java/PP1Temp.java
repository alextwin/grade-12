// The "PP1Temp" class.
public class PP1Temp
{
    public static void main (String[] args)
    {
	int fahrenheit, celsius;

	System.out.print ("Enter a temperature in celsius: ");
	celsius = ReadLib.readInt ();

	fahrenheit = (int)(9.0 / 5 * celsius + 32);

	System.out.print ("Temperature in fahrenheit: " + fahrenheit);
    } // main method
} // PP1Temp class
