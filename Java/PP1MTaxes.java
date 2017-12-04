// The "PP1MTaxes" class.
public class PP1MTaxes
{
    public static void main (String[] args)
    {
	final double HST = 1.13;
	double price, total;

	System.out.print("Enter price: ")
	price = ReadLib.readInt();
	
	total = price * HST;
	System.out.print(total);
    } // main method
} // PP1MTaxes class
