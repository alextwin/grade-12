// The "PP1MRoots" class.
public class PP1MRoots
{
    public static void main (String[] args)
    {
	double a, b, c, discriminant;
	
	System.out.print("a: ");
	a = ReadLib.readDouble();
	System.out.print("b: ");
	b = ReadLib.readDouble();
	System.out.print("c: ");
	c = ReadLib.readDouble();
	
	discriminant = b * b - 4 * a * c;
	if(discriminant >= 0){
	    System.out.print("The roots are real");
	}else{
	    System.out.print("The roots are not real");
	}
    } // main method
} // PP1MRoots class
