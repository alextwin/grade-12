//The "Die" helper class generates a random
//number between 1-6 (compile, not run)

public class Die
{
    private int value;  //local to class
    
    public Die()    //constructor
    {
	roll();
    }
    
    public void roll()
    {
	value = (int)(Math.random() * 6 + 1);
    }
    
    public int getValue()
    {
	return value;
    }
}//Die class
