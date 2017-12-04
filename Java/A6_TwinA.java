//Name: Alex Twin
//Date: April 24, 2017
//Purpose: To simulate a game played between the computer and a single player
//The computer selects a scret word and the player attempts to guess it

// The "A6_TwinA" class.
public class A6_TwinA
{
    public static void main (String[] args)
    {
	String[] wordBank = {"JASON", "COMPUTER", "BOX", "FOUR", "APPLE", "JUICE", "DICE", "DATA", "BASKETBALL", "TEN"};

	int numGuess; //Counter
	int randomWord; //Random number to choose a word from word bank
	int k;
	String guess; //For guessed letter or word
	String replay; //To see if user wants to play again

	while (true)
	{
	    //Set the counter to 0 after every new game
	    numGuess = 0;
	    //Pick a random word
	    randomWord = (int) (Math.random () * 10);
	    //StringBuffer hidden is assigned random word
	    StringBuffer hidden = new StringBuffer (wordBank [randomWord]);

	    System.out.println ("Word Guessing Game:");
	    for (k = 0 ; k < wordBank [randomWord].length () ; k++)
	    {
		//Change letters into dashes
		hidden.setCharAt (k, '-');
	    }

	    System.out.println (hidden);
	    
	    //Keep looping until player wins or loses
	    while (true)
	    {
		System.out.print ("Enter a letter ($ for the entire word): ");
		//Read input
		guess = ReadLib.readString ();
		//Increment counter
		numGuess++;

		if (guess.equals ("$"))
		{
		    System.out.println ();
		    System.out.print ("What is your guess? ");
		    //Reassign guess the player's guessed word
		    guess = ReadLib.readString ();
		    
		    //See if the guess is correct
		    if (wordBank [randomWord].equalsIgnoreCase (guess))
		    {
			System.out.print ("You won!  ");
		    }
		    else
		    {
			System.out.print ("Sorry, you lost!  ");
		    }
		    
		    //Break out of loop
		    break;
		}
		else
		{
		    //Check to see if any letters match
		    for (k = 0 ; k < wordBank [randomWord].length () ; k++)
		    {
			if (guess.equalsIgnoreCase (Character.toString (wordBank [randomWord].charAt (k))))
			{
			    //Reveal letter by replacing the dash with the guessed character
			    hidden.setCharAt (k, wordBank [randomWord].charAt (k));
			}

		    }
		    
		    //If the player has guessed all letters of the hidden word
		    //The player wins
		    if (wordBank [randomWord].equals (hidden.toString ()))
		    {
			System.out.println ();
			System.out.print ("You won!  ");
			//Break out of loop
			break;
		    }
		    
		    //Display hidden word
		    System.out.println (hidden);

		}
	    }
	    
	    //Output secret word and number of guesses
	    System.out.println ("Secret word was " + wordBank [randomWord]);
	    System.out.println ("Total number of guesses: " + numGuess);
	    
	    //Ask user if they want to play again
	    while (true)
	    {
		System.out.println ();
		System.out.println ("Would you like to play again? (Y/N)");
		//Read input
		replay = ReadLib.readString ();

		if (replay.equalsIgnoreCase ("N"))
		{
		    //This ends execution the program
		    return;
		}
		else if (replay.equalsIgnoreCase ("Y"))
		{
		    //Breaks this loop to replay again
		    break;
		}
		else
		{
		    //If invalid command, keep on asking user
		    System.out.println ("Invalid command!");
		}
	    }

	    System.out.println ();

	}
    } // main method
} // A6_TwinA class
