# Microsoft-Excel-Conditional-Formatting-nth-Row-Formula
## Summary

This is a how-to on using the Conditional Formatting rules inside Microsoft Excel to format rows within a spreadsheet by every nth row in a sequence. [My original post can be found here.](https://community.spiceworks.com/how_to/149441-microsoft-excel-conditional-formatting-nth-row-formula)
### Example of the Result:
![enter image description here](https://static.spiceworks.com/images/how_to_steps/0012/3279/ff0fe16f651ee582f3aab97e165465030336113ede043ddfcd79b47d4cdba9e5_Capture4.png)

## Intro
The general form of the formulas are "=MOD(ROW(),n)=0" [ending] and "=1" [beginning]

General form is as follows:

    =MOD(ROW(),n)=0
    =MOD(ROW(),n)=n-1
    =MOD(ROW(),n)=n-2
    =MOD(ROW(),n)=n-3
    ...
    Till n-x = 2
    ...
    =MOD(ROW(),n)=2
    =1

Our example will highlight a sequence of 4 rows for our sample data. The color sequence will be Orange, Blue, Yellow, and Red; of course yours can be anything or any format style you desire.
Some background of the modulus operator is very handy.

## Steps

1. **Locate the Conditional Formatting Settings**
	
    In Office 365 and Office 2016, Conditional Formatting is on the "Home" tab. Click button for a drop down of specific settings including helpful pre-made rules, colors, and icons.
    
    We are going to select "Manage Rules" 
    Then select "Show formatting rules for:" drop down as "This worksheet" to see if any rules already exist.
    
2. **Add a Rule to Manage Rules**
  
     Select the "New Rule..." button
    
    There are several options for the basis of varying rule types. We are going to select the one at the bottom: "Use a Formula to determine which cells to format".
    
    Here you can adjust the format of the rows by clicking "Format..".

3. **Adjust the "Applies to" field in the Manage Rules Window**
	![enter image description here](https://static.spiceworks.com/images/how_to_steps/0012/3278/f9c8415c97ed09cca266c1b5705ef3edb6c10543b9ce10cee936b25151a7b569_Capture3.png)
    After a rule is added or before you open the Rules Manger window, you will need to adjust the data these rules will apply to.
    
    To adjust before you start, drag your mouse selection so that bounding box covers the data you want this applied to.
    
    To adjust after you start a rule, when you select New Rule and click "OK" in the Rules Manager window, you'll need to select the up arrow next to the text box under the heading "Applies to" then drag the bounding box over the data. Or, simply type in a formula to include your data.

4. **First Rule formula is "=1"**
	
    First rule is "=1" as this is a layered approach to highlighting the rows. This is the base pain so to speak. We are also highlighting this orange. 
    This will pain the whole of the selection area. The colors later will be applied on top of this.

5. **Second Rule formula is "=MOD(ROW(),4)=2"**
	We are using the color blue for this rule.
6. **Third Rule formula is "=MOD(ROW(),4)=3"**
	We are using the color yellow for this rule.
7. **Fourth Rule formula is "=MOD(ROW(),4)=0"**
	We are using the color red for this rule.
	Notice how we are setting this equal to 0 (zero)
8. **Verify the Rules are in Order**
	The Rules must descend from most nth, in our case 4th row rule, at 	the top descending downwards.

	Order is as follows for this example: 
	Color RED =MOD(ROW(),4)=0 
	Color Yellow =MOD(ROW(),4)=3 
	Color Blue =MOD(ROW(),4)=2 
	Color Orange = 1

	Click Apply to apply the rules.
9. **End Result**
	![enter image description here](https://static.spiceworks.com/images/how_to_steps/0012/3279/ff0fe16f651ee582f3aab97e165465030336113ede043ddfcd79b47d4cdba9e5_Capture4.png) 
	You can verify the color sequence is now descending and repeating for our data set.

    Orange
	  Blue
	  Yellow
	  Red

## Summary
 The hard part here is understanding the expansion of the formula. In our example of 4 rows and sequences, notice where the pieces of the puzzle are located. 
4th row:
=MOD(ROW(),4)=0
4 row mod 4 = 0
Each time the 4th row is addressed 4 modulus 4 will equal 0. We want a sequence to repeat every 4th row. This is our starting point. When 4 mod 4 is equal to 0 when know we are on the fourth row. 
row 4 mod 4 = 0
row 8 mod 4 = 0
row 12 mod 4 = 0

3rd row:
=MOD(ROW(),4)=3
Now we want every 3rd row in our 4 row sequence. 
row 3 mod 4 = 3
row 7 mod 4 = 3
row 11 mod 4 = 3

2nd row:
=MOD(ROW(),4)=2
Like above, every 2nd row in our 4 row sequence. 
row 2 mod 4 = 2
row 4 mod 4 = 0
row 6 mod 4 = 2
row 8 mod 4 = 0
row 10 mod 4 = 2
row 12 mod 4 = 0

1st row: 
As long as it's not zero, it is "true" and the format will apply to the area selected. 
=1

Notice how sometimes we have an overlap? This is ok because the way Excel evaluates these rules. It creates a layered effect . I don't like how it's "bottom up" , but visually it makes sense if you view the Manage Rules window as point you have put down already. Top rules will apply over bottom rules.

The order I input them. New rules will be added "on top" of the prior rules. 
[first down]
"=1", row orange blankets everything.
"=MOD(ROW(),4)=2", row blue blankets all even rows.
"=MOD(ROW(),4)=3", row yellow takes over only the times, in the 4 set sequence, the 3rd row of that set is used. 
"=MOD(ROW(),4)=0", row red takes every 4th row. 
[last down]

Hopefully this will help anyone else if a user asks for something similar. If there's another way to do this, I'd like to hear about it!

Otherwise hopefully this "brute force method" will help anyone else stuck on how to use the Conditional Formatting, because I couldn't find this type of idea somewhere else.
