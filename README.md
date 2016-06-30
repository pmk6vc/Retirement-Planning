# Retirement-Planning

As mentioned above, the macro I've supplied is meant to help people get a rough projection of their savings upon retirement and estimate how much monthly income they can afford for the rest of their expected life. Inspired by my desire to help my dad with his retirement planning. Also, I was bored one afternoon. I want to emphasize that this is just a rough outline based on potentially volatile parameters, and that I excluded several important variables from this simple projection (i.e. taxes, mortgages, monthly living expenses, market fluctuations, etc.). This isn't a complete and personalized financial plan by a long shot, so don't treat it like one. 

This program offers two key features:

-Estimate both nominal savings and real purchasing power upon retirement. The key parameters here are the initial savings, current annual income, expected income growth, % of income retained, expected investment return on savings, expected inflation, and years to retirement. -Estimate how much money (in nominal terms) can be withdrawn to leave behind a certain inheritance (in real terms).** They key parameters here are, on top of the ones above, expected lifespan after retirement and desired inheritance to leave behind.

# Instructions for the common user

I'm assuming you have Excel. If you do, great! Simply download the example Excel workbook and change the parameters in the top left to reflect your personal situation, and then click the button. The Summary section should provide all of the information you need. 

Potential prerequisites (read: sources of error) to be aware of:
-You must enable macros (while this is potentially dangerous, I have neither the inclination nor the ability to do any harm). 
-You must have the Solver add-in installed
-You must create the Solver reference so that VBA knows my code is using Solver (Googling "VBA Solver reference" should help)
-Don't mess with the formulas in Cells B17-B18
-If Solver cannot find a solution to leave behind your desired inheritance, then it's possible that a solution does not exist. However, sometimes Solver won't check values that are far from the initial value. So, try fidgeting with the "Proposed nominal monthly withdrawal" figure to number that approach your desired inheritance, and click the button again.
-The workbook must be in automatic calculation mode (this shouldn't be a problem and my macro does this automatically anyway, but sometimes Solver does weird things if it can't find a solution in my experience)

# Information for the developer

If you're looking to examine or modify my code, go for it. However, there are some "bad practices" that I should explain first. I wrote this macro with no intention of scaling it or optimizing it. You'll notice that I haven't declared any variable data type and that I freely use functions and subroutines interchangeably. Bad practice, yes, but I'm not proficient in VBA and for such a simple program, I didn't feel the need to polish the end product. The use case that I wanted to solve is indeed solved by this program, so I'm happy, but beware if you want to take this program and use it as part of a larger application. 

**You may be wondering why the monthly income quote is in nominal terms, while the inheritance constraint is in real terms. I made this choice because parents generally have a better idea of how much purchasing power they want to leave behind to their children (AKA how much money in today's terms) than how much they'll want month over month several years from now. 
