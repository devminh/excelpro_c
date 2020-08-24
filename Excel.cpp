#include "Excel.h"

void Excel::readSheet()
{
	fstream fs;
	fs.open("Data.txt");

	string r;
	string sub;
	  
	while (!fs.eof())
	{
		getline(fs, r);
		if (r != "")
		{


			sub = r.substr(0, 6);

			transform(sub.begin(), sub.end(), sub.begin(), toupper);

			if (sub.substr(0, 5) == "SHEET")
			{
				add = sub;

			}
			else
			{
				ex[add].addCells(r);
				
			}

		}
	}
	fs.close();

}

void Excel::addSheet()
{
	cout << "\n How many sheets do you want to add ? ";
	int n; cin >> n;
	
	string s1;
	
	Sheets s2;

	int t = ex.size();

	for (int i = 0; i < n ; i++)
	{
		
		s1 = "SHEET" + to_string(  t + i + 1);
		
		ex[s1].CreateSheet();
		
		
	}

}

string Excel::getSheetAddress()
{
	int a;
	cout << "\n Choose sheet you want to update : ";
	cin >> a;
	
	return "SHEET"+ to_string(a) ;
}

void Excel::update()
{
	
   ex[this->getSheetAddress()].input();
		
}



void Excel::print()
{
	HANDLE  hConsole = GetStdHandle(STD_OUTPUT_HANDLE);
	SetConsoleTextAttribute(hConsole, 175);
	cout << "\n -------------------EXCEL PROGRAM-------------------------" << endl;
	SetConsoleTextAttribute(hConsole, 15);

	map<string, Sheets> exsub= ex;    //this is solved
	map<string, Sheets> :: iterator i;

	for (i = exsub.begin(); i != exsub.end(); i++)
	{
		   i->second.countInSheet();
		    i->second.fillScan(exsub);
			i->second.scanSheet(exsub);
			

			i->second.PrintSheet();
			cout << endl;
		

	}
}


